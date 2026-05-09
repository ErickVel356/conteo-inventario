const express = require('express');
const http    = require('http');
const WebSocket = require('ws');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');

const app    = express();
const server = http.createServer(app);
const wss    = new WebSocket.Server({ server });

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── In-memory state (resets on server restart) ──────────────────────────
let state = {
  teorico:      {},   // { contId: { items, type } }
  fisico:       {},   // { contId: [{fisico, daniado}] | null }
  asignaciones: {},   // { contId: [names] }
  historial:    [],   // [{hora, usuario, accion, detalle}]
  costos:       {},   // { sku: avgCost }
  date:         new Date().toDateString()
};

function resetIfNewDay() {
  const today = new Date().toDateString();
  if(state.date !== today) {
    state = { teorico:{}, fisico:{}, asignaciones:{}, historial:[], costos:{}, date:today };
    broadcast({ type:'reset', state: publicState() });
  }
}

function publicState() {
  return {
    teorico:      state.teorico,
    fisico:       state.fisico,
    asignaciones: state.asignaciones,
    historial:    state.historial.slice(-50),
    date:         state.date
  };
}

function addHistorial(usuario, accion, detalle) {
  state.historial.push({
    hora:    new Date().toLocaleTimeString('es'),
    usuario, accion, detalle: detalle||''
  });
  if(state.historial.length > 200) state.historial.shift();
}

// ── WebSocket broadcast ──────────────────────────────────────────────────
function broadcast(msg, exceptWs) {
  const data = JSON.stringify(msg);
  wss.clients.forEach(client => {
    if(client !== exceptWs && client.readyState === WebSocket.OPEN)
      client.send(data);
  });
}

wss.on('connection', ws => {
  resetIfNewDay();
  // Send current state to new client
  ws.send(JSON.stringify({ type:'init', state: publicState() }));

  ws.on('message', raw => {
    try {
      const msg = JSON.parse(raw);
      resetIfNewDay();

      if(msg.type === 'save_conteo') {
        state.fisico[msg.cont] = msg.data;
        addHistorial(msg.usuario, 'Conteo guardado', msg.cont);
        broadcast({ type:'conteo_saved', cont:msg.cont, data:msg.data,
          historial: state.historial.slice(-50) });
      }

      else if(msg.type === 'add_asign') {
        if(!state.asignaciones[msg.cont]) state.asignaciones[msg.cont] = [];
        if(!state.asignaciones[msg.cont].includes(msg.name))
          state.asignaciones[msg.cont].push(msg.name);
        addHistorial(msg.usuario, 'Asignación', msg.cont + ' → ' + msg.name);
        broadcast({ type:'asign_update', asignaciones:state.asignaciones,
          historial: state.historial.slice(-50) });
      }

      else if(msg.type === 'remove_asign') {
        if(state.asignaciones[msg.cont])
          state.asignaciones[msg.cont] = state.asignaciones[msg.cont].filter(n => n !== msg.name);
        addHistorial(msg.usuario, 'Asignación removida', msg.cont + ' ← ' + msg.name);
        broadcast({ type:'asign_update', asignaciones:state.asignaciones,
          historial: state.historial.slice(-50) });
      }

      else if(msg.type === 'self_assign') {
        if(!state.asignaciones[msg.cont]) state.asignaciones[msg.cont] = [];
        if(!state.asignaciones[msg.cont].includes(msg.name))
          state.asignaciones[msg.cont].push(msg.name);
        addHistorial(msg.name, 'Auto-asignación', msg.cont);
        broadcast({ type:'asign_update', asignaciones:state.asignaciones,
          historial: state.historial.slice(-50) });
      }
    } catch(e) { console.error('WS msg error:', e.message); }
  });
});

// ── REST: Upload teorico ─────────────────────────────────────────────────
const upload = multer({ storage: multer.memoryStorage() });

app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    resetIfNewDay();
    const wb   = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const usuario = req.body.usuario || '—';
    let loaded = [];

    wb.SheetNames.forEach(sheetName => {
      const nl   = sheetName.toLowerCase();
      const type = nl.includes('traslado') ? 'Traslados' : nl.includes('embarque') ? 'Embarques' : null;
      if(!type) return;

      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, type);
      if(count > 0) {
        loaded.push(sheetName + ' (' + count + ' contenedores)');
        addHistorial(usuario, 'Teórico cargado', type + ' — ' + count + ' contenedores');
      }
    });

    // If no known sheets, load first
    if(loaded.length === 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, 'General');
      if(count > 0) loaded.push(wb.SheetNames[0] + ' (' + count + ')');
      addHistorial(usuario, 'Teórico cargado', loaded.join(', '));
    }

    broadcast({ type:'teorico_loaded', state: publicState() });
    res.json({ ok:true, loaded });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── REST: Upload costos ──────────────────────────────────────────────────
app.post('/api/costos', upload.single('file'), (req, res) => {
  try {
    const wb   = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const sn   = wb.SheetNames.find(n => n.toLowerCase().includes('existencia') || n.toLowerCase().includes('sap')) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header:1, defval:'', raw:false });
    const hdr  = rows[0].map(h => norm(String(h)));
    const cSku = findCol(hdr, ['articulo','artículo','sku','codigo']);
    const cCost= findCol(hdr, ['costo promedio','costo']);
    if(cSku<0||cCost<0) return res.status(400).json({ ok:false, error:'Columnas no encontradas' });

    const raw = {};
    rows.slice(1).forEach(row => {
      const sku  = String(row[cSku]||'').trim();
      const cost = parseFloat(String(row[cCost]).replace(',','.'));
      if(sku && !isNaN(cost) && cost>0) {
        if(!raw[sku]) raw[sku]={sum:0,n:0};
        raw[sku].sum+=cost; raw[sku].n++;
      }
    });
    state.costos = {};
    let cnt = 0;
    Object.keys(raw).forEach(sku => { state.costos[sku]=raw[sku].sum/raw[sku].n; cnt++; });
    res.json({ ok:true, count:cnt });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── REST: Get costos for export ──────────────────────────────────────────
app.get('/api/costos', (req,res) => res.json(state.costos));

// ── Helper functions ─────────────────────────────────────────────────────
function norm(s){ return String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function findCol(hdr, terms) {
  for(const t of terms) { const i=hdr.findIndex(h=>h===norm(t)); if(i>=0) return i; }
  for(const t of terms) { const i=hdr.findIndex(h=>h.includes(norm(t))); if(i>=0) return i; }
  return -1;
}
function mergeSheet(rows, type) {
  if(!rows||rows.length<2) return 0;
  const hdr     = rows[0].map(h=>norm(h));
  const colCont = findCol(hdr,['numero','número']);
  const colSku  = findCol(hdr,['sku']);
  const colQty  = findCol(hdr,['cant.','cant','cantidad']);
  let colDesc=-1;
  for(let i=colSku+1;i<hdr.length;i++){if(hdr[i].includes('nombre')||hdr[i].includes('descripcion')){colDesc=i;break;}}
  if(colDesc<0) colDesc=findCol(hdr,['nombre','descripcion']);
  if(colCont<0||colSku<0||colDesc<0||colQty<0) return 0;

  const extra = {
    cFecha:   findCol(hdr,['fecha']),
    cCodProv: findCol(hdr,['proveedor','código de proveedor']),
    cLineas:  findCol(hdr,['lineas','líneas']),
    cStatus:  findCol(hdr,['status']),
    cOC:      findCol(hdr,['orden de compra','# orden']),
    cIngr:    findCol(hdr,['ingresado']),
    cColoc:   findCol(hdr,['colocado']),
    cFalt:    findCol(hdr,['faltantes']),
    cSobr:    findCol(hdr,['sobrantes']),
    cDan:     findCol(hdr,['dañado','danado']),
    cOrigen:  findCol(hdr,['origen']),
    cIngreso: findCol(hdr,['# ingreso','ingreso']),
    cDocSap:  findCol(hdr,['doc. sap','doc sap']),
    cTipo:    findCol(hdr,['tipo']),
    cUnidad:  findCol(hdr,['unidad']),
    cUnidades:findCol(hdr,['unidades']),
    cDestino: findCol(hdr,['destino']),
  };

  const newConts = {};
  rows.slice(1).filter(r=>r.some(c=>String(c).trim()!=='')).forEach(row=>{
    const cont=String(row[colCont]||'').trim(); if(!cont) return;
    if(!newConts[cont]) newConts[cont]=[];
    const g=(i)=>i>=0?row[i]:'';
    newConts[cont].push({
      sku:  String(row[colSku]||'').trim(),
      desc: String(row[colDesc]||'').trim(),
      qty:  parseFloat(String(row[colQty]).replace(',','.'))||0,
      raw:{
        fecha:g(extra.cFecha),     codProv:g(extra.cCodProv),
        lineas:g(extra.cLineas),   status:g(extra.cStatus),
        oc:g(extra.cOC),           ingresado:g(extra.cIngr),
        colocado:g(extra.cColoc),  faltantes:g(extra.cFalt),
        sobrantes:g(extra.cSobr),  daniado:g(extra.cDan),
        origen:g(extra.cOrigen),   ingreso:g(extra.cIngreso),
        docSap:g(extra.cDocSap),   tipo:g(extra.cTipo),
        unidad:g(extra.cUnidad),   unidades:g(extra.cUnidades),
        destino:g(extra.cDestino),
      }
    });
  });

  Object.keys(newConts).forEach(cont=>{
    const exists = state.fisico[cont];
    state.teorico[cont]={items:newConts[cont],type};
    if(!exists) state.fisico[cont]=null;
  });
  return Object.keys(newConts).length;
}

const PORT = process.env.PORT || 3000;
server.listen(PORT, ()=>console.log(`Conteo app running on port ${PORT}`));
