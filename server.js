const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── State ────────────────────────────────────────────────────────────────
let state = {
  teorico:      {},
  fisico:       {},
  asignaciones: {},
  historial:    [],
  costos:       {},
  date:         new Date().toDateString(),
  version:      0
};

// Active users: { name: lastSeenTimestamp }
let activeUsers = {};
const ACTIVE_TIMEOUT = 15000; // 15 seconds

function getActiveUsers() {
  const now = Date.now();
  // Remove users not seen in last 15s
  Object.keys(activeUsers).forEach(name => {
    if(now - activeUsers[name] > ACTIVE_TIMEOUT) delete activeUsers[name];
  });
  return Object.keys(activeUsers);
}

function resetIfNewDay() {
  const today = new Date().toDateString();
  if(state.date !== today) {
    state = { teorico:{}, fisico:{}, asignaciones:{}, historial:[],
              costos:{}, date:today, version:0 };
  }
}

function addHistorial(usuario, accion, detalle) {
  state.historial.push({
    hora: new Date().toLocaleTimeString('es'), usuario, accion, detalle: detalle||''
  });
  if(state.historial.length > 200) state.historial.shift();
  state.version++;
}

function publicState() {
  return { teorico:state.teorico, fisico:state.fisico,
           asignaciones:state.asignaciones, historial:state.historial.slice(-50),
           date:state.date, version:state.version,
           activeUsers: getActiveUsers() };
}

// ── Polling API ──────────────────────────────────────────────────────────
// Client polls this every 3 seconds with its current version
// Heartbeat — client sends name, server tracks as active
app.post('/api/heartbeat', (req, res) => {
  const { name } = req.body;
  if(name) activeUsers[name] = Date.now();
  res.json({ active: getActiveUsers() });
});

app.get('/api/state', (req, res) => {
  resetIfNewDay();
  const s = publicState();
  s.activeUsers = getActiveUsers();
  res.json(s);
});

// Save conteo
app.post('/api/conteo', (req, res) => {
  resetIfNewDay();
  const { cont, data, usuario } = req.body;
  if(!cont || !data) return res.status(400).json({ ok:false });
  state.fisico[cont] = data;
  addHistorial(usuario||'—', 'Conteo guardado', cont);
  res.json({ ok:true, version:state.version });
});

// Asignaciones
app.post('/api/asign', (req, res) => {
  resetIfNewDay();
  const { cont, name, action, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.asignaciones[cont]) state.asignaciones[cont] = [];
  if(action === 'add') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(usuario||'—', 'Asignación', cont + ' → ' + name);
  } else if(action === 'remove') {
    state.asignaciones[cont] = state.asignaciones[cont].filter(n => n !== name);
    addHistorial(usuario||'—', 'Asignación removida', cont + ' ← ' + name);
  } else if(action === 'self') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(name, 'Auto-asignación', cont);
  }
  state.version++;
  res.json({ ok:true, version:state.version });
});

// Upload teorico
const upload = multer({ storage: multer.memoryStorage() });
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    resetIfNewDay();
    const wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const usuario = req.body.usuario || '—';
    let loaded = [];
    wb.SheetNames.forEach(sheetName => {
      const nl   = sheetName.toLowerCase();
      const type = nl.includes('traslado') ? 'Traslados' : nl.includes('embarque') ? 'Embarques' : null;
      if(!type) return;
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, type);
      if(count > 0) { loaded.push(sheetName+' ('+count+')'); addHistorial(usuario,'Teórico cargado',type+' — '+count+' contenedores'); }
    });
    if(loaded.length === 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, 'General');
      if(count > 0) { loaded.push(wb.SheetNames[0]+'('+count+')'); addHistorial(usuario,'Teórico cargado',loaded.join()); }
    }
    state.version++;
    res.json({ ok:true, loaded });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// Upload costos
app.post('/api/costos', upload.single('file'), (req, res) => {
  try {
    const wb   = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const sn   = wb.SheetNames.find(n=>n.toLowerCase().includes('existencia')||n.toLowerCase().includes('sap'))||wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header:1, defval:'', raw:false });
    const hdr  = rows[0].map(h=>norm(String(h)));
    const cSku = findCol(hdr,['articulo','artículo','sku','codigo']);
    const cCost= findCol(hdr,['costo promedio','costo']);
    if(cSku<0||cCost<0) return res.status(400).json({ ok:false, error:'Columnas no encontradas' });
    const raw = {};
    rows.slice(1).forEach(row => {
      const sku=String(row[cSku]||'').trim(), cost=parseFloat(String(row[cCost]).replace(',','.'));
      if(sku&&!isNaN(cost)&&cost>0){if(!raw[sku])raw[sku]={sum:0,n:0};raw[sku].sum+=cost;raw[sku].n++;}
    });
    state.costos={};let cnt=0;
    Object.keys(raw).forEach(sku=>{state.costos[sku]=raw[sku].sum/raw[sku].n;cnt++;});
    res.json({ ok:true, count:cnt });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

app.get('/api/costos', (req,res) => res.json(state.costos));

// ── Helpers ──────────────────────────────────────────────────────────────
function norm(s){ return String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function findCol(hdr,terms){
  for(const t of terms){const i=hdr.findIndex(h=>h===norm(t));if(i>=0)return i;}
  for(const t of terms){const i=hdr.findIndex(h=>h.includes(norm(t)));if(i>=0)return i;}
  return -1;
}
function mergeSheet(rows,type){
  if(!rows||rows.length<2)return 0;
  const hdr=rows[0].map(h=>norm(h));
  const colCont=findCol(hdr,['numero','número']);
  const colSku=findCol(hdr,['sku']);
  const colQty=findCol(hdr,['cant.','cant','cantidad']);
  let colDesc=-1;
  for(let i=colSku+1;i<hdr.length;i++){if(hdr[i].includes('nombre')||hdr[i].includes('descripcion')){colDesc=i;break;}}
  if(colDesc<0)colDesc=findCol(hdr,['nombre','descripcion']);
  if(colCont<0||colSku<0||colDesc<0||colQty<0)return 0;
  const g=(row,i)=>i>=0?row[i]:'';
  const ex={
    cFecha:findCol(hdr,['fecha']),cCodProv:findCol(hdr,['proveedor','código de proveedor']),
    cLineas:findCol(hdr,['lineas','líneas']),cStatus:findCol(hdr,['status']),
    cOC:findCol(hdr,['orden de compra','# orden']),cIngr:findCol(hdr,['ingresado']),
    cColoc:findCol(hdr,['colocado']),cFalt:findCol(hdr,['faltantes']),
    cSobr:findCol(hdr,['sobrantes']),cDan:findCol(hdr,['dañado','danado']),
    cOrigen:findCol(hdr,['origen']),cIngreso:findCol(hdr,['# ingreso','ingreso']),
    cDocSap:findCol(hdr,['doc. sap','doc sap']),cTipo:findCol(hdr,['tipo']),
    cUnidad:findCol(hdr,['unidad']),cUnidades:findCol(hdr,['unidades']),cDestino:findCol(hdr,['destino'])
  };
  const newConts={};
  rows.slice(1).filter(r=>r.some(c=>String(c).trim()!=='')).forEach(row=>{
    const cont=String(row[colCont]||'').trim();if(!cont)return;
    if(!newConts[cont])newConts[cont]=[];
    newConts[cont].push({sku:String(row[colSku]||'').trim(),desc:String(row[colDesc]||'').trim(),
      qty:parseFloat(String(row[colQty]).replace(',','.'))||0,
      raw:{fecha:g(row,ex.cFecha),codProv:g(row,ex.cCodProv),lineas:g(row,ex.cLineas),
        status:g(row,ex.cStatus),oc:g(row,ex.cOC),ingresado:g(row,ex.cIngr),
        colocado:g(row,ex.cColoc),faltantes:g(row,ex.cFalt),sobrantes:g(row,ex.cSobr),
        daniado:g(row,ex.cDan),origen:g(row,ex.cOrigen),ingreso:g(row,ex.cIngreso),
        docSap:g(row,ex.cDocSap),tipo:g(row,ex.cTipo),unidad:g(row,ex.cUnidad),
        unidades:g(row,ex.cUnidades),destino:g(row,ex.cDestino)}});
  });
  Object.keys(newConts).forEach(cont=>{
    state.teorico[cont]={items:newConts[cont],type};
    if(!state.fisico[cont])state.fisico[cont]=null;
  });
  return Object.keys(newConts).length;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log('Conteo app on port '+PORT));
