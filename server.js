const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const https   = require('https');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Supabase config (set via environment variables in Render) ─────────────
const SUPABASE_URL = process.env.SUPABASE_URL;       // https://xxx.supabase.co
const SUPABASE_KEY = process.env.SUPABASE_KEY;       // service_role key

// ── Simple Supabase REST client ───────────────────────────────────────────
function supabase(method, table, body, query) {
  return new Promise((resolve, reject) => {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return resolve(null); // No Supabase configured — use memory only
    }
    const url  = new URL(`${SUPABASE_URL}/rest/v1/${table}${query||''}`);
    const data = body ? JSON.stringify(body) : null;
    const opts = {
      hostname: url.hostname,
      path:     url.pathname + url.search,
      method,
      headers: {
        'apikey':        SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type':  'application/json',
        'Prefer':        'return=representation'
      }
    };
    if(data) opts.headers['Content-Length'] = Buffer.byteLength(data);
    const req = https.request(opts, res => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => {
        try { resolve(raw ? JSON.parse(raw) : null); }
        catch(e) { resolve(null); }
      });
    });
    req.on('error', reject);
    if(data) req.write(data);
    req.end();
  });
}

async function dbGet(key) {
  const rows = await supabase('GET', 'app_state', null, `?key=eq.${key}&select=value`);
  if(rows && rows.length > 0) {
    try { return JSON.parse(rows[0].value); } catch(e) { return null; }
  }
  return null;
}

async function dbSet(key, value) {
  const data = { key, value: JSON.stringify(value) };
  // Upsert
  await supabase('POST', 'app_state', data, '?on_conflict=key');
}

// ── In-memory state ───────────────────────────────────────────────────────
let state = {
  teorico:      {},
  fisico:       {},
  asignaciones: {},
  historial:    [],
  costos:       {},
  cdg:          {},
  date:         new Date().toDateString(),
  version:      0
};

// Field locks: { "contId:rowIdx": { user, since, expires } }
let fieldLocks = {};
const LOCK_TIMEOUT = 90000; // 90 seconds max

function cleanLocks() {
  const now = Date.now();
  Object.keys(fieldLocks).forEach(k => {
    if(fieldLocks[k].expires < now) delete fieldLocks[k];
  });
}

function getLocks() {
  cleanLocks();
  return fieldLocks;
}
let activeUsers = {};
const ACTIVE_TIMEOUT = 15000;

// ── Load state from Supabase on startup ───────────────────────────────────
async function loadState() {
  try {
    const saved = await dbGet('daily_state');
    if(saved && saved.date === new Date().toDateString()) {
      state = { ...state, ...saved };
      console.log('State restored from Supabase ✓');
    } else {
      console.log('New day or no saved state — starting fresh');
    }
  } catch(e) {
    console.log('Could not load from Supabase:', e.message);
  }
}

// ── Save state to Supabase (debounced) ───────────────────────────────────
let saveTimer = null;
function scheduleSave() {
  if(saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      await dbSet('daily_state', {
        teorico:      state.teorico,
        fisico:       state.fisico,
        asignaciones: state.asignaciones,
        historial:    state.historial.slice(-100),
        cdg:          state.cdg,
        date:         state.date,
        version:      state.version
      });
    } catch(e) { console.log('Save error:', e.message); }
  }, 1000); // save 1s after last change
}

function resetIfNewDay() {
  const today = new Date().toDateString();
  if(state.date !== today) {
    state = { teorico:{}, fisico:{}, asignaciones:{}, historial:[],
              costos:{}, cdg:{}, date:today, version:0 };
    scheduleSave();
  }
}

function addHistorial(usuario, accion, detalle) {
  state.historial.push({
    hora: new Date().toLocaleTimeString('es'), usuario, accion, detalle:detalle||''
  });
  if(state.historial.length > 200) state.historial.shift();
  state.version++;
  scheduleSave();
}

function publicState() {
  return {
    teorico:      state.teorico,
    fisico:       state.fisico,
    asignaciones: state.asignaciones,
    historial:    state.historial.slice(-50),
    cdg:          state.cdg,
    date:         state.date,
    version:      state.version,
    activeUsers:  getActiveUsers(),
    locks:        getLocks()
  };
}

function getActiveUsers() {
  const now = Date.now();
  Object.keys(activeUsers).forEach(n => { if(now - activeUsers[n] > ACTIVE_TIMEOUT) delete activeUsers[n]; });
  return Object.keys(activeUsers);
}

// ── API ───────────────────────────────────────────────────────────────────
app.post('/api/heartbeat', (req, res) => {
  const { name } = req.body;
  if(name) activeUsers[name] = Date.now();
  res.json({ active: getActiveUsers() });
});

app.get('/api/state', (req, res) => {
  resetIfNewDay();
  res.json(publicState());
});

// Full state save (from client doServerSave)
app.post('/api/state/save', (req, res) => {
  resetIfNewDay();
  const { teorico: t, fisico: f, asignaciones: a, historial: h } = req.body;
  // Merge teorico: add new containers but never remove existing ones
  if(t && Object.keys(t).length > 0) {
    Object.keys(t).forEach(cont => {
      if(!state.teorico[cont]) state.teorico[cont] = t[cont];
    });
  }
  if(f) {
    // Merge fisico: server data wins over client (server has field-saves)
    Object.keys(f).forEach(cont => {
      if(!state.fisico[cont]) {
        state.fisico[cont] = f[cont];
      } else if(Array.isArray(f[cont])) {
        f[cont].forEach((item, i) => {
          // Only use client data if server has nothing for this cell
          if(item && item.lastAt && (!state.fisico[cont] || !state.fisico[cont][i] || !state.fisico[cont][i].lastAt)) {
            if(!state.fisico[cont]) state.fisico[cont] = [];
            state.fisico[cont][i] = item;
          }
        });
      }
    });
  }
  if(a) state.asignaciones = a;
  if(h) { state.historial = h; }
  state.version++;
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

app.post('/api/conteo', (req, res) => {
  resetIfNewDay();
  const { cont, data, usuario } = req.body;
  if(!cont || !data) return res.status(400).json({ ok:false });
  state.fisico[cont] = data;
  addHistorial(usuario||'—', 'Conteo guardado', cont);
  res.json({ ok:true, version:state.version });
});

app.post('/api/asign', (req, res) => {
  resetIfNewDay();
  const { cont, name, action, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.asignaciones[cont]) state.asignaciones[cont] = [];
  if(action==='add') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(usuario||'—', 'Asignación', cont+' → '+name);
  } else if(action==='remove') {
    state.asignaciones[cont] = state.asignaciones[cont].filter(n=>n!==name);
    addHistorial(usuario||'—', 'Asignación removida', cont+' ← '+name);
  } else if(action==='self') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(name, 'Auto-asignación', cont);
  }
  state.version++;
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// CDG
app.get('/api/cdg', (req,res) => res.json(state.cdg||{}));

app.post('/api/cdg/save', (req,res) => {
  const { contId, items, usuario } = req.body;
  if(!contId) return res.status(400).json({ok:false});
  if(!state.cdg[contId]) state.cdg[contId]={items:[],status:'open',autor:usuario,fecha:new Date().toLocaleDateString('es')};
  state.cdg[contId].items=items; state.cdg[contId].lastEditor=usuario;
  addHistorial(usuario,'CDG guardado',contId);
  state.version++; scheduleSave();
  res.json({ok:true,version:state.version});
});

app.post('/api/cdg/finalizar', (req,res) => {
  const { contId, items, usuario, traslado } = req.body;
  if(!contId) return res.status(400).json({ok:false});
  state.cdg[contId]={items,status:'closed',autor:usuario,fecha:new Date().toLocaleDateString('es'),traslado};
  const num=traslado||contId;
  if(!state.teorico[num]) {
    state.teorico[num]={type:'Traslados',fromCDG:true,cdgRef:contId,
      items:items.map(i=>({sku:i.sku,desc:i.desc,qty:i.qty,
        raw:{origen:'CDG',status:'CDG Validado',fecha:new Date().toLocaleDateString('es')}}))};
    state.fisico[num]=null;
  } else {
    state.teorico[num].cdgValidated=true;
  }
  addHistorial(usuario,'CDG finalizado → Traslado',contId+' → '+num);
  state.version++; scheduleSave();
  res.json({ok:true,traslado:num,version:state.version});
});

// Costos
const upload = multer({ storage: multer.memoryStorage() });
app.post('/api/costos', upload.single('file'), (req,res) => {
  try {
    const wb=XLSX.read(req.file.buffer,{type:'buffer',raw:false});
    const sn=wb.SheetNames.find(n=>n.toLowerCase().includes('existencia')||n.toLowerCase().includes('sap'))||wb.SheetNames[0];
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
    const hdr=rows[0].map(h=>norm(String(h)));
    const cSku=findCol(hdr,['articulo','artículo','sku','codigo']);
    const cCost=findCol(hdr,['costo promedio','costo']);
    if(cSku<0||cCost<0) return res.status(400).json({ok:false,error:'Columnas no encontradas'});
    const raw={};
    rows.slice(1).forEach(row=>{
      const sku=String(row[cSku]||'').trim(),cost=parseFloat(String(row[cCost]).replace(',','.'));
      if(sku&&!isNaN(cost)&&cost>0){if(!raw[sku])raw[sku]={sum:0,n:0};raw[sku].sum+=cost;raw[sku].n++;}
    });
    state.costos={}; let cnt=0;
    Object.keys(raw).forEach(sku=>{state.costos[sku]=raw[sku].sum/raw[sku].n;cnt++;});
    res.json({ok:true,count:cnt});
  } catch(e){res.status(500).json({ok:false,error:e.message});}
});
app.get('/api/costos',(req,res)=>res.json(state.costos));

// ── Field locking ─────────────────────────────────────────────────────────
app.post('/api/lock', (req, res) => {
  cleanLocks();
  const { cont, idx, user } = req.body;
  const key = cont + ':' + idx;
  const existing = fieldLocks[key];
  // Already locked by someone else and not expired
  if(existing && existing.user !== user && existing.expires > Date.now()) {
    return res.json({ ok:false, lockedBy: existing.user, since: existing.since });
  }
  fieldLocks[key] = { user, since: Date.now(), expires: Date.now() + 90000 };
  state.version++;
  res.json({ ok:true });
});

app.post('/api/unlock', (req, res) => {
  const { cont, idx, user } = req.body;
  const key = cont + ':' + idx;
  if(fieldLocks[key] && fieldLocks[key].user === user) delete fieldLocks[key];
  state.version++;
  res.json({ ok:true });
});

// Auto-save single field
app.post('/api/conteo/field', (req, res) => {
  resetIfNewDay();
  const { cont, idx, fisico, daniado, usuario } = req.body;
  if(cont === undefined || idx === undefined) return res.status(400).json({ok:false});
  if(!state.fisico[cont]) state.fisico[cont] = [];
  if(!Array.isArray(state.fisico[cont])) state.fisico[cont] = [];
  state.fisico[cont][idx] = {
    fisico:   fisico  !== undefined ? fisico  : (state.fisico[cont][idx]||{}).fisico||0,
    daniado:  daniado !== undefined ? daniado : (state.fisico[cont][idx]||{}).daniado||0,
    lastUser: usuario,
    lastAt:   Date.now()
  };
  state.version++;
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// Upload teorico
app.post('/api/upload', upload.single('file'), (req,res) => {
  try {
    resetIfNewDay();
    const wb=XLSX.read(req.file.buffer,{type:'buffer',raw:false});
    const usuario=req.body.usuario||'—';
    let loaded=[];
    wb.SheetNames.forEach(sn=>{
      const nl=sn.toLowerCase();
      const type=nl.includes('traslado')?'Traslados':nl.includes('embarque')?'Embarques':null;
      if(!type) return;
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
      const count=mergeSheet(rows,type);
      if(count>0){loaded.push(sn+'('+count+')');addHistorial(usuario,'Teórico cargado',type+' — '+count+' contenedores');}
    });
    if(loaded.length===0){
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:'',raw:false});
      const count=mergeSheet(rows,'General');
      if(count>0) loaded.push(wb.SheetNames[0]+'('+count+')');
      addHistorial(usuario,'Teórico cargado',loaded.join());
    }
    state.version++; scheduleSave();
    res.json({ok:true,loaded});
  } catch(e){res.status(500).json({ok:false,error:e.message});}
});

// ── Helpers ───────────────────────────────────────────────────────────────
function norm(s){return String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');}
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
    // Don't overwrite CDG-validated containers from teorico upload
    if(state.teorico[cont]&&state.teorico[cont].fromCDG) {
      state.teorico[cont].cdgValidated=true;
      return;
    }
    state.teorico[cont]={items:newConts[cont],type};
    // Preserve existing fisico data — never overwrite conteo work
    if(!state.fisico[cont]) state.fisico[cont]=null;
    // else: keep existing fisico[cont] as-is
  });
  return Object.keys(newConts).length;
}

// ── Start ─────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
// Always start server — even if Supabase fails
app.listen(PORT, () => console.log('Conteo app on port ' + PORT));
// Load state after server is up
loadState().catch(e => console.log('State load failed (non-fatal):', e.message));
