const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const https   = require('https');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

// Límite a 5MB es suficiente para conteos por campo, asignaciones, hallazgos,
// metadata, CDG. Antes era 50MB lo que abría puerta a payloads gigantes que
// matan la RAM del free tier. Uploads de teorico/costos NO pasan por aquí,
// usan multer (memoryStorage), que tiene su propio límite.
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Supabase config (set via environment variables in Render) ─────────────
const SUPABASE_URL = process.env.SUPABASE_URL;  // https://xxx.supabase.co
const SUPABASE_KEY = process.env.SUPABASE_KEY;  // anon/publishable key (RLS allow_all)

// ── Simple Supabase REST client ───────────────────────────────────────────
function supabase(method, table, body, query) {
  return new Promise((resolve, reject) => {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return resolve(null); // No Supabase configured — use memory only
    }
    const url  = new URL(`${SUPABASE_URL}/rest/v1/${table}${query||''}`);
    const data = body ? JSON.stringify(body) : null;
    // FIX BUG LATENTE (mié 20-may-2026): agregar 'resolution=merge-duplicates'
    // al header Prefer. Sin esto, los POST con ?on_conflict=key fallan con
    // HTTP 409 'duplicate key' porque PostgREST no sabe que debe hacer UPSERT.
    // Antes el error se tragaba silenciosamente (resolve null) y el save NUNCA
    // persistía vía server.js — los datos llegaban a Supabase solo por
    // saveDirectToSupabase del cliente (que usa PATCH directo).
    // Con el fix v8 que ahora detecta errores HTTP, este 409 se hizo visible.
    // Esta línea lo arregla de raíz.
    const opts = {
      hostname: url.hostname,
      path:     url.pathname + url.search,
      method,
      headers: {
        'apikey':        SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type':  'application/json',
        'Prefer':        'resolution=merge-duplicates,return=representation'
      }
    };
    if(data) opts.headers['Content-Length'] = Buffer.byteLength(data);
    const req = https.request(opts, res => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => {
        // FIX CRÍTICO (mié 20-may-2026, rev ChatGPT): rechazar la Promise
        // si Supabase responde con error HTTP (4xx/5xx). Antes resolvíamos
        // SIEMPRE con null aunque el server hubiera rechazado el guardado,
        // lo que dejaba a saveDailyState y dbSet creyendo que todo OK.
        // Esto era la causa raíz REAL del bug que vimos en producción.
        if(res.statusCode >= 400) {
          var errMsg = 'Supabase HTTP ' + res.statusCode;
          try {
            var errBody = raw ? JSON.parse(raw) : null;
            if(errBody && errBody.message) errMsg += ': ' + errBody.message;
            else if(errBody && errBody.error) errMsg += ': ' + errBody.error;
            else if(raw) errMsg += ': ' + raw.slice(0, 200);
          } catch(e) {
            if(raw) errMsg += ': ' + raw.slice(0, 200);
          }
          return reject(new Error(errMsg));
        }
        try { resolve(raw ? JSON.parse(raw) : null); }
        catch(e) { resolve(null); }
      });
    });
    req.on('error', reject);
    if(data) req.write(data);
    req.end();
  });
}

// FIX CRÍTICO (mié 20-may-2026, rev Claude2): timeout para evitar que un
// Supabase lento cuelgue al cliente hasta el límite de Render/Cloudflare
// (~100s). Si dbSet tarda más de 15s, rechazamos para que el endpoint
// responda con error custom al cliente. El dbSet sigue en vuelo en el
// background — si eventualmente completa, mejor (UPSERT idempotente).
function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(
      () => reject(new Error(label + ' timeout después de ' + ms + 'ms')),
      ms
    ))
  ]);
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
  await supabase('POST', 'app_state', data, '?on_conflict=key');
}

// ── In-memory state ───────────────────────────────────────────────────────
let state = {
  teorico:        {},
  fisico:         {},
  asignaciones:   {},
  historial:      [],
  costos:         {},
  cdg:            {},
  puertas:        {},
  hallazgos:      [],
  conteoMetadata: {},
  alertasWMS:     {},  // FIX (server v15): tracking de cambios del WMS por contenedor congelado
  date:           new Date().toDateString(),
  version:        0
};

// Field locks: { "contId:rowIdx": { user, since, expires } }
let fieldLocks = {};
const LOCK_TIMEOUT   = 90000; // 90 seconds max
const ACTIVE_TIMEOUT = 15000;

let activeUsers = {};

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

function getActiveUsers() {
  const now = Date.now();
  Object.keys(activeUsers).forEach(n => {
    if(now - activeUsers[n] > ACTIVE_TIMEOUT) delete activeUsers[n];
  });
  return Object.keys(activeUsers);
}

// Build the full daily_state payload from current memory state
function buildDailyStatePayload() {
  return {
    teorico:        state.teorico,
    fisico:         state.fisico,
    asignaciones:   state.asignaciones,
    historial:      state.historial.slice(-100),
    cdg:            state.cdg,
    puertas:        state.puertas        || {},
    hallazgos:      state.hallazgos      || [],
    conteoMetadata: state.conteoMetadata || {},
    alertasWMS:     state.alertasWMS     || {},  // FIX (server v15)
    date:           state.date,
    version:        state.version
  };
}

function saveDailyState(label) {
  return dbSet('daily_state', buildDailyStatePayload())
    .catch(e => console.log((label||'save')+' error:', e.message));
}
// FIX (lun 1-jun-2026, v19): versión estricta que SÍ rechaza si dbSet falla.
// saveDailyState() tiene .catch interno — withTimeout() recibe una promesa
// resuelta aunque Supabase devuelva 500, y el endpoint respondería ok:true
// sin haber persistido. saveDailyStateStrict() propaga el error para que el
// caller pueda responder 500 al cliente en operaciones críticas (ej. cerrar v2).
function saveDailyStateStrict(label) {
  return dbSet('daily_state', buildDailyStatePayload());
}

// ── Load state from Supabase on startup ───────────────────────────────────
async function loadState() {
  try {
    const saved = await dbGet('daily_state');
    if(saved && saved.teorico) {
      state = { ...state, ...saved };
      console.log('State restored from Supabase ✓ date:', saved.date);
    } else {
      console.log('No saved state found — starting fresh');
    }
    if(saved && saved.puertas)        state.puertas        = saved.puertas;
    if(saved && saved.hallazgos)      state.hallazgos      = saved.hallazgos;
    if(saved && saved.conteoMetadata) state.conteoMetadata = saved.conteoMetadata;
    if(saved && saved.alertasWMS)     state.alertasWMS     = saved.alertasWMS;  // FIX (server v15)

    const savedCostos = await dbGet('costos_state');
    if(savedCostos && savedCostos.costos) {
      state.costos = savedCostos.costos;
      console.log('Costos restored:', Object.keys(state.costos).length, 'SKUs');
    }
  } catch(e) {
    console.log('Could not load from Supabase:', e.message);
  }
}

// ── Save state to Supabase (debounced) ───────────────────────────────────
// Debounce a 1500ms para agrupar escrituras frecuentes (varios users tecleando
// simultáneamente). Antes era 200ms — provocaba serializar el state completo
// (~7MB) por cada tecleo, lo que dispara el uso de RAM en Render free (512MB)
// y causa "heap out of memory". 1.5s aún es lo suficientemente rápido para no
// perder datos en cierres normales (beforeunload del frontend hace flush).
let saveTimer = null;
function scheduleSave() {
  if(saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { saveDailyState('Debounced save'); }, 1500);
}

function addHistorial(usuario, accion, detalle) {
  state.historial.push({
    hora: new Date().toLocaleTimeString('es'),
    usuario,
    accion,
    detalle: detalle || ''
  });
  if(state.historial.length > 200) state.historial.shift();
  state.version++;
  scheduleSave();
}

function publicState() {
  return {
    teorico:        state.teorico,
    fisico:         state.fisico,
    asignaciones:   state.asignaciones,
    historial:      state.historial.slice(-50),
    cdg:            state.cdg,
    puertas:        state.puertas        || {},
    hallazgos:      state.hallazgos      || [],
    conteoMetadata: state.conteoMetadata || {},
    alertasWMS:     state.alertasWMS     || {},  // FIX (rev ChatGPT BLOQUEANTE v5.2.22): sin esto el cliente nunca recibía las alertas y el badge no aparecía
    date:           state.date,
    version:        state.version,
    activeUsers:    getActiveUsers(),
    locks:          getLocks()
  };
}

// ── API ───────────────────────────────────────────────────────────────────
app.post('/api/heartbeat', (req, res) => {
  const { name } = req.body;
  if(name) activeUsers[name] = Date.now();
  res.json({ active: getActiveUsers() });
});

app.get('/api/state', (req, res) => {
  var ps = publicState();
  try {
    // Diagnóstico de tamaño: permite monitorear crecimiento del state sin crashear.
    // FIX (mar 2-jun-2026, server v20): campo stateSizeBytes agregado para auditoría.
    ps.stateSizeBytes = Buffer.byteLength(JSON.stringify(buildDailyStatePayload()), 'utf8');
  } catch(e) { /* no bloquear el estado si falla el cálculo */ }
  res.json(ps);
});

// Full state save (from client). Merges, never deletes existing data.
app.post('/api/conteo', (req, res) => {
  const { cont, data, usuario } = req.body;
  if(!cont || !data) return res.status(400).json({ ok:false });
  state.fisico[cont] = data;
  addHistorial(usuario||'—', 'Conteo guardado', cont);
  // Debounced — agrupa escrituras concurrentes. addHistorial ya llama scheduleSave,
  // así que aquí no hace falta llamar de nuevo.
  res.json({ ok:true, version:state.version });
});

app.post('/api/asign', (req, res) => {
  const { cont, name, action, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.asignaciones[cont]) state.asignaciones[cont] = [];
  if(action === 'add') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(usuario||'—', 'Asignación', cont+' → '+name);
  } else if(action === 'remove') {
    state.asignaciones[cont] = state.asignaciones[cont].filter(n => n !== name);
    addHistorial(usuario||'—', 'Asignación removida', cont+' ← '+name);
  } else if(action === 'self') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(name, 'Auto-asignación', cont);
  }
  state.version++;
  // Debounced — addHistorial dentro de cada rama ya llamó scheduleSave.
  res.json({ ok:true, version:state.version });
});

// ── CDG ───────────────────────────────────────────────────────────────────
app.get('/api/cdg', (req, res) => res.json(state.cdg||{}));

app.post('/api/cdg/save', (req, res) => {
  const { contId, items, usuario, tipo, fotoGral, fotos } = req.body;
  if(!contId) return res.status(400).json({ ok:false });
  // FIX (sáb 30-may-2026, v19): validación defensiva de items.
  // Si items llega null/undefined (payload mal formado, tablet vieja, etc.),
  // rechazar antes de tocar el state. Sin esto, items:null pasa la guardia
  // ((null||[]).length = 0) y luego state.cdg[contId].items = null rompe el conteo.
  if(!Array.isArray(items)) {
    return res.status(400).json({
      ok: false,
      error: 'No se guardó: el conteo llegó sin lista de líneas válida. Recargá la pantalla e intentá de nuevo.'
    });
  }
  if(!state.cdg[contId]) {
    state.cdg[contId] = {
      items:  [],
      status: 'open',
      autor:  usuario,
      fecha:  new Date().toLocaleDateString('es')
    };
  }
  // FIX (sáb 30-may-2026, v19): protección contra sobreescritura accidental.
  // Escenario: Julio guarda 50 items. Otro usuario abre CDG con lista vacía y
  // guarda 1 item → borra los 50 de Julio (Bug B1 CDG v1).
  // Guardia: si ya hay items y el nuevo array tiene menos de la mitad Y el
  // usuario no es el autor original → rechazar con 409 y mensaje claro.
  // Nota: items ya validado como Array arriba, usar .length directo.
  var existentes    = (state.cdg[contId].items || []).length;
  var nuevos        = items.length;
  var autorOriginal = state.cdg[contId].autor || state.cdg[contId].lastEditor;
  if(existentes > 2 && nuevos < existentes / 2 && autorOriginal && autorOriginal !== usuario) {
    console.log('CDG save BLOQUEADO: ' + usuario + ' intentó reducir ' + contId + ' de ' + existentes + ' a ' + nuevos + ' items (autor: ' + autorOriginal + ')');
    return res.status(409).json({
      ok:    false,
      error: 'Este conteo ya tiene ' + existentes + ' líneas. Recargá el conteo antes de guardar. Para reemplazarlo usá desbloqueo de supervisor o CDG v2 multiusuario.'
    });
  }
  state.cdg[contId].items      = items;
  state.cdg[contId].lastEditor = usuario;
  if(tipo)     state.cdg[contId].tipo     = tipo;
  if(fotoGral) state.cdg[contId].fotoGral = fotoGral;
  if(fotos)    state.cdg[contId].fotos    = fotos;
  addHistorial(usuario, 'CDG guardado', contId);
  state.version++;
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// CDG Unlock (supervisor only — enforced client-side)
app.post('/api/cdg/unlock', (req, res) => {
  const { contId, usuario } = req.body;
  if(!contId) return res.status(400).json({ ok:false });
  if(state.cdg[contId]) {
    state.cdg[contId].bloqueado       = false;
    state.cdg[contId].desbloqueadoPor = usuario;
    state.cdg[contId].desbloqueadoTs  = new Date().toLocaleString('es');
  }
  // Also unblock in teorico
  if(state.teorico[contId]) state.teorico[contId].cdgBloqueado = false;
  addHistorial(usuario||'—', 'Desbloqueó CDG', contId);
  state.version++;
  // Debounced — addHistorial ya disparó scheduleSave.
  res.json({ ok:true });
});


// ── cdgEnriquecerItemsWMS ──────────────────────────────────────────────────
// FIX (lun 1-jun-2026, server v20 rev): devuelve la UNIÓN de items CDG + SKUs WMS.
//
// CORRECCIÓN vs versión anterior (solo hacía map sobre items CDG):
//   La versión anterior agregaba teoricoWMS a SKUs que CDG sí contó, pero dejaba
//   invisible cualquier SKU que existiera en WMS y CDG no hubiera contado.
//   Un SKU WMS con 50 unidades y qty CDG = 0 quedaba fuera de Hamilton → faltante
//   invisible. Esto era bloqueante.
//
// Comportamiento correcto:
//   1. SKU en CDG y WMS → qty = Validado CDG, teoricoWMS = cantidad WMS.
//   2. SKU en CDG, NO en WMS → item sin cambios (sin teoricoWMS). Fallback legacy OK.
//   3. SKU en WMS, NO en CDG → nueva línea: qty=0, teoricoWMS=cantidad WMS.
//      Aparece en Hamilton: Teórico s/WMS=N, Validado CDG=0.
//
// DEUDA TÉCNICA (no resuelta en este fix):
//   Los exports de Traslados CDG y updateSummaryCard usan item.qty para calcular
//   diferencias. Con los nuevos items WMS-only (qty=0), el resumen Hamilton puede
//   contar más faltantes de lo que muestra la tabla si no se actualiza esa lógica.
//   Fix pendiente: ajustar buildReporte / exportCDG para leer teoricoWMS como base
//   de diferencia cuando esCDG, igual que renderConteo ya hace.
//
// Si no hay WMS o falla la consulta → devuelve items sin cambios (backward-compatible).
// Se usa en /api/cdg/finalizar (v1) y en el cierre CDG v2.
async function cdgEnriquecerItemsWMS(licenciaId, items) {
  if(!SUPABASE_URL || !SUPABASE_KEY || !licenciaId) return items || [];
  try {
    var wmsRows = await supabase('GET', 'cdg_wms', null,
      '?licencia_id=eq.' + encodeURIComponent(cdgNormId(licenciaId)) + '&limit=1');
    if(!Array.isArray(wmsRows) || !wmsRows.length) return items || []; // sin WMS — OK
    var wmsRow = wmsRows[0];
    var skusWMS = typeof wmsRow.skus === 'string' ? JSON.parse(wmsRow.skus) : (wmsRow.skus || []);

    // Mapa WMS: SKU_NORMALIZADO → { cantidad, descripcion }
    var wmsMap = {};
    skusWMS.forEach(function(s){
      var sk = String(s.sku || '').trim().toUpperCase();
      if(!wmsMap[sk]) {
        wmsMap[sk] = { cantidad: 0, descripcion: s.descripcion || '' };
      }
      wmsMap[sk].cantidad += (Number(s.cantidad) || 0);
      // Preferir descripción no vacía si la anterior era vacía
      if(!wmsMap[sk].descripcion && s.descripcion) wmsMap[sk].descripcion = s.descripcion;
    });

    // Conjunto de SKUs ya cubiertos por CDG (para detectar solo-WMS al final)
    var itemsCdg = Array.isArray(items) ? items : [];
    var skusCDG = {};
    var resultado = itemsCdg.map(function(item) {
      var sk = String(item.sku || '').trim().toUpperCase();
      skusCDG[sk] = true;
      if(wmsMap[sk] !== undefined) {
        // SKU en CDG y en WMS → agregar teoricoWMS
        return Object.assign({}, item, { teoricoWMS: wmsMap[sk].cantidad });
      }
      // SKU en CDG pero no en WMS → sin teoricoWMS (fallback legacy)
      return item;
    });

    // SKUs en WMS que CDG NO contó → agregar como líneas con qty=0
    var tipoEfectivo = 'CDG'; // valor por defecto; el llamador lo conoce pero no se pasa aquí
    Object.keys(wmsMap).forEach(function(sk){
      if(skusCDG[sk]) return; // ya cubierto por CDG
      var entrada = wmsMap[sk];
      resultado.push({
        sku:         sk,
        desc:        entrada.descripcion || '',
        qty:         0,           // Validado CDG = 0 (no fue contado)
        teoricoWMS:  entrada.cantidad,
        raw:         { origen: tipoEfectivo, status: tipoEfectivo + ' Validado' },
        _soloWMS:    true         // marca interna para diagnóstico; no afecta renderConteo
      });
    });

    return resultado;
  } catch(e) {
    console.log('cdgEnriquecerItemsWMS warn (non-fatal):', e.message);
    return items || []; // fallo silencioso — no romper el cierre CDG
  }
}

app.post('/api/cdg/finalizar', async (req, res) => {
  const { contId, items, usuario, traslado, tipo, fotoGral, fotos, bloqueado } = req.body;
  if(!contId) return res.status(400).json({ ok:false });
  state.cdg[contId] = {
    items,
    status:    'closed',
    autor:     usuario,
    fecha:     new Date().toLocaleDateString('es'),
    traslado,
    tipo:      tipo || 'CDG',
    fotoGral:  fotoGral || null,
    fotos:     fotos    || null,
    bloqueado: bloqueado || false
  };
  const num = traslado || contId;
  if(!state.teorico[num]) {
    // FIX (sáb 23-may-2026): raw.origen y raw.status ahora reflejan el tipo real
    // (KTM, CDG, Otros). Antes era hardcoded 'CDG' lo que hacía que el export
    // de Traslados mostrara "CDG" para TODOS los contenedores finalizados desde
    // esta sección, incluyendo los KTM. Erick + Ever confirmaron 23-may que KTM
    // debe verse como categoría distinta en el export.
    var tipoEfectivo = tipo || 'CDG';
    // FIX (lun 1-jun-2026, server v20): enriquecer items con teoricoWMS desde cdg_wms.
    // qty = Validado CDG (el conteo físico CDG). teoricoWMS = teórico del WMS si existe.
    var itemsEnriquecidos = await cdgEnriquecerItemsWMS(num, items.map(i => ({
      sku:  i.sku,
      desc: i.desc,
      qty:  i.qty,
      raw:  { origen:tipoEfectivo, status:tipoEfectivo+' Validado', fecha:new Date().toLocaleDateString('es') }
    })));
    state.teorico[num] = {
      type:         'Traslados',
      fromCDG:      true,
      cdgRef:       contId,
      cdgValidado:  true,
      cdgBloqueado: true,
      cdgTipo:      tipoEfectivo,
      items: itemsEnriquecidos
    };
    state.fisico[num] = null;
  } else {
    state.teorico[num].cdgValidado = true;
    // Asegurar que el tipo quede registrado también si el contenedor ya existía
    // (defensivo, no rompe nada si ya estaba)
    if(tipo && !state.teorico[num].cdgTipo) state.teorico[num].cdgTipo = tipo;
  }
  addHistorial(usuario, 'CDG finalizado → Traslado', contId+' → '+num);
  state.version++;

  // FIX CRÍTICO (mié 20-may-2026): mismo patrón que upload. Cancelar debounced
  // pending y awaitear el save, responder con error si Supabase falla.
  // FIX (rev Claude2): timeout 15s. FIX (rev ChatGPT): mensaje operativo.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'CDG final save'
    );
    console.log('CDG final save: persisted to Supabase ✓');
  } catch(saveErr) {
    console.log('CDG final save FAILED:', saveErr.message);
    return res.status(500).json({
      ok: false,
      error: 'El CDG sí se procesó, pero NO quedó guardado. No cierres la app. Volvé a finalizarlo. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, traslado:num, version:state.version });
});

// ── Costos ────────────────────────────────────────────────────────────────
app.post('/api/costos', upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const sn = wb.SheetNames.find(n =>
      n.toLowerCase().includes('existencia') || n.toLowerCase().includes('sap')
    ) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header:1, defval:'', raw:false });
    const hdr  = rows[0].map(h => norm(String(h)));
    const cSku  = findCol(hdr, ['articulo','artículo','sku','codigo']);
    const cCost = findCol(hdr, ['costo promedio','costo']);
    if(cSku < 0 || cCost < 0) {
      return res.status(400).json({ ok:false, error:'Columnas no encontradas' });
    }
    const raw = {};
    rows.slice(1).forEach(row => {
      const sku  = String(row[cSku]||'').trim();
      const cost = parseFloat(String(row[cCost]).replace(',','.'));
      if(sku && !isNaN(cost) && cost > 0) {
        if(!raw[sku]) raw[sku] = { sum:0, n:0 };
        raw[sku].sum += cost;
        raw[sku].n++;
      }
    });
    state.costos = {};
    let cnt = 0;
    Object.keys(raw).forEach(sku => {
      state.costos[sku] = raw[sku].sum / raw[sku].n;
      cnt++;
    });
    res.json({ ok:true, count:cnt });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// Save pre-processed costos from client
// FIX CRÍTICO (mié 20-may-2026): dbSet ya no se hace sin await. Mismo bug
// que upload teorico — si Render reiniciaba después de responder, los costos
// se perdían silenciosamente.
// FIX (rev Claude2): timeout 15s. FIX (rev ChatGPT): mensaje operativo.
app.post('/api/costos-save', async (req, res) => {
  const { costos: c } = req.body;
  if(!c || Object.keys(c).length === 0) {
    return res.json({ ok:false });
  }
  state.costos = c;
  try {
    await withTimeout(
      dbSet('costos_state', { costos: c }),
      15000,
      'Costos save'
    );
    console.log('Costos save: persisted to Supabase ✓');
    res.json({ ok:true, count: Object.keys(c).length });
  } catch(e) {
    console.log('Costos save FAILED:', e.message);
    res.status(500).json({
      ok: false,
      error: 'Los costos sí se procesaron, pero NO quedaron guardados. No cierres la app. Volvé a subirlos. (' + e.message + ')'
    });
  }
});

app.get('/api/costos', (req, res) => res.json(state.costos));

// ── Hallazgos ─────────────────────────────────────────────────────────────
app.post('/api/hallazgo', (req, res) => {
  const { hallazgo, action, id } = req.body;
  if(!state.hallazgos) state.hallazgos = [];
  if(action === 'add' && hallazgo) {
    state.hallazgos.push(hallazgo);
  } else if(action === 'edit' && id) {
    const idx = state.hallazgos.findIndex(h => h.id === id);
    if(idx >= 0) state.hallazgos[idx] = hallazgo;
  } else if(action === 'delete' && id) {
    // FIX (mié 20-may-2026 noche): soporte para eliminar hallazgos desde el
    // botón "🗑 Eliminar" del cliente (visible solo para Erick Vela). Sin
    // este case el server caería en el branch implícito y reescribiría
    // Supabase reviviendo el hallazgo eliminado.
    state.hallazgos = state.hallazgos.filter(h => h.id !== id);
  }
  state.version++;
  // Debounced — hallazgos pueden venir en ráfaga
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// ── Metadata (puerta, fechaIngreso, fechaFurgon, placas per container) ────
app.post('/api/metadata', (req, res) => {
  const { cont, metadata, puerta, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.conteoMetadata) state.conteoMetadata = {};
  if(!state.puertas)        state.puertas        = {};
  if(metadata)             state.conteoMetadata[cont] = metadata;
  if(puerta !== undefined) state.puertas[cont]        = puerta;
  addHistorial(usuario||'—', 'Metadata actualizada', cont);
  state.version++;
  // Debounced — addHistorial ya disparó scheduleSave.
  res.json({ ok:true, version:state.version });
});

// FIX (mar 19-may-2026): edición manual de fechaCarga del teorico.
// Permite asignar/cambiar la fecha de trabajo de un contenedor desde la UI,
// por ejemplo para contenedores históricos que no tenían fechaCarga.
// Formato esperado: 'YYYY-MM-DD' o cadena vacía ('') para limpiar.
//
// FIX (rev Claude2): valida también la fecha calendáricamente, no solo
// el formato. Rechaza overflows como 2026-02-30 o 9999-99-99.
//
// FIX (mié 20-may-2026, post-deploy v8.1): mismo patrón que upload teorico.
// Antes el endpoint dependía de addHistorial → scheduleSave debounced (1.5s)
// para persistir. Si Render reiniciaba en ese gap, la edición se perdía
// silenciosamente. Ahora cancelamos pending, awaitamos con timeout, y
// respondemos error si Supabase falla. Igual que /api/upload y CDG finalizar.
app.post('/api/teorico/fecha-carga', async (req, res) => {
  const { cont, fechaCarga, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false, error:'falta cont' });
  if(!state.teorico[cont]) return res.status(404).json({ ok:false, error:'cont no existe' });
  // Validar formato + fecha calendárica real
  if(fechaCarga !== '' && fechaCarga !== null && fechaCarga !== undefined) {
    var s = String(fechaCarga);
    if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      return res.status(400).json({ ok:false, error:'formato debe ser YYYY-MM-DD o vacío' });
    }
    // Validar fecha calendárica: el truco es comparar contra round-trip via Date.
    // Si entró 2026-02-30, new Date lo normaliza a 2026-03-02, y la comparación falla.
    var d = new Date(s + 'T00:00:00Z');
    if(isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) {
      return res.status(400).json({ ok:false, error:'fecha calendárica inválida' });
    }
  }
  // FIX (mié 28-may-2026, server v18): rollback ante fallo de persistencia.
  // Antes: si Supabase fallaba, el cambio quedaba aplicado en memoria pero NO
  // en Supabase. El cliente veía "error" pero la fecha cambiada en pantalla
  // hasta el próximo polling. Confusión + divergencia.
  // Ahora: snapshot del valor previo + version + historial. En catch, restaura.
  //
  // FIX (rev cruzada Claude3 + ChatGPT, post-v18 borrador): addHistorial()
  // muta state.version e state.historial internamente. Snapshot debe capturar
  // AMBOS antes de cualquier mutación para revertir correctamente. Antes del
  // fix, el rollback hacía state.version-- (revertía solo 1 de 2 incrementos)
  // y dejaba el historial con la entrada del intento fallido.
  var prevFechaCarga    = state.teorico[cont].fechaCarga;
  var snapshotVersion   = state.version;
  var snapshotHistorial = state.historial.slice();

  state.teorico[cont].fechaCarga = fechaCarga || null;
  addHistorial(usuario||'—', 'Cambió fecha de trabajo a ' + (fechaCarga || '(vacío)'), cont);
  state.version++;

  // Cancelar debounced pending y await el save para garantizar persistencia.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'Fecha-carga save'
    );
    console.log('Fecha-carga save: persisted to Supabase ✓');
  } catch(saveErr) {
    console.log('Fecha-carga save FAILED:', saveErr.message);
    // Rollback completo: restaurar valor + historial + version (revierte
    // tanto el version++ del endpoint como el version++ interno de addHistorial)
    state.teorico[cont].fechaCarga = prevFechaCarga;
    state.historial = snapshotHistorial;
    state.version   = snapshotVersion;
    return res.status(500).json({
      ok: false,
      error: 'La fecha NO quedó guardada. Reintentá. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, version:state.version, fechaCarga: state.teorico[cont].fechaCarga });
});

// FIX (sáb 23-may-2026): endpoint para clasificación manual de contenedores.
// Resuelve el bug B1 (server-memory desactualizada): antes el cliente llamaba
// saveDirectToSupabase con PATCH directo, lo que persistía en Supabase pero
// dejaba al server.js con state.teorico EN MEMORIA sin la clasificación.
// En el siguiente GET /api/state el server respondía con su memoria vieja
// y el cliente perdía visualmente la clasificación manual.
//
// Ahora el cliente llama este endpoint. El server:
//   1. Actualiza state.teorico[cont].clasificacion + clasificacionManual
//   2. Awaitea el save a Supabase con timeout 15s (mismo patrón que fecha-carga)
//   3. Cancela debounced pending para que no pise el cambio
//
// Permite limpiar el override (clasificacion=null, manual=false) para
// soportar la lógica "quitar manual si coincide con el automático" del UI.
app.post('/api/clasificacion/set', async (req, res) => {
  const { cont, clasificacion, manual, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false, error:'falta cont' });
  if(!state.teorico[cont]) return res.status(404).json({ ok:false, error:'cont no existe' });

  // Validar clasificación si viene con valor
  if(clasificacion !== null && clasificacion !== undefined && clasificacion !== '') {
    var allowed = ['Auditado', 'En Revisión', 'No auditado'];
    if(!allowed.includes(String(clasificacion))) {
      return res.status(400).json({ ok:false, error:'clasificacion inválida (debe ser ' + allowed.join('|') + ')' });
    }
  }

  // FIX (mié 28-may-2026, server v18): rollback ante fallo de persistencia.
  // Snapshot ANTES de mutar (los 2 campos que vamos a tocar).
  // FIX (rev cruzada post-v18 borrador): también capturamos version + historial
  // porque addHistorial() los muta internamente. Sin esto, el rollback dejaba
  // version en N+1 y el historial con la entrada del intento fallido.
  var prevClasificacion       = state.teorico[cont].clasificacion;
  var prevClasificacionManual = state.teorico[cont].clasificacionManual;
  var prevTeniaClasif         = 'clasificacion'       in state.teorico[cont];
  var prevTeniaManual         = 'clasificacionManual' in state.teorico[cont];
  var snapshotVersion         = state.version;
  var snapshotHistorial       = state.historial.slice();

  // Si clasificacion es null/vacío Y manual no es true → quitar override
  if((clasificacion === null || clasificacion === undefined || clasificacion === '') && !manual) {
    delete state.teorico[cont].clasificacion;
    delete state.teorico[cont].clasificacionManual;
    addHistorial(usuario||'—', 'Quitó clasificación manual', cont);
  } else {
    state.teorico[cont].clasificacion = clasificacion;
    state.teorico[cont].clasificacionManual = !!manual;
    addHistorial(usuario||'—', 'Clasificación ' + (manual ? 'manual' : 'auto') + ': ' + clasificacion, cont);
  }
  state.version++;

  // Cancelar debounced pending y await el save para garantizar persistencia.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'Clasificación save'
    );
    console.log('Clasificación save: persisted to Supabase ✓');
  } catch(saveErr) {
    console.log('Clasificación save FAILED:', saveErr.message);
    // Rollback completo: restaurar campos previos (respetando si existían),
    // historial y version (revierte tanto el version++ del endpoint como el
    // version++ interno de addHistorial).
    if(prevTeniaClasif) state.teorico[cont].clasificacion = prevClasificacion;
    else delete state.teorico[cont].clasificacion;
    if(prevTeniaManual) state.teorico[cont].clasificacionManual = prevClasificacionManual;
    else delete state.teorico[cont].clasificacionManual;
    state.historial = snapshotHistorial;
    state.version   = snapshotVersion;
    return res.status(500).json({
      ok: false,
      error: 'La clasificación NO quedó guardada. Reintentá. (' + saveErr.message + ')'
    });
  }

  res.json({
    ok: true,
    version: state.version,
    clasificacion: state.teorico[cont].clasificacion || null,
    manual: !!state.teorico[cont].clasificacionManual
  });
});

// FIX (sáb 23-may-2026): endpoint para borrar CDG.
// Misma motivación que /api/clasificacion/set: evitar bug B1.
// Cuando se borra un CDG, también se borra el contenedor Traslado que
// se creó a partir de él (decisión Erick 23-may: limpio total).
//
// Identificación del Traslado asociado:
//   - El CDG tiene state.cdg[contId].traslado (asignado en /api/cdg/finalizar)
//   - El Traslado tiene state.teorico[num].cdgRef === contId
//   - Borramos AMBOS con todas sus referencias (teorico + fisico + asignaciones)
app.post('/api/cdg/delete', async (req, res) => {
  const { contId, usuario } = req.body;
  if(!contId) return res.status(400).json({ ok:false, error:'falta contId' });
  if(!state.cdg[contId]) return res.status(404).json({ ok:false, error:'CDG no existe' });

  // Identificar el contenedor Traslado asociado
  var trasladoNum = state.cdg[contId].traslado || null;

  // FIX (mié 28-may-2026, server v18): rollback ante fallo de persistencia.
  // ANTES de borrar nada, snapshot completo de TODO lo que vamos a borrar.
  // Esto incluye: el CDG, el Traslado (en sus 6 ubicaciones), y los traslados
  // detectados por fallback (cdgRef). El rollback restaura cada referencia.
  //
  // FIX (rev cruzada post-v18 borrador): también capturamos version + historial
  // porque addHistorial() los muta internamente. Sin esto, el rollback dejaba
  // version en N+1 y el historial con la entrada del intento fallido.
  var rollbackData = {
    cdg:           state.cdg[contId],
    trasladoPrimary: null,    // el del .traslado
    trasladosFallback: []     // los detectados por cdgRef
  };
  var snapshotVersion   = state.version;
  var snapshotHistorial = state.historial.slice();

  // Helper para capturar TODO el sub-state asociado a un traslado.
  // Usa 'in' para distinguir "propiedad existe con valor null" vs "propiedad ausente".
  // Esto es importante para que el rollback restaure exactamente el estado previo.
  function captureTraslado(num) {
    return {
      num:           num,
      teorico:       state.teorico[num],
      teniaFisico:        ('fisico' in state && num in state.fisico),
      fisico:        (state.fisico && num in state.fisico) ? state.fisico[num] : null,
      teniaAsign:         ('asignaciones' in state && num in state.asignaciones),
      asignaciones:  (state.asignaciones && num in state.asignaciones) ? state.asignaciones[num] : null,
      teniaPuerta:        (state.puertas && num in state.puertas),
      puerta:        (state.puertas && num in state.puertas) ? state.puertas[num] : null,
      teniaMetadata:      (state.conteoMetadata && num in state.conteoMetadata),
      metadata:      (state.conteoMetadata && num in state.conteoMetadata) ? state.conteoMetadata[num] : null,
      teniaAlerta:        (state.alertasWMS && num in state.alertasWMS),
      alertaWMS:     (state.alertasWMS && num in state.alertasWMS) ? state.alertasWMS[num] : null
    };
  }

  if(trasladoNum && state.teorico[trasladoNum] && state.teorico[trasladoNum].cdgRef === contId) {
    rollbackData.trasladoPrimary = captureTraslado(trasladoNum);
  }
  Object.keys(state.teorico).forEach(function(k){
    if(state.teorico[k] && state.teorico[k].cdgRef === contId
       && k !== trasladoNum) {  // ya capturado arriba si match
      rollbackData.trasladosFallback.push(captureTraslado(k));
    }
  });

  // Helper de borrado seguro: usa 'in' en vez de truthy check. Antes:
  //   if(state.fisico[trasladoNum]) delete state.fisico[trasladoNum]
  // no borraba si el valor era null (común en contenedores no contados).
  // Ahora: si la propiedad existe (aunque sea null), se borra.
  function deleteTrasladoRefs(num) {
    delete state.teorico[num];
    if(state.fisico         && num in state.fisico)         delete state.fisico[num];
    if(state.asignaciones   && num in state.asignaciones)   delete state.asignaciones[num];
    if(state.puertas        && num in state.puertas)        delete state.puertas[num];
    if(state.conteoMetadata && num in state.conteoMetadata) delete state.conteoMetadata[num];
    if(state.alertasWMS     && num in state.alertasWMS)     delete state.alertasWMS[num];
  }

  // Borrar el CDG
  delete state.cdg[contId];

  // Borrar el contenedor Traslado asociado si existe y vino del CDG
  if(trasladoNum && rollbackData.trasladoPrimary) {
    deleteTrasladoRefs(trasladoNum);
  }

  // Fallback: por si el traslado no estaba en .traslado pero sí en teorico con cdgRef
  // (escenario raro pero defensivo)
  rollbackData.trasladosFallback.forEach(function(t){
    deleteTrasladoRefs(t.num);
  });

  addHistorial(usuario||'—', 'CDG eliminado (+ Traslado asociado)', contId + (trasladoNum ? ' → ' + trasladoNum : ''));
  state.version++;

  // Cancelar debounced pending y await el save
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'CDG delete save'
    );
    console.log('CDG delete save: persisted to Supabase ✓ (CDG ' + contId + ', Traslado ' + (trasladoNum||'—') + ')');
  } catch(saveErr) {
    console.log('CDG delete save FAILED:', saveErr.message);
    // Rollback completo: CDG, traslado primario, traslados fallback, historial, version.
    // Usa los flags tenia* para distinguir "propiedad existía con valor null/falsy"
    // vs "propiedad no existía" y restaurar el estado EXACTO previo.
    state.cdg[contId] = rollbackData.cdg;
    function restoreTraslado(t) {
      state.teorico[t.num] = t.teorico;
      if(t.teniaFisico   && state.fisico)         state.fisico[t.num]         = t.fisico;
      if(t.teniaAsign    && state.asignaciones)   state.asignaciones[t.num]   = t.asignaciones;
      if(t.teniaPuerta   && state.puertas)        state.puertas[t.num]        = t.puerta;
      if(t.teniaMetadata && state.conteoMetadata) state.conteoMetadata[t.num] = t.metadata;
      if(t.teniaAlerta   && state.alertasWMS)     state.alertasWMS[t.num]     = t.alertaWMS;
    }
    if(rollbackData.trasladoPrimary) restoreTraslado(rollbackData.trasladoPrimary);
    rollbackData.trasladosFallback.forEach(restoreTraslado);
    state.historial = snapshotHistorial;
    state.version   = snapshotVersion;
    return res.status(500).json({
      ok: false,
      error: 'El CDG NO se borró (memoria restaurada). Reintentá. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, version:state.version, deletedCDG:contId, deletedTraslado:trasladoNum });
});

// FIX (dom 24-may-2026 PM, server v15): endpoint para "Sincronizar con WMS".
// Cuando un contenedor está congelado (tiene físico contado) y el WMS tiene
// cambios pendientes en state.alertasWMS, el supervisor puede aplicar esos
// cambios manualmente con este endpoint.
//
// Comportamiento:
//   1. Validar que la alerta existe para el contenedor
//   2. Aplicar los items del WMS al teorico (sobreescribe items)
//   3. Preservar fechaCarga, clasificacion, etc. (Object.assign)
//   4. Eliminar la alerta del state.alertasWMS
//   5. Registrar en historial
//
// Permisos: solo supervisores (validado en el cliente; el server confía
// pero loggea quién hizo la sincronización).
app.post('/api/wms/sincronizar', async (req, res) => {
  const { cont, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false, error:'falta cont' });
  if(!state.alertasWMS || !state.alertasWMS[cont]) {
    return res.status(404).json({ ok:false, error:'no hay alerta WMS pendiente para este contenedor' });
  }
  if(!state.teorico[cont]) {
    return res.status(404).json({ ok:false, error:'contenedor no existe en teorico' });
  }

  var alerta = state.alertasWMS[cont];
  if(!Array.isArray(alerta.itemsWMS) || alerta.itemsWMS.length === 0) {
    return res.status(400).json({ ok:false, error:'la alerta no tiene snapshot de items del WMS' });
  }

  // FIX (rev Claude2+ChatGPT BLOQUEANTE v5.2.22): re-mapear el fisico por SKU
  // antes de aplicar los items nuevos. Esto evita el riesgo de desalineamiento
  // que detectaron ambos validadores: el fisico se alinea por índice posicional
  // con teorico.items[idx]. Si el WMS elimina/reordena líneas, los conteos
  // físicos quedan apuntando a SKUs equivocados.
  //
  // Estrategia: indexar el fisico viejo por el SKU que tenía, y reconstruir
  // un fisico nuevo donde cada conteo se asigna al índice del MISMO SKU en
  // la lista nueva. Los SKUs eliminados del WMS pierden su conteo (decisión
  // operativa: el WMS dice "esa línea ya no existe"). Los SKUs agregados
  // quedan con físico null (sin contar todavía, correcto).
  var prev = state.teorico[cont];
  var prevItems = Array.isArray(prev.items) ? prev.items : [];
  var prevFisico = Array.isArray(state.fisico[cont]) ? state.fisico[cont] : null;
  var newFisico = null;
  if(prevFisico) {
    // Indexar fisico viejo por SKU
    var fisicoBySku = {};
    prevItems.forEach(function(it, i){
      if(it && it.sku && prevFisico[i]) {
        fisicoBySku[it.sku] = prevFisico[i];
      }
    });
    // Construir nuevo fisico alineado con los items del WMS
    newFisico = alerta.itemsWMS.map(function(it){
      return (it && it.sku && fisicoBySku[it.sku]) ? fisicoBySku[it.sku] : null;
    });
  }

  // Aplicar items del WMS, preservar resto de propiedades (auditoría)
  state.teorico[cont] = Object.assign({}, prev, {
    items: alerta.itemsWMS,
    meta:  alerta.metaWMS || prev.meta || {}   // FIX (server v16): aplicar meta del WMS al sincronizar
  });
  // Aplicar fisico re-mapeado (solo si había fisico antes)
  if(newFisico) state.fisico[cont] = newFisico;

  // Contar cuántos conteos se mantuvieron / perdieron por SKUs eliminados
  var conteosMantenidos = 0, conteosPerdidos = 0;
  if(prevFisico && newFisico) {
    prevFisico.forEach(function(f){
      if(f && f.fisico !== undefined && f.fisico !== null && f.fisico !== '') {
        // Estaba contado antes. Ver si quedó en el nuevo.
        var sigueEnNuevo = newFisico.some(function(nf){
          return nf === f;
        });
        if(sigueEnNuevo) conteosMantenidos++;
        else conteosPerdidos++;
      }
    });
  }

  // Snapshot version + historial ANTES de mutar (addHistorial los mutará).
  // FIX (rev cruzada post-v18 borrador): el patrón original solo hacía version--
  // que revertía solo 1 de 2 incrementos y dejaba historial con la entrada
  // del intento fallido. Mismo fix que aplicamos a los otros 4 endpoints v18.
  var snapshotVersion   = state.version;
  var snapshotHistorial = state.historial.slice();

  // Eliminar la alerta
  delete state.alertasWMS[cont];

  addHistorial(usuario||'—', 'Sincronizó con WMS: ' + alerta.totalDiffs + ' cambios aplicados (conteos: ' + conteosMantenidos + ' mantenidos, ' + conteosPerdidos + ' perdidos)', cont);
  state.version++;

  // Persistir await (mismo patrón que clasificacion/set)
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'Sincronizar WMS save'
    );
    console.log('Sincronizar WMS save: persisted to Supabase ✓');
  } catch(saveErr) {
    console.log('Sincronizar WMS save FAILED:', saveErr.message);
    // Rollback completo: items + fisico + alerta + historial + version.
    state.teorico[cont] = prev;
    if(prevFisico) state.fisico[cont] = prevFisico;
    state.alertasWMS[cont] = alerta;
    state.historial = snapshotHistorial;
    state.version   = snapshotVersion;
    return res.status(500).json({
      ok: false,
      error: 'La sincronización NO quedó guardada. Reintentá. (' + saveErr.message + ')'
    });
  }

  res.json({
    ok: true,
    version: state.version,
    cont: cont,
    itemsAplicados: alerta.itemsWMS.length,
    cambiosAplicados: alerta.totalDiffs,
    conteosMantenidos: conteosMantenidos,
    conteosPerdidos:   conteosPerdidos
  });
});

// ── Chat ──────────────────────────────────────────────────────────────────
// ── Field locking ─────────────────────────────────────────────────────────
app.post('/api/lock', (req, res) => {
  cleanLocks();
  const { cont, idx, user } = req.body;
  const key = cont + ':' + idx;
  const existing = fieldLocks[key];
  if(existing && existing.user !== user && existing.expires > Date.now()) {
    return res.json({ ok:false, lockedBy: existing.user, since: existing.since });
  }
  fieldLocks[key] = { user, since: Date.now(), expires: Date.now() + LOCK_TIMEOUT };
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
  const { cont, idx, fisico, daniado, cobertura, calcExpr, usuario } = req.body;
  if(cont === undefined || idx === undefined) return res.status(400).json({ ok:false });
  if(!Array.isArray(state.fisico[cont])) state.fisico[cont] = [];
  const prev = state.fisico[cont][idx] || {};
  // Preserve null/undefined distinction — only fall back to prev when the
  // new value wasn't sent (undefined). Don't coerce null to 0.
  state.fisico[cont][idx] = {
    fisico:    fisico    !== undefined ? fisico    : prev.fisico,
    daniado:   daniado   !== undefined ? daniado   : prev.daniado,
    cobertura: cobertura !== undefined ? cobertura : (prev.cobertura || 'En revisión'),
    // FIX (mar 19-may-2026): el cliente nuevo ya no envía calcExpr (campo
    // Cálculo eliminado de la UI). Pero clientes con caché viejo todavía
    // pueden mandar calcExpr: ''. Tratamos null/undefined/'' como "no enviar"
    // y preservamos siempre el histórico previo. Solo aceptamos valores
    // no-vacíos (compatibilidad con clientes legacy que aún calculan).
    calcExpr:  (calcExpr != null && calcExpr !== '') ? calcExpr : ((prev && prev.calcExpr) || ''),
    quien:     usuario,
    ts:        new Date().toLocaleString('es'),
    lastUser:  usuario,
    lastAt:    Date.now()
  };
  state.version++;
  // Debounced (1.5s) — agrupa multiples tecleos seguidos del mismo o varios
  // usuarios. Antes hacía saveDailyState() sync por cada keystroke, lo que
  // serializaba el state completo (~7MB) en cada request y agotaba el heap.
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// ── Upload teórico ────────────────────────────────────────────────────────
// FIX CRÍTICO (mié 20-may-2026): el endpoint era sync con respuesta antes
// del save real. Eso causó pérdida silenciosa de contenedores ayer (HP26-0357,
// HP26-0591 entre otros). Dos problemas que se combinaban:
//   1) saveDailyState() retornaba Promesa pero NO se awaiteaba. La respuesta
//      {ok:true} se enviaba antes que Supabase confirmara. Si Render
//      reiniciaba en esos ms, el upload se perdía.
//   2) Si había un scheduleSave debounced pendiente de otro usuario, los
//      dos saves competían en paralelo. El debounced podía ganar y pisar
//      el upload con un state pre-upload en memoria.
// Fix: (a) cancelar saveTimer pendiente antes del upload, (b) hacer await
// del save crítico, (c) responder con error si el save falla.
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const usuario = req.body.usuario || '—';

    // FIX (mié 28-may-2026, server v18): rollback ante fallo de persistencia.
    // mergeSheet muta state.teorico, state.fisico y state.alertasWMS reemplazando
    // entradas por contenedor. Snapshot somero de esos 3 objetos preserva las
    // referencias previas; al revertir, los contenedores modificados vuelven a
    // apuntar a sus objetos originales (los no tocados nunca cambiaron).
    // Costo: tres spreads superficiales (~ms), no deep clone.
    //
    // LIMITACIÓN CONOCIDA del snapshot somero (deuda en backlog):
    // mergeSheet también muta PROPIEDADES de objetos existentes — concretamente
    // hace `state.teorico[cont].cdgValidado = true` para contenedores que ya
    // tienen `fromCDG=true` (línea ~1352). Como el snapshot guarda referencias,
    // esa mutación NO se revierte al hacer rollback: el contenedor preserva su
    // objeto original, pero ese objeto ahora tiene `cdgValidado=true` adentro.
    // Severidad baja:
    //   - el flag no destruye datos (solo marca como validado un contenedor
    //     que ya tenía fromCDG=true; semánticamente coherente)
    //   - el próximo upload exitoso lo normaliza igual
    //   - el operario no percibe nada distinto en la UI
    // Fix completo en backlog: deep clone selectivo de los fromCDG, o refactor
    // de mergeSheet para que no mute propiedades (que devuelva los cambios a
    // aplicar y el caller los aplique). Hoy NO se hace por costo/riesgo vs
    // beneficio marginal.
    var snapshotTeorico    = Object.assign({}, state.teorico);
    var snapshotFisico     = Object.assign({}, state.fisico);
    var snapshotAlertasWMS = Object.assign({}, state.alertasWMS || {});
    var snapshotHistorial  = state.historial.slice();
    var snapshotVersion    = state.version;

    let loaded = [];
    wb.SheetNames.forEach(sn => {
      const nl = sn.toLowerCase();
      const type = nl.includes('traslado') ? 'Traslados'
                 : nl.includes('embarque') ? 'Embarques'
                 : null;
      if(!type) return;
      const rows  = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, type);
      if(count > 0) {
        loaded.push(sn+'('+count+')');
        addHistorial(usuario, 'Teórico cargado', type+' — '+count+' contenedores');
      }
    });
    if(loaded.length === 0) {
      const rows  = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, 'General');
      if(count > 0) loaded.push(wb.SheetNames[0]+'('+count+')');
      addHistorial(usuario, 'Teórico cargado', loaded.join());
    }
    state.version++;

    // Cancelar cualquier debounced save pendiente: si existe, contendría
    // una copia del state PRE-upload y al ejecutarse pisaría el upload.
    if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }

    // Await: NO responder éxito hasta que Supabase confirme.
    // Usar buildDailyStatePayload directo (no saveDailyState que swallow errores).
    // FIX (rev Claude2): timeout 15s. FIX (rev ChatGPT): mensaje operativo.
    try {
      await withTimeout(
        dbSet('daily_state', buildDailyStatePayload()),
        15000,
        'Upload save'
      );
      console.log('Upload save: persisted to Supabase ✓');
    } catch(saveErr) {
      console.log('Upload save FAILED:', saveErr.message);
      // Rollback: restaurar los 3 objetos shallow + historial + version.
      // Esto deshace el mergeSheet completo (todos los contenedores tocados
      // vuelven a apuntar a sus objetos previos).
      state.teorico    = snapshotTeorico;
      state.fisico     = snapshotFisico;
      state.alertasWMS = snapshotAlertasWMS;
      state.historial  = snapshotHistorial;
      state.version    = snapshotVersion;
      return res.status(500).json({
        ok: false,
        error: 'El archivo NO se guardó (memoria restaurada). Volvé a subirlo. (' + saveErr.message + ')'
      });
    }

    res.json({ ok:true, loaded });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── Fotos: upload a Supabase Storage ──────────────────────────────────────
// Endpoint que recibe una foto en base64 desde el cliente y la sube al
// bucket 'app-fotos' de Supabase Storage. Devuelve la URL pública para
// que el cliente la guarde en el campo `foto` del hallazgo/CDG en vez de
// almacenar el base64 completo dentro del state.
//
// FIX (mié 20-may-2026): migración fotos a Storage para resolver
// PayloadTooLargeError + state inflado por base64. Decisiones del diseño:
//   - bucket público (URLs largas+random como protección por oscuridad)
//   - validación de tamaño max 5MB (cliente debe comprimir antes)
//   - validación de mime type permitidos
//   - timeout 15s (subida puede ser lenta en red mala)
//   - path por kind: 'hallazgo-{id}.jpg', 'cdg-gral-{id}.jpg', 'cdg-sku-{id}.jpg'
//   - si el path ya existe, sobrescribe (Prefer: x-upsert=true)
function uploadToStorage(buffer, path, contentType) {
  return new Promise((resolve, reject) => {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return reject(new Error('Supabase no configurado'));
    }
    var url = new URL(`${SUPABASE_URL}/storage/v1/object/app-fotos/${path}`);
    var opts = {
      hostname: url.hostname,
      path:     url.pathname,
      method:   'POST',
      headers: {
        'apikey':        SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type':  contentType,
        'Content-Length': buffer.length,
        'x-upsert':      'true', // sobrescribir si existe
        'Cache-Control': '3600'
      }
    };
    var req = https.request(opts, function(res){
      var raw = '';
      res.on('data', function(c){ raw += c; });
      res.on('end', function(){
        if(res.statusCode >= 400) {
          var errMsg = 'Storage HTTP ' + res.statusCode;
          try {
            var b = raw ? JSON.parse(raw) : null;
            if(b && b.message) errMsg += ': ' + b.message;
            else if(b && b.error) errMsg += ': ' + b.error;
            else if(raw) errMsg += ': ' + raw.slice(0, 200);
          } catch(e) {
            if(raw) errMsg += ': ' + raw.slice(0, 200);
          }
          return reject(new Error(errMsg));
        }
        // Construir URL pública. Para buckets públicos:
        // https://{project}.supabase.co/storage/v1/object/public/app-fotos/{path}
        var publicUrl = `${SUPABASE_URL}/storage/v1/object/public/app-fotos/${path}`;
        resolve({ ok:true, url: publicUrl, key: path });
      });
    });
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}

const ALLOWED_PHOTO_MIME = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp'];
const MAX_PHOTO_BYTES = 5 * 1024 * 1024; // 5MB

app.post('/api/foto/upload', async (req, res) => {
  try {
    const { foto, kind, refId, usuario } = req.body;
    if(!foto || !kind || !refId) {
      return res.status(400).json({ ok:false, error:'faltan campos: foto, kind, refId' });
    }
    // kind permitido: hallazgo, cdg-gral, cdg-sku
    if(!['hallazgo', 'cdg-gral', 'cdg-sku'].includes(kind)) {
      return res.status(400).json({ ok:false, error:'kind inválido (debe ser hallazgo|cdg-gral|cdg-sku)' });
    }
    // refId sanitizado: solo alfanumérico, guiones, slashes, puntos.
    if(!/^[A-Za-z0-9_\-./]+$/.test(String(refId))) {
      return res.status(400).json({ ok:false, error:'refId inválido (solo A-Z 0-9 _ - . /)' });
    }

    // El foto viene como data URL: "data:image/jpeg;base64,XXX..."
    var match = String(foto).match(/^data:([^;]+);base64,(.+)$/);
    if(!match) {
      return res.status(400).json({ ok:false, error:'foto debe ser data URL base64' });
    }
    var mimeType = match[1].toLowerCase();
    var base64Data = match[2];

    if(!ALLOWED_PHOTO_MIME.includes(mimeType)) {
      return res.status(400).json({ ok:false, error:'mime type no permitido: ' + mimeType });
    }

    // Decodificar base64 a buffer
    var buffer;
    try {
      buffer = Buffer.from(base64Data, 'base64');
    } catch(e) {
      return res.status(400).json({ ok:false, error:'base64 inválido' });
    }

    if(buffer.length > MAX_PHOTO_BYTES) {
      return res.status(413).json({
        ok: false,
        error: 'foto muy grande (' + Math.round(buffer.length/1024) + ' KB). Máximo: ' + Math.round(MAX_PHOTO_BYTES/1024) + ' KB. Comprimí antes de subir.'
      });
    }

    // Construir path en el bucket. Extensión a partir del mime.
    var ext = (mimeType.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
    var timestamp = Date.now();
    var safeRefId = String(refId).replace(/[^A-Za-z0-9_\-]/g, '_');
    var path = `${kind}/${safeRefId}_${timestamp}.${ext}`;

    // Subir a Storage con timeout
    var result;
    try {
      result = await withTimeout(
        uploadToStorage(buffer, path, mimeType),
        15000,
        'Foto upload'
      );
      console.log('Foto upload OK:', path, '(' + Math.round(buffer.length/1024) + ' KB)');
    } catch(uploadErr) {
      console.log('Foto upload FAILED:', uploadErr.message);
      return res.status(500).json({
        ok: false,
        error: 'No se pudo subir la foto al Storage. Reintentá. (' + uploadErr.message + ')'
      });
    }

    res.json({
      ok: true,
      url: result.url,
      key: result.key,
      bytes: buffer.length
    });
  } catch(e) {
    console.log('Foto upload exception:', e.message);
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── Helpers ───────────────────────────────────────────────────────────────
function norm(s) {
  return String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function findCol(hdr, terms) {
  for(const t of terms) { const i = hdr.findIndex(h => h === norm(t)); if(i >= 0) return i; }
  for(const t of terms) { const i = hdr.findIndex(h => h.includes(norm(t))); if(i >= 0) return i; }
  return -1;
}

function mergeSheet(rows, type) {
  if(!rows || rows.length < 2) return 0;
  // Find actual header row (contains 'sku' or 'numero')
  let hdrRowIdx = 0;
  for(let ri = 0; ri < Math.min(rows.length, 10); ri++) {
    const r = rows[ri].map(h => norm(h));
    if(r.some(h => h === 'sku' || h.includes('sku')) ||
       r.some(h => h === 'numero' || h === 'número')) {
      hdrRowIdx = ri;
      break;
    }
  }
  const hdr      = rows[hdrRowIdx].map(h => norm(h));
  const dataRows = rows.slice(hdrRowIdx + 1);
  const colCont = findCol(hdr, ['numero','número']);
  const colSku  = findCol(hdr, ['sku']);
  const colQty  = findCol(hdr, ['cant.','cant','cantidad']);
  let colDesc = -1;
  for(let i = colSku + 1; i < hdr.length; i++) {
    if(hdr[i].includes('nombre') || hdr[i].includes('descripcion')) { colDesc = i; break; }
  }
  if(colDesc < 0) colDesc = findCol(hdr, ['nombre','descripcion']);
  if(colCont < 0 || colSku < 0 || colDesc < 0 || colQty < 0) return 0;

  const g = (row, i) => i >= 0 ? row[i] : '';
  // FIX (sáb 23-may-2026): el WMS de Embarques trae DOS columnas relevantes para
  // proveedor: 'Proveedor' (código, ej. HPP001199) y 'Nombre' (razón social, ej.
  // LONGTAI TRADING FZCO). Antes solo se capturaba el código en codProv, así que
  // raw.nomProv quedaba undefined y getOrigenDesc (export) recibía el código →
  // 'Descripción de origen' salía como el código en vez de Truper China/México.
  // Capturamos el 'Nombre' que aparece JUSTO DESPUÉS de la columna Proveedor,
  // para no confundirlo con el 'Nombre' de descripción del SKU (que va después de SKU).
  const cCodProvIdx = findCol(hdr, ['proveedor','código de proveedor','codigo de proveedor']);
  let cNomProvIdx = -1;
  if(cCodProvIdx >= 0) {
    for(let i = cCodProvIdx + 1; i < hdr.length; i++) {
      if(hdr[i].includes('nombre')) { cNomProvIdx = i; break; }
      // No buscar más allá del SKU: el 'Nombre' de descripción va después del SKU
      if(hdr[i] === 'sku' || hdr[i].includes('sku')) break;
    }
  }
  const ex = {
    cFecha:    findCol(hdr, ['fecha']),
    cCodProv:  cCodProvIdx,
    cNomProv:  cNomProvIdx,
    cLineas:   findCol(hdr, ['lineas','líneas']),
    cStatus:   findCol(hdr, ['status']),
    cOC:       findCol(hdr, ['orden de compra','# orden']),
    cIngr:     findCol(hdr, ['ingresado']),
    cColoc:    findCol(hdr, ['colocado']),
    cFalt:     findCol(hdr, ['faltantes']),
    cSobr:     findCol(hdr, ['sobrantes']),
    cDan:      findCol(hdr, ['dañado','danado']),
    cOrigen:   findCol(hdr, ['origen']),
    cIngreso:  findCol(hdr, ['# ingreso','ingreso']),
    cDocSap:   findCol(hdr, ['doc. sap','doc sap']),
    cTipo:     findCol(hdr, ['tipo']),
    cUnidad:   findCol(hdr, ['unidad']),
    cUnidades: findCol(hdr, ['unidades']),
    cDestino:  findCol(hdr, ['destino'])
  };
  const newConts = {};
  // FIX (dom 25-may-2026, server v16): OPTIMIZACIÓN DE TAMAÑO DEL STATE.
  // Antes, cada item guardaba 18 campos en raw. Medido contra el WMS real,
  // 10 de esos campos (fecha, codProv, nomProv, lineas, status, origen,
  // ingreso, docSap, tipo, destino) son CONSTANTES dentro de un mismo
  // contenedor — se repetían idénticos en cada una de sus ~150 líneas.
  // Ahora esos campos van UNA sola vez a newMeta[cont] (nivel contenedor),
  // y raw queda solo con los que VARÍAN por línea (oc, ingresado, colocado,
  // faltantes, sobrantes, daniado, unidad, unidades). Reducción ~70% del teorico.
  // Compatibilidad: el cliente lee raw heredando de meta (Object.assign), así
  // que los exports siguen funcionando y los contenedores VIEJOS (con raw
  // completo y sin meta) también, porque su raw viejo gana sobre el meta ausente.
  const newMeta = {};
  dataRows
    .filter(r => r && r.some(c => String(c).trim() !== ''))
    .forEach(row => {
      const cont = String(row[colCont]||'').trim();
      if(!cont) return;
      // FIX (mar 26-may-2026, server v17): LISTA BLANCA de nombre de contenedor.
      // El export de Power BI trae texto al pie ("Filtros aplicados: ... Embarque
      // es HP26-XXXX Aplica es 1", "Total general", etc.) que el parser leía como
      // contenedores basura. Ahora SOLO se acepta una fila si su nombre tiene el
      // formato real: empieza con H o U, dígitos, guion, dígitos, con sufijo
      // opcional (ej. /CS). HP26-0540, H274-0123, U25-161410, U147-0089.
      // Cualquier otra cosa (texto largo, notas, totales) se ignora automáticamente.
      if(!/^[HU][A-Z0-9]*-[0-9]+(\/[A-Z0-9]+)?$/i.test(cont)) return;
      // FIX (mar 26-may-2026, server v17): los contenedores terminados en /CS
      // existen en la base del WMS pero el equipo NO los usa. Antes había que
      // borrarlos a mano tras cada carga del maestro porque se volvían a meter.
      // Ahora se ignoran automáticamente al cargar.
      if(/\/CS$/i.test(cont)) return;
      if(!newConts[cont]) newConts[cont] = [];
      // Meta a nivel contenedor: se captura de la primera fila (campos constantes).
      if(!newMeta[cont]) {
        newMeta[cont] = {
          fecha:    g(row, ex.cFecha),
          codProv:  g(row, ex.cCodProv),
          nomProv:  g(row, ex.cNomProv),
          lineas:   g(row, ex.cLineas),
          status:   g(row, ex.cStatus),
          origen:   g(row, ex.cOrigen),
          ingreso:  g(row, ex.cIngreso),
          docSap:   g(row, ex.cDocSap),
          tipo:     g(row, ex.cTipo),
          destino:  g(row, ex.cDestino)
        };
      }
      newConts[cont].push({
        sku:  String(row[colSku] ||'').trim(),
        desc: String(row[colDesc]||'').trim(),
        qty:  parseFloat(String(row[colQty]).replace(',','.')) || 0,
        raw: {
          // Solo campos que VARÍAN por línea:
          oc:        g(row, ex.cOC),
          ingresado: g(row, ex.cIngr),
          colocado:  g(row, ex.cColoc),
          faltantes: g(row, ex.cFalt),
          sobrantes: g(row, ex.cSobr),
          daniado:   g(row, ex.cDan),
          unidad:    g(row, ex.cUnidad),
          unidades:  g(row, ex.cUnidades)
        }
      });
    });

  // FIX (dom 24-may-2026 PM, server v15): FREEZE de contenedores ya contados.
  // Contexto: el equipo va a subir el Excel del Power BI varias veces al día.
  // El WMS puede cambiar cantidades o eliminar líneas de un contenedor que
  // YA empezamos a contar. Si dejáramos que el upload sobreescriba, el equipo
  // perdería referencia (físico contado vs teórico modificado).
  //
  // Regla: si state.fisico[cont] tiene AL MENOS 1 línea con valor (contada),
  // el contenedor se considera "frozen" y NO se actualiza su teorico desde el
  // upload. EN CAMBIO, se registra una alerta en state.alertasWMS para que el
  // supervisor decida si "descongelar" y sincronizar manualmente.
  //
  // Helper: detectar si un contenedor ya tiene físico contado
  function hasPhysicalCounted(contKey) {
    var f = state.fisico && state.fisico[contKey];
    if(!Array.isArray(f)) return false;
    for(var i = 0; i < f.length; i++) {
      var x = f[i];
      if(x && x.fisico !== undefined && x.fisico !== null && x.fisico !== '') return true;
    }
    return false;
  }

  // Helper: comparar items prev vs new para detectar diferencias específicas
  function compareItems(prevItems, newItems) {
    var prev = Array.isArray(prevItems) ? prevItems : [];
    var nw = Array.isArray(newItems) ? newItems : [];
    var diffs = [];
    // Index por SKU para comparación
    var prevBySku = {}, newBySku = {};
    prev.forEach(function(it){ if(it && it.sku) prevBySku[it.sku] = it; });
    nw.forEach(function(it){ if(it && it.sku) newBySku[it.sku] = it; });
    // SKUs eliminados (en prev pero no en new)
    Object.keys(prevBySku).forEach(function(sku){
      if(!newBySku[sku]) diffs.push({ tipo:'eliminado', sku:sku, qty:prevBySku[sku].qty });
    });
    // SKUs nuevos (en new pero no en prev)
    Object.keys(newBySku).forEach(function(sku){
      if(!prevBySku[sku]) diffs.push({ tipo:'agregado', sku:sku, qty:newBySku[sku].qty });
    });
    // SKUs con cantidad cambiada
    Object.keys(prevBySku).forEach(function(sku){
      if(newBySku[sku] && Number(prevBySku[sku].qty) !== Number(newBySku[sku].qty)) {
        diffs.push({ tipo:'qty_cambio', sku:sku, qtyAntes:prevBySku[sku].qty, qtyNueva:newBySku[sku].qty });
      }
    });
    return diffs;
  }

  // Inicializar el contenedor de alertas si no existe
  if(!state.alertasWMS) state.alertasWMS = {};

  var congeladosCount = 0;
  var congeladosConCambio = 0;

  Object.keys(newConts).forEach(cont => {
    // Don't overwrite CDG-validated containers from teorico upload
    if(state.teorico[cont] && state.teorico[cont].fromCDG) {
      state.teorico[cont].cdgValidado = true;
      return;
    }

    var prev = state.teorico[cont] || {};
    var fechaCargaPrev = prev.fechaCarga || null;
    var hoy = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Guatemala' });

    // ── FREEZE check (server v15) ───────────────────────────────
    // Si el contenedor ya tiene físico contado, NO actualizar items.
    // Pero SÍ comparar y registrar alerta si hay cambios.
    var isFrozen = prev.items && hasPhysicalCounted(cont);

    if(isFrozen) {
      congeladosCount++;
      var diffs = compareItems(prev.items, newConts[cont]);
      if(diffs.length > 0) {
        congeladosConCambio++;
        // Guardar alerta CON los items nuevos del WMS. Esto permite al
        // supervisor "Sincronizar con WMS" después, aplicando estos items.
        // NO acumular alertas viejas: cada upload reemplaza la alerta con
        // el snapshot más reciente del WMS, para no inflar el state.
        state.alertasWMS[cont] = {
          detectadoEn: hoy,
          itemsAntes:  prev.items.length,
          itemsNuevos: newConts[cont].length,
          diffs:       diffs.slice(0, 30),  // límite defensivo: top 30 diffs
          totalDiffs:  diffs.length,
          // Snapshot completo de items del WMS para "Sincronizar" después
          // (necesario porque sin esto el supervisor no podría aplicar el WMS
          // ya que el upload solo se guarda si NO está congelado)
          itemsWMS:    newConts[cont],
          metaWMS:     newMeta[cont] || {}   // FIX (server v16): meta para aplicar al sincronizar
        };
      } else {
        // Sin diferencias: limpiar cualquier alerta vieja para este contenedor
        if(state.alertasWMS[cont]) delete state.alertasWMS[cont];
      }
      // Solo actualizar metadatos no-críticos del raw del primer item
      // (status del WMS, etc.) — pero NO items
      // Por simplicidad y seguridad: NO tocamos nada cuando está frozen.
      // El supervisor puede sincronizar manualmente vía /api/wms/sincronizar.
      return;
    }

    // Contenedor SIN físico contado: comportamiento normal (preservar auditoría)
    state.teorico[cont] = Object.assign({}, prev, {
      items: newConts[cont],
      meta:  newMeta[cont] || prev.meta || {},   // FIX (server v16): meta a nivel contenedor
      type,
      fechaCarga: fechaCargaPrev || hoy
    });
    // Si había una alerta vieja, limpiarla (el upload nuevo se aplicó OK)
    if(state.alertasWMS && state.alertasWMS[cont]) delete state.alertasWMS[cont];
    // Preserve existing fisico data — never overwrite conteo work
    if(!state.fisico[cont]) state.fisico[cont] = null;
  });

  console.log('mergeSheet:', Object.keys(newConts).length, 'contenedores procesados,',
              congeladosCount, 'congelados (' + congeladosConCambio + ' con cambios del WMS)');
  return Object.keys(newConts).length;
}

// ═══════════════════════════════════════════════════════════════════════════
// CDG v2 — Módulo colaborativo (server v19, 29-may-2026)
// Adaptado sáb 30-may-2026: corregido comentario service_role→anon,
// rutas sku-catalog/* movidas antes de /:id (bug de routing Express).
//
// Arquitectura:
//   cdg_meta_{id}  → registro en app_state con metadata, usuarios, bitácora
//   cdg_lineas     → tabla Supabase independiente, una fila por línea
//
// Esto evita reescribir el array completo de líneas en cada operación.
// Dos usuarios agregando líneas simultáneamente hacen INSERTs independientes,
// sin conflicto de versión.
//
// IMPORTANTE: estos endpoints NO tocan state (el monolito Hamilton).
// NO aparecen en buildDailyStatePayload() ni en publicState().
// La red de seguridad saveDirectToSupabase del cliente sigue operando
// sobre daily_state como antes — no se ve afectada.
//
// Backup del módulo CDG v1 (legacy): los endpoints /api/cdg/* originales
// se conservan sin cambios más abajo para rollback en caso necesario.
// ═══════════════════════════════════════════════════════════════════════════

// ── Helpers CDG v2 ─────────────────────────────────────────────────────────

// Genera un UUID v4 simple (sin dependencias externas)
// FIX (lun 1-jun-2026, v19): lock en memoria por licencia para operaciones CDG v2.
// Resuelve la race condition donde un POST /linea en vuelo puede insertar después
// de que el cierre ya leyó las líneas para Hamilton.
// Estructura: { [licenciaId]: { activeWrites: 0, closing: false } }
var cdgLocks = {};
function cdgLockAcquire(licenciaId) {
  if(!cdgLocks[licenciaId]) cdgLocks[licenciaId] = { activeWrites: 0, closing: false };
  if(cdgLocks[licenciaId].closing) return false; // rechazar — cierre en progreso
  cdgLocks[licenciaId].activeWrites++;
  return true;
}
function cdgLockRelease(licenciaId) {
  var l = cdgLocks[licenciaId];
  if(l && l.activeWrites > 0) l.activeWrites--;
}
function cdgLockStartClosing(licenciaId) {
  if(!cdgLocks[licenciaId]) cdgLocks[licenciaId] = { activeWrites: 0, closing: false };
  cdgLocks[licenciaId].closing = true;
}
function cdgLockClear(licenciaId) {
  delete cdgLocks[licenciaId];
}
async function cdgLockWaitDrain(licenciaId, timeoutMs) {
  var l = cdgLocks[licenciaId];
  if(!l) return;
  var start = Date.now();
  while(l.activeWrites > 0 && Date.now() - start < timeoutMs) {
    await new Promise(function(r){ setTimeout(r, 10); });
  }
  if(l.activeWrites > 0) {
    // FIX: lanzar error en lugar de proceder — el cierre no puede correr con
    // writes activos. El caller limpia el lock y responde error al cliente.
    throw new Error('Timeout esperando que terminen ' + l.activeWrites + ' escritura(s) activa(s). Reintentá en unos segundos.');
  }
}

function uuidv4() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

// Lee la metadata de una licencia CDG desde app_state
async function cdgGetMeta(licenciaId) {
  // FIX (lun 1-jun-2026, v19): normalizar id con trim().toUpperCase() para evitar
  // claves duplicadas por espacios o mayúsculas (iPad autocapitaliza).
  return await dbGet('cdg_meta_' + cdgNormId(licenciaId));
}

// Guarda la metadata de una licencia CDG en app_state
async function cdgSaveMeta(licenciaId, meta) {
  await dbSet('cdg_meta_' + cdgNormId(licenciaId), meta);
}

// FIX (lun 1-jun-2026, v19): helper central de normalización.
// Todos los endpoints v2 usan cdgNormId() al extraer req.params.id y al
// construir licencia_id en filas de cdg_lineas. Garantiza que meta y líneas
// usen siempre la misma clave canónica (sin espacios, mayúsculas).
function cdgNormId(id) {
  return String(id || '').trim().toUpperCase();
}

// Lee las líneas de una licencia desde cdg_lineas (tabla Supabase).
// Carga inicial (sin desde): solo líneas NO eliminadas.
// Delta (con desde): todas las líneas modificadas después de ese momento,
// INCLUYENDO las eliminadas — necesario para que el polling propague deletes
// a otras tablets. El cliente descarta las que tengan eliminada=true.
async function cdgGetLineas(licenciaId, desde) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return [];
  var nId   = cdgNormId(licenciaId);
  var query = '?licencia_id=eq.' + encodeURIComponent(nId)
            + '&order=ts_creacion.asc';
  if(desde) {
    // Delta: incluir eliminadas para que las tablets las descarten localmente
    // FIX (lun 1-jun-2026, v19): usar gte. (mayor o igual) en lugar de gt.
    // Con gt. estricto, dos líneas con el mismo ts_modif pierden la del borde
    // del cursor — nunca llega en el delta. gte. la reenvía, y el merge del
    // cliente la descarta como duplicado (findIndex por id). Sin pérdida de datos.
    query += '&ts_modif=gte.' + encodeURIComponent(desde);
  } else {
    // Carga completa: solo activas
    query += '&eliminada=eq.false';
  }
  var rows = await supabase('GET', 'cdg_lineas', null, query);
  return Array.isArray(rows) ? rows : [];
}

// Inserta una línea nueva en cdg_lineas
async function cdgInsertLinea(linea) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var rows = await supabase('POST', 'cdg_lineas', linea, '');
  return Array.isArray(rows) ? rows[0] : rows;
}

// Actualiza una línea existente (solo campos permitidos)
async function cdgUpdateLinea(lineaId, patch, autorEsperado) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  patch.ts_modif = new Date().toISOString();
  var query = '?id=eq.' + encodeURIComponent(lineaId)
            + '&autor=eq.' + encodeURIComponent(autorEsperado)
            + '&eliminada=eq.false';
  var rows = await supabase('PATCH', 'cdg_lineas', patch, query);
  return Array.isArray(rows) ? rows[0] : rows;
}

// Soft-delete de una línea (solo el autor puede borrar la suya)
async function cdgSoftDeleteLinea(lineaId, autorEsperado) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var patch = { eliminada: true, ts_modif: new Date().toISOString() };
  var query = '?id=eq.' + encodeURIComponent(lineaId)
            + '&autor=eq.' + encodeURIComponent(autorEsperado)
            + '&eliminada=eq.false';
  var rows = await supabase('PATCH', 'cdg_lineas', patch, query);
  return Array.isArray(rows) ? rows[0] : rows;
}

// Actualiza tsUltimaModifLineas en la metadata cuando cambian las líneas
// FIX (sáb 30-may-2026, server v19): re-leer meta fresco desde DB antes de bump.
// SIN este fix: si usuario A leyó meta con estado='activo' y luego usuario B
// cerró la licencia (estado='cerrado'), el write de A sobreescribe con su meta
// stale y reabre la licencia — riesgo operativo real.
// CON el fix: siempre escribimos sobre la base más reciente de Supabase,
// preservando estado/finalizador/bitacora actuales. Solo propagamos lastActivity
// del usuario local (campo seguro de pisar — a lo sumo queda unos segundos atrás).
// deltaTotalLineas: +1 al insertar, -1 al eliminar, 0 o undefined al editar.
// Sin deltaTotalLineas el contador no se actualiza (no lo podemos inferir).
// Costo: 1 GET extra por operación de línea. Aceptable para equipos de 4-6.
async function cdgBumpLineasTs(licenciaId, metaLocal, deltaTotalLineas) {
  var metaFresh = await cdgGetMeta(licenciaId);
  var base = metaFresh || metaLocal; // fallback si la licencia fue borrada
  base.tsUltimaModifLineas = new Date().toISOString();
  base.version = (base.version || 0) + 1;
  // Aplicar delta de totalLineas sobre el valor FRESCO (no el stale de metaLocal)
  if(deltaTotalLineas !== undefined && deltaTotalLineas !== 0) {
    base.totalLineas = Math.max(0, (base.totalLineas || 0) + deltaTotalLineas);
  }
  // Propagar solo lastActivity del usuario que operó (nunca estado de licencia)
  var usuarios = metaLocal.usuarios || {};
  Object.keys(usuarios).forEach(function(u) {
    if(!base.usuarios) base.usuarios = {};
    if(!base.usuarios[u]) base.usuarios[u] = { estado: 'activo', lastActivity: null };
    var localTs = (usuarios[u] || {}).lastActivity;
    if(localTs && localTs > (base.usuarios[u].lastActivity || '')) {
      base.usuarios[u].lastActivity = localTs;
    }
  });
  await cdgSaveMeta(licenciaId, base);
  // Retornar estadoActual para que el caller detecte si la licencia fue cerrada
  // mientras la operación estaba en vuelo y pueda responder 409 al cliente.
  return { estadoActual: base.estado };
}

// Agrega una entrada a la bitácora de la licencia
function cdgBitacora(meta, usuario, accion, detalle) {
  if(!meta.bitacora) meta.bitacora = [];
  meta.bitacora.push({
    ts:      new Date().toISOString(),
    usuario: usuario || '—',
    accion:  accion,
    detalle: detalle || ''
  });
  // Límite defensivo: máx 200 entradas en bitácora
  if(meta.bitacora.length > 200) meta.bitacora = meta.bitacora.slice(-200);
}

// Verifica si una licencia acepta nuevas líneas
function cdgAceptaLineas(meta) {
  return meta && meta.estado === 'activo';
}

// ── GET /api/cdg/v2/listar ─────────────────────────────────────────────────
// Lista licencias CDG v2.
// Sin parámetros: activas, pausadas y cerradas en últimas 48h (para captura/manifiesto).
// ?historico=true: todas las licencias cerradas sin límite de fecha (para Reportes).
//   En modo histórico, enriquece cada licencia con totalUnidades y costoTotal
//   calculados desde cdg_lineas vigentes (eliminada=false).
// FIX (lun 1-jun-2026, server v20): modo histórico para Reportes/Historial CDG.
app.get('/api/cdg/v2/listar', async (req, res) => {
  try {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return res.json({ ok: true, licencias: [] });
    }

    var historico = req.query.historico === 'true';

    // Busca claves que empiecen con "cdg_meta_" en app_state
    var query = '?key=like.cdg_meta_*&order=key.asc&select=key,value';
    var rows = await supabase('GET', 'app_state', null, query);
    if(!Array.isArray(rows)) return res.json({ ok: true, licencias: [] });

    var licencias = [];
    var hace48h = new Date(Date.now() - 48*60*60*1000).toISOString();

    rows.forEach(function(row) {
      try {
        var meta = typeof row.value === 'string' ? JSON.parse(row.value) : row.value;
        if(!meta) return;
        var incluir;
        if(historico) {
          // Modo histórico: incluir todas (activas, pausadas, cerradas sin límite de fecha)
          incluir = true;
        } else {
          // Modo operativo: activas, pausadas, cerradas recientes (últimas 48h)
          incluir = meta.estado === 'activo' || meta.estado === 'pausado' ||
            (meta.estado === 'cerrado' && meta.fechaCierre && meta.fechaCierre > hace48h);
        }
        if(incluir) {
          licencias.push({
            id:            meta.id,
            tipo:          meta.tipo,
            estado:        meta.estado,
            creadoPor:     meta.creadoPor,
            finalizador:   meta.finalizador,
            fechaCreacion: meta.fechaCreacion,
            fechaCierre:   meta.fechaCierre || null,
            totalLineas:   meta.totalLineas || 0,
            usuarios:      Object.keys(meta.usuarios || {}),
            // Agregados: se llenan en modo histórico; null en modo operativo
            totalUnidades: null,
            costoTotal:    null
          });
        }
      } catch(e) { /* skip malformed */ }
    });

    // Modo histórico: enriquecer con agregados desde cdg_lineas
    if(historico && licencias.length > 0) {
      try {
        // Una sola query: todas las líneas vigentes de todas las licencias encontradas
        var ids = licencias.map(function(l){ return encodeURIComponent(l.id); }).join(',');
        var lineasQuery = '?licencia_id=in.(' + ids + ')&eliminada=eq.false&select=licencia_id,cantidad,costo_unit';
        var lineasRows = await supabase('GET', 'cdg_lineas', null, lineasQuery);

        if(Array.isArray(lineasRows)) {
          // Acumular por licencia_id
          var agregados = {};
          lineasRows.forEach(function(l){
            var lid = l.licencia_id;
            if(!agregados[lid]) agregados[lid] = { unidades: 0, costo: 0 };
            var qty  = Number(l.cantidad) || 0;
            var cost = (l.costo_unit !== null && l.costo_unit !== undefined) ? Number(l.costo_unit) : 0;
            agregados[lid].unidades += qty;
            agregados[lid].costo    += qty * (isNaN(cost) ? 0 : cost);
          });
          // Pegar los agregados en cada licencia
          licencias.forEach(function(l){
            var ag = agregados[l.id];
            if(ag) {
              l.totalUnidades = ag.unidades;
              l.costoTotal    = ag.costo;
            } else {
              l.totalUnidades = 0;
              l.costoTotal    = 0;
            }
          });
        }
      } catch(aggErr) {
        // Si falla el enriquecimiento, devolver licencias sin agregados (no bloqueante)
        console.log('CDG v2 listar agregados WARN:', aggErr.message);
      }
    }

    res.json({ ok: true, licencias: licencias });
  } catch(e) {
    console.log('CDG v2 listar error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── POST /api/cdg/v2/crear ─────────────────────────────────────────────────
// Crea una nueva licencia CDG.
// Body: { id, tipo, usuario }
// id: correlativo U25-XXXX. Si no se pasa, se genera automáticamente.
app.post('/api/cdg/v2/crear', async (req, res) => {
  var { id, tipo, usuario } = req.body;
  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });

  // Generar ID si no se pasó
  if(!id) {
    var fecha = new Date();
    var mm = String(fecha.getMonth() + 1).padStart(2, '0');
    var dd = String(fecha.getDate()).padStart(2, '0');
    var rand = Math.floor(Math.random() * 9000) + 1000;
    id = 'U25-' + mm + dd + rand;
  }
  // FIX (lun 1-jun-2026, v19): normalizar id con trim().toUpperCase() para evitar
  // duplicados por espacios o mayúsculas (iPad autocapitaliza la primera letra).
  id = cdgNormId(id); // usar helper central

  // Verificar que no exista ya
  try {
    var existing = await cdgGetMeta(id);
    if(existing) {
      return res.status(409).json({ ok: false, error: 'Ya existe una licencia con ese ID. Usá /api/cdg/v2/' + id + ' para unirte.' });
    }
  } catch(e) { /* no existe, continuar */ }

  var now = new Date().toISOString();
  var meta = {
    id:                   id,
    tipo:                 tipo || 'CDG',
    estado:               'activo',
    creadoPor:            usuario,
    finalizador:          usuario,
    fechaCreacion:        now,
    fechaCierre:          null,
    fotosEncabezado:      [],
    usuarios: {},
    totalLineas:          0,
    tsUltimaModifLineas:  now,
    bitacora:             [],
    version:              1
  };
  meta.usuarios[usuario] = {
    estado:       'activo',
    lastActivity: now
  };
  cdgBitacora(meta, usuario, 'licencia_creada', tipo || 'CDG');

  try {
    await withTimeout(cdgSaveMeta(id, meta), 15000, 'CDG crear save');
    console.log('CDG v2 crear: licencia', id, 'por', usuario);
    res.json({ ok: true, licencia: meta });
  } catch(e) {
    console.log('CDG v2 crear FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo crear la licencia. Reintentá. (' + e.message + ')' });
  }
});

// FIX (sáb 30-may-2026, server v19): rutas /sku-catalog/* ANTES de /:id.
// Express hace match en orden de definición. Si /:id va primero, captura
// "sku-catalog" como parámetro :id y las rutas de catálogo nunca se alcanzan.
// ── GET /api/cdg/v2/sku-catalog/sugerir ───────────────────────────────────
// Devuelve hasta 15 SKUs que matcheen el texto ingresado (sku o descripcion).
// Usado para autocomplete en el form CDG v2. Sin cargar los 41k SKUs al cliente.
// Query: ?q=texto (mínimo 2 chars para evitar queries demasiado amplias)
// FIX (lun 1-jun-2026, server v19): endpoint nuevo para autocomplete.
app.get('/api/cdg/v2/sku-catalog/sugerir', async (req, res) => {
  var q = String(req.query.q || '').trim();
  if(q.length < 2) return res.json({ ok: true, resultados: [] });

  try {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return res.json({ ok: true, resultados: [] });
    }
    // Buscar por SKU exacto primero (si es numérico) O por descripción ilike
    // PostgREST: or=(sku.eq.X,descripcion.ilike.*Y*)
    var qEnc = encodeURIComponent('%' + q + '%');
    var skuEnc = encodeURIComponent(q);
    var query = '?or=(sku.eq.' + skuEnc + ',descripcion.ilike.' + qEnc + ')'
              + '&select=sku,descripcion,costo&limit=15';
    var rows = await supabase('GET', 'sku_catalog', null, query);
    if(!Array.isArray(rows)) return res.json({ ok: true, resultados: [] });
    var resultados = rows.map(function(r){
      return {
        sku:        r.sku,
        descripcion: r.descripcion || '',
        costo:      (r.costo != null && Number(r.costo) > 0) ? Math.round(Number(r.costo)) : null
      };
    });
    res.json({ ok: true, resultados: resultados });
  } catch(e) {
    console.log('CDG v2 sku-catalog sugerir error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── GET /api/cdg/v2/sku-catalog/buscar ────────────────────────────────────
// Busca la descripción y costo de un SKU en el catálogo.
// Query: ?sku=XXXX
// FIX (lun 1-jun-2026, server v19): incluir costo en select y respuesta.
// Antes solo devolvía sku,descripcion — el cliente no podía autollenar el costo.
app.get('/api/cdg/v2/sku-catalog/buscar', async (req, res) => {
  var sku = req.query.sku;
  if(!sku) return res.status(400).json({ ok: false, error: 'falta sku' });

  try {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return res.json({ ok: true, encontrado: false });
    }
    var query = '?sku=eq.' + encodeURIComponent(String(sku).trim()) + '&select=sku,descripcion,costo';
    var rows  = await supabase('GET', 'sku_catalog', null, query);
    if(Array.isArray(rows) && rows.length > 0) {
      var row = rows[0];
      // costo: devolver como número o null (nunca como string con símbolo)
      var costo = (row.costo != null && row.costo !== '' && Number(row.costo) > 0)
        ? Number(row.costo)
        : null;
      res.json({ ok: true, encontrado: true, sku: row.sku, descripcion: row.descripcion, costo: costo });
    } else {
      res.json({ ok: true, encontrado: false });
    }
  } catch(e) {
    console.log('CDG v2 sku-catalog buscar error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── POST /api/cdg/v2/sku-catalog/upload ───────────────────────────────────
// Carga o actualiza el catálogo de SKUs desde un archivo Excel/CSV.
// Hace UPSERT por SKU (actualiza si existe, inserta si no).
// Multer field: 'file'. Columnas esperadas: SKU + Descripción (flexible).
app.post('/api/cdg/v2/sku-catalog/upload', upload.single('file'), async (req, res) => {
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    var wb   = XLSX.read(req.file.buffer, { type: 'buffer', raw: false });
    var sn   = wb.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: false });

    if(rows.length < 2) {
      return res.status(400).json({ ok: false, error: 'El archivo está vacío o solo tiene encabezado.' });
    }

    var hdr  = rows[0].map(function(h) { return norm(String(h)); });
    var cSku  = findCol(hdr, ['sku', 'articulo', 'artículo', 'codigo', 'código']);
    var cDesc = findCol(hdr, ['descripcion', 'descripción', 'desc', 'nombre', 'producto']);

    if(cSku < 0 || cDesc < 0) {
      return res.status(400).json({
        ok: false,
        error: 'No se encontraron columnas SKU y Descripción. Encabezados detectados: ' + rows[0].join(', ')
      });
    }

    // Construir lote de upserts
    var ahora   = new Date().toISOString();
    var lote    = [];
    var saltados = 0;
    rows.slice(1).forEach(function(row) {
      var sku  = String(row[cSku]  || '').trim();
      var desc = String(row[cDesc] || '').trim();
      if(!sku || !desc) { saltados++; return; }
      lote.push({ sku: sku, descripcion: desc, ts_actualizacion: ahora });
    });

    if(lote.length === 0) {
      return res.status(400).json({ ok: false, error: 'No se encontraron filas válidas (SKU + Descripción requeridos).' });
    }

    // Upsert en lotes de 500 para no superar límites de Supabase
    var insertados = 0;
    var tamLote    = 500;
    for(var i = 0; i < lote.length; i += tamLote) {
      var chunk = lote.slice(i, i + tamLote);
      await supabase('POST', 'sku_catalog', chunk, '?on_conflict=sku');
      insertados += chunk.length;
    }

    console.log('CDG v2 sku-catalog upload:', insertados, 'SKUs, saltados:', saltados);
    res.json({ ok: true, insertados: insertados, saltados: saltados });
  } catch(e) {
    console.log('CDG v2 sku-catalog upload error:', e.message);
    res.status(500).json({ ok: false, error: 'Error procesando el archivo: ' + e.message });
  }
});

// ── GET /api/cdg/v2/:id ────────────────────────────────────────────────────
// Polling del estado de la licencia. El cliente pasa su último timestamp
// conocido de líneas para recibir solo el delta.
// Query params: lineasDesde (ISO timestamp, opcional)
app.get('/api/cdg/v2/:id', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineasDesde = req.query.lineasDesde || null;

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });

    // FIX (sáb 30-may-2026, server v19): siempre consultar cdg_lineas cuando
    // el cliente manda lineasDesde. La optimización anterior (skip si
    // tsUltimaModifLineas no cambió) ocultaba líneas cuando el bump de metadata
    // fallaba — la línea existía en cdg_lineas pero ninguna tablet la veía
    // hasta hacer un refresh completo.
    // Sin la optimización: 1 query extra a Supabase por poll (aceptable).
    var lineas = await cdgGetLineas(licenciaId, lineasDesde);
    var respuesta = { ok: true, meta: meta, lineas: lineas };

    res.json(respuesta);
  } catch(e) {
    console.log('CDG v2 GET', licenciaId, 'error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── POST /api/cdg/v2/:id/linea ─────────────────────────────────────────────
// Agrega una línea nueva a la licencia.
// Body: { usuario, sku, descripcion, cantidad, costoUnit, fotos[] }
// fotos[]: array de paths en Storage (ya subidos antes de llamar este endpoint)
app.post('/api/cdg/v2/:id/linea', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, sku, descripcion, cantidad, costoUnit, fotos } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!sku)     return res.status(400).json({ ok: false, error: 'falta sku' });
  if(cantidad === undefined || cantidad === null) {
    return res.status(400).json({ ok: false, error: 'falta cantidad' });
  }

  // Lock: registrar escritura activa. Si la licencia está cerrando → rechazar.
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando. Ya no se pueden agregar líneas.' });
  }

  var fotosArr = Array.isArray(fotos) ? fotos : [];
  if(fotosArr.length > 3) {
    cdgLockRelease(licenciaId);
    return res.status(400).json({ ok: false, error: 'máximo 3 fotos por línea' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(!cdgAceptaLineas(meta)) {
      return res.status(403).json({ ok: false, error: 'La licencia está ' + meta.estado + '. No se pueden agregar líneas.' });
    }

    var now = new Date().toISOString();
    var linea = {
      licencia_id:  cdgNormId(licenciaId),
      sku:          String(sku).trim(),
      descripcion:  descripcion || '',
      cantidad:     Number(cantidad),
      costo_unit:   costoUnit !== undefined ? Number(costoUnit) : null,
      autor:        usuario,
      fotos:        fotosArr,
      ts_creacion:  now,
      ts_modif:     now,
      eliminada:    false
    };

    var lineaCreada = await withTimeout(
      cdgInsertLinea(linea),
      15000,
      'CDG insertar linea'
    );

    // Actualizar metadata: totalLineas y tsUltimaModifLineas
    // Registrar al usuario como activo si es la primera vez
    // FIX (sáb 30-may-2026, server v19): bump de metadata en try separado.
    // Si cdgInsertLinea ya persistió la línea y cdgBumpLineasTs falla,
    // el catch externo respondería 500 → el cliente reintentaría → línea duplicada.
    // Separando el try: si el bump falla, la línea está guardada, loggeamos y
    // respondemos 200 igual. El contador totalLineas se puede recalcular desde
    // cdg_lineas; tsUltimaModifLineas se corrige en el próximo bump exitoso.
    var licenciaCerradaMidFlight = false;
    try {
      if(!meta.usuarios[usuario]) {
        meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
        cdgBitacora(meta, usuario, 'usuario_unido', '');
      } else {
        meta.usuarios[usuario].lastActivity = now;
      }
      meta.totalLineas = (meta.totalLineas || 0) + 1;
      var bumpResult = await withTimeout(cdgBumpLineasTs(licenciaId, meta, +1), 15000, 'CDG bump ts');
      // Detectar si la licencia se cerró mientras esta operación estaba en vuelo
      if(bumpResult && bumpResult.estadoActual === 'cerrado') {
        licenciaCerradaMidFlight = true;
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata FAILED (línea ya insertada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea agregada:', licenciaId, 'sku:', sku, 'por:', usuario);
    if(licenciaCerradaMidFlight) {
      // La línea quedó guardada en cdg_lineas, pero la licencia se cerró
      // mientras operabas. ok:true porque el dato está seguro.
      return res.status(409).json({
        ok: true,
        linea: lineaCreada,
        aviso: 'La licencia fue cerrada mientras agregabas la línea. Tu línea quedó guardada en el sistema.'
      });
    }
    res.json({ ok: true, linea: lineaCreada });
  } catch(e) {
    console.log('CDG v2 agregar linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo guardar la línea. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId); // liberar escritura activa siempre
  }
});

// ── PATCH /api/cdg/v2/:id/linea/:lineaId ──────────────────────────────────
// Edita una línea propia (datos o fotos). Solo el autor puede editar su línea.
// Body: { usuario, sku?, descripcion?, cantidad?, costoUnit?, fotos? }
app.patch('/api/cdg/v2/:id/linea/:lineaId', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineaId    = req.params.lineaId;
  var { usuario, sku, descripcion, cantidad, costoUnit, fotos } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(meta.estado === 'cerrado') {
      return res.status(403).json({ ok: false, error: 'La licencia está cerrada. No se puede editar.' });
    }

    // Construir patch solo con campos enviados
    var patch = {};
    if(sku       !== undefined) patch.sku          = String(sku).trim();
    if(descripcion !== undefined) patch.descripcion = descripcion;
    if(cantidad  !== undefined) patch.cantidad      = Number(cantidad);
    if(costoUnit !== undefined) patch.costo_unit    = Number(costoUnit);
    if(fotos     !== undefined) {
      var fotosArr = Array.isArray(fotos) ? fotos : [];
      if(fotosArr.length > 3) {
        return res.status(400).json({ ok: false, error: 'máximo 3 fotos por línea' });
      }
      patch.fotos = fotosArr;
    }

    if(Object.keys(patch).length === 0) {
      return res.status(400).json({ ok: false, error: 'nada que actualizar' });
    }

    var lineaActualizada = await withTimeout(
      cdgUpdateLinea(lineaId, patch, usuario),
      15000,
      'CDG editar linea'
    );

    if(!lineaActualizada) {
      return res.status(403).json({ ok: false, error: 'No se encontró la línea o no sos el autor.' });
    }

    // Actualizar tsUltimaModifLineas para que el próximo poll la incluya.
    // FIX (sáb 30-may-2026, server v19): try separado — si la edición ya
    // persistió y el bump falla, respondemos 200 igual (la línea está guardada).
    try {
      if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
      var bumpEditResult = await withTimeout(cdgBumpLineasTs(licenciaId, meta, 0), 15000, 'CDG bump ts edit');
      if(bumpEditResult && bumpEditResult.estadoActual === 'cerrado') {
        console.log('CDG v2 linea editada (licencia cerrada mid-flight):', lineaId);
        return res.status(409).json({ ok: true, linea: lineaActualizada, aviso: 'La licencia fue cerrada mientras editabas. Tu cambio quedó guardado.' });
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata (edit) FAILED (línea ya editada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea editada:', lineaId, 'por:', usuario);
    res.json({ ok: true, linea: lineaActualizada });
  } catch(e) {
    console.log('CDG v2 editar linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo editar la línea. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// ── DELETE /api/cdg/v2/:id/linea/:lineaId ─────────────────────────────────
// Soft-delete de una línea. Solo el autor puede borrar su propia línea.
// Body: { usuario }
app.delete('/api/cdg/v2/:id/linea/:lineaId', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineaId    = req.params.lineaId;
  var { usuario } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(meta.estado === 'cerrado') {
      return res.status(403).json({ ok: false, error: 'La licencia está cerrada. No se puede eliminar.' });
    }

    var resultado = await withTimeout(
      cdgSoftDeleteLinea(lineaId, usuario),
      15000,
      'CDG delete linea'
    );

    if(!resultado) {
      return res.status(403).json({ ok: false, error: 'No se encontró la línea o no sos el autor.' });
    }

    // Actualizar metadata.
    // FIX (sáb 30-may-2026, server v19): try separado — si el soft-delete ya
    // persistió y el bump falla, respondemos 200 igual (la línea ya está eliminada).
    try {
      meta.totalLineas = Math.max(0, (meta.totalLineas || 1) - 1);
      if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
      var bumpDelResult = await withTimeout(cdgBumpLineasTs(licenciaId, meta, -1), 15000, 'CDG bump ts delete');
      if(bumpDelResult && bumpDelResult.estadoActual === 'cerrado') {
        console.log('CDG v2 linea eliminada (licencia cerrada mid-flight):', lineaId);
        return res.status(409).json({ ok: true, aviso: 'La licencia fue cerrada mientras eliminabas. Tu cambio quedó guardado.' });
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata (delete) FAILED (línea ya eliminada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea eliminada:', lineaId, 'por:', usuario);
    res.json({ ok: true });
  } catch(e) {
    console.log('CDG v2 delete linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo eliminar la línea. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// ── POST /api/cdg/v2/:id/fotos-encabezado ─────────────────────────────────
// Agrega fotos de encabezado a la licencia (máx 5).
// Body: { usuario, fotos[] } — fotos son paths en Storage ya subidos
app.post('/api/cdg/v2/:id/fotos-encabezado', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, fotos } = req.body;

  if(!usuario)               return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!Array.isArray(fotos) || fotos.length === 0) {
    return res.status(400).json({ ok: false, error: 'falta fotos' });
  }
  // Lock: registrar write de metadata. Si cerrando → 423.
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(!cdgAceptaLineas(meta)) {
      return res.status(403).json({ ok: false, error: 'La licencia está ' + meta.estado });
    }

    var actuales = meta.fotosEncabezado || [];
    var total    = actuales.length + fotos.length;
    if(total > 5) {
      return res.status(400).json({
        ok: false,
        error: 'Máximo 5 fotos de encabezado. Ya tenés ' + actuales.length + ', intentás agregar ' + fotos.length + '.'
      });
    }

    meta.fotosEncabezado = actuales.concat(fotos);
    meta.version = (meta.version || 0) + 1;
    if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();

    // FIX (lun 1-jun-2026, v19): re-leer meta fresco antes de guardar para
    // detectar cierre concurrente — mismo patrón que /accion bloque compartido.
    var metaFreshFotos = await cdgGetMeta(licenciaId);
    if(metaFreshFotos && metaFreshFotos.estado === 'cerrado') {
      return res.status(409).json({ ok: false, error: 'La licencia fue cerrada mientras subías la foto. No se guardó.' });
    }
    await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG fotos encabezado save');
    console.log('CDG v2 fotos encabezado:', licenciaId, '+' + fotos.length + ' fotos');
    res.json({ ok: true, fotosEncabezado: meta.fotosEncabezado });
  } catch(e) {
    console.log('CDG v2 fotos encabezado FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudieron guardar las fotos. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// ── POST /api/cdg/v2/:id/accion ────────────────────────────────────────────
// Centraliza todas las operaciones de estado de licencia y usuarios.
// Body: { usuario, tipo, ...params }
//
// Tipos soportados:
//   unirse          → usuario se une a la licencia
//   guardar_progreso → registra lastActivity del usuario (sin cerrar)
//   pausar          → usuario pasa a estado pausado
//   reanudar        → usuario vuelve a activo
//   marcar_inactivo → cliente informa que el usuario llegó a 45 min sin actividad
//   delegar         → finalizador delega el rol a otro usuario activo
//                     params: { nuevoFinalizador }
//   cerrar          → finalizador cierra la licencia definitivamente
//   desbloquear     → supervisor reabre una licencia cerrada
//                     params: { supervisor: true }
app.post('/api/cdg/v2/:id/accion', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, tipo, nuevoFinalizador, supervisor } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!tipo)    return res.status(400).json({ ok: false, error: 'falta tipo de acción' });

  // Lock para acciones que escriben metadata (no-cerrar).
  // 'cerrar' gestiona el lock internamente con cdgLockStartClosing/WaitDrain.
  var lockAdquirido = false;
  if(tipo !== 'cerrar') {
    if(!cdgLockAcquire(licenciaId)) {
      return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
    }
    lockAdquirido = true;
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });

    var now = new Date().toISOString();

    // ── unirse ──────────────────────────────────────────────────────────────
    if(tipo === 'unirse') {
      if(meta.estado === 'cerrado') {
        return res.status(403).json({ ok: false, error: 'La licencia está cerrada.' });
      }
      if(!meta.usuarios[usuario]) {
        meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
        cdgBitacora(meta, usuario, 'usuario_unido', '');
      } else {
        // Ya estaba registrado — reactivar
        meta.usuarios[usuario].estado       = 'activo';
        meta.usuarios[usuario].lastActivity = now;
        cdgBitacora(meta, usuario, 'usuario_reingreso', '');
      }
    }

    // ── guardar_progreso ────────────────────────────────────────────────────
    else if(tipo === 'guardar_progreso') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].lastActivity = now;
      meta.usuarios[usuario].estado       = 'activo';
      cdgBitacora(meta, usuario, 'progreso_guardado', '');
    }

    // ── pausar ──────────────────────────────────────────────────────────────
    else if(tipo === 'pausar') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado       = 'pausado';
      meta.usuarios[usuario].lastActivity = now;
      cdgBitacora(meta, usuario, 'usuario_pausado', '');
    }

    // ── reanudar ────────────────────────────────────────────────────────────
    else if(tipo === 'reanudar') {
      if(meta.estado === 'cerrado') {
        return res.status(403).json({ ok: false, error: 'La licencia está cerrada.' });
      }
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado       = 'activo';
      meta.usuarios[usuario].lastActivity = now;
      cdgBitacora(meta, usuario, 'usuario_reanudo', '');
    }

    // ── marcar_inactivo ─────────────────────────────────────────────────────
    else if(tipo === 'marcar_inactivo') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado = 'inactivo';
      cdgBitacora(meta, usuario, 'usuario_inactivo_auto', '45 min sin actividad');

      // Reasignación automática lazy del finalizador si quedó inactivo
      if(meta.finalizador === usuario) {
        var candidatos = Object.keys(meta.usuarios).filter(function(u) {
          return u !== usuario && meta.usuarios[u].estado === 'activo';
        });
        if(candidatos.length > 0) {
          // Criterio: usuario con actividad más reciente
          candidatos.sort(function(a, b) {
            return (meta.usuarios[b].lastActivity || '') > (meta.usuarios[a].lastActivity || '') ? 1 : -1;
          });
          var nuevo = candidatos[0];
          meta.finalizador = nuevo;
          cdgBitacora(meta, 'sistema', 'finalizador_reasignado_auto',
            'De ' + usuario + ' a ' + nuevo + ' (inactividad)');
        }
      }
    }

    // ── delegar ─────────────────────────────────────────────────────────────
    else if(tipo === 'delegar') {
      if(meta.finalizador !== usuario) {
        return res.status(403).json({ ok: false, error: 'Solo el finalizador actual puede delegar.' });
      }
      if(!nuevoFinalizador) {
        return res.status(400).json({ ok: false, error: 'falta nuevoFinalizador' });
      }
      if(!meta.usuarios[nuevoFinalizador]) {
        return res.status(400).json({ ok: false, error: nuevoFinalizador + ' no está en esta licencia.' });
      }
      meta.finalizador = nuevoFinalizador;
      cdgBitacora(meta, usuario, 'finalizador_delegado',
        'De ' + usuario + ' a ' + nuevoFinalizador);
    }

    // ── cerrar ──────────────────────────────────────────────────────────────
    else if(tipo === 'cerrar') {
      if(meta.estado === 'cerrado') {
        return res.status(400).json({ ok: false, error: 'La licencia ya está cerrada.' });
      }
      // FIX (lun 1-jun-2026, v19): permitir cierre a cualquier usuario activo
      // unido a la licencia, no solo al finalizador.
      // Razón operativa: en multiusuario el finalizador puede no estar disponible.
      // Supervisor (isSup desde el cliente) siempre puede cerrar.
      var esFinalizador   = meta.finalizador === usuario;
      var esActivoEnLic   = meta.usuarios && meta.usuarios[usuario]
                            && meta.usuarios[usuario].estado === 'activo';
      var esSupervisorCDG = req.body.esSupervisor === true;
      if(!esFinalizador && !esActivoEnLic && !esSupervisorCDG) {
        return res.status(403).json({
          ok: false,
          error: 'Solo usuarios activos de la licencia o el finalizador pueden cerrar.'
        });
      }
      // FIX (lun 1-jun-2026, v19): orden correcto — lock primero, re-leer fresco,
      // luego aplicar TODOS los campos de cierre al meta fresco para no perderlos.
      // 1. closing=true en memoria → nuevos writes reciben 423 de inmediato
      // 2. drain → esperar writes activos (hasta 15s; lanza si no drena)
      // 3. re-leer meta fresco → tiene fotos/bitácora de todos los writes anteriores
      // 4. aplicar TODOS los campos de cierre al fresco
      // 5. guardar en Supabase → lock distribuido
      // 6. leer líneas con garantía
      cdgLockStartClosing(licenciaId);
      await cdgLockWaitDrain(licenciaId, 15000);

      var metaFresco = await cdgGetMeta(licenciaId);
      var base = metaFresco || meta;

      // Aplicar TODOS los campos de cierre sobre el meta más reciente
      base.estado             = 'cerrado';
      base.fechaCierre        = now;
      base.cerradoPor         = usuario;
      base.finalizadorOriginal = base.finalizador;
      base.version            = (base.version || 0) + 1;
      cdgBitacora(base, usuario, 'licencia_cerrada',
        esFinalizador ? 'por finalizador' : esSupervisorCDG ? 'por supervisor' : 'por usuario activo');
      Object.keys(base.usuarios || {}).forEach(function(u){
        base.usuarios[u].estado = 'inactivo';
      });
      meta = base;

      await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG cerrar lock');

      // Snapshot de state ANTES de mutar para rollback si falla Hamilton.
      var snapCdg          = state.cdg     ? JSON.parse(JSON.stringify(state.cdg))     : {};
      var snapTeorico      = state.teorico ? JSON.parse(JSON.stringify(state.teorico)) : {};
      var snapFisico       = state.fisico  ? JSON.parse(JSON.stringify(state.fisico))  : {};
      var snapVersion      = state.version;
      var snapHistorialLen = (state.historial||[]).length;

      var lineasV2 = await cdgGetLineas(licenciaId, null);
      var itemsV2  = lineasV2.map(function(l) {
        return {
          sku:   l.sku,
          desc:  l.descripcion || '',
          qty:   l.cantidad,
          costo: l.costo_unit != null ? l.costo_unit : null,
          fotos: Array.isArray(l.fotos) ? l.fotos : [],
          autor: l.autor
        };
      });
      if(!state.cdg) state.cdg = {};
      state.cdg[licenciaId] = {
        items:      itemsV2,
        status:     'locked',
        autor:      meta.creadoPor,
        fecha:      new Date().toLocaleDateString('es'),
        tipo:       meta.tipo || 'CDG',
        bloqueado:  true,
        lastEditor: usuario,
        fromCDGv2:  true
      };
      if(!state.teorico) state.teorico = {};
      if(!state.teorico[licenciaId]) {
        // FIX (lun 1-jun-2026, server v20): enriquecer items v2 con teoricoWMS.
        var itemsV2base = itemsV2.map(function(s) {
          return { sku: s.sku, desc: s.desc, qty: s.qty,
            raw: { origen: 'CDG', status: 'CDG Validado', tipo: meta.tipo || 'CDG' } };
        });
        var itemsV2enc = await cdgEnriquecerItemsWMS(licenciaId, itemsV2base);
        state.teorico[licenciaId] = {
          items: itemsV2enc,
          type:        meta.tipo || 'CDG',
          fromCDG:     true,
          cdgRef:      licenciaId,
          cdgValidado: true,
          cdgBloqueado: true
        };
        if(!state.fisico) state.fisico = {};
        state.fisico[licenciaId] = null;
      }
      addHistorial(usuario, 'CDG v2 finalizado', licenciaId);
      state.version++;
      // Persistir de forma directa y estricta. Si falla → revertir y lanzar
      // para que el catch del endpoint responda 500 (no ok:true mentiroso).
      try {
        await withTimeout(saveDailyStateStrict('CDG v2 cerrar Hamilton'), 20000, 'CDG v2 cerrar Hamilton save');
      } catch(saveErr) {
        // Rollback: restaurar state a como estaba antes de las mutaciones
        state.cdg      = snapCdg;
        state.teorico  = snapTeorico;
        state.fisico   = snapFisico;
        state.version  = snapVersion;
        if(state.historial) state.historial.length = snapHistorialLen;
        console.log('CDG v2 cerrar Hamilton ROLLBACK por fallo de persistencia:', saveErr.message);
        // Best-effort: reabrir la licencia en Supabase restaurando estado + usuarios
        meta.estado = 'activo'; meta.fechaCierre = null;
        Object.keys(meta.usuarios || {}).forEach(function(u){
          if(meta.usuarios[u]) meta.usuarios[u].estado = 'activo';
        });
        cdgLockClear(licenciaId); // limpiar lock para permitir nuevas escrituras
        cdgSaveMeta(licenciaId, meta).catch(function(e2){
          console.log('CDG v2 cerrar: no se pudo reabrir licencia tras rollback:', e2.message);
        });
        throw saveErr; // el catch externo responde 500
      }
      console.log('CDG v2 cerrar: entrada Hamilton creada para', licenciaId, '-', itemsV2.length, 'líneas');
      cdgLockClear(licenciaId); // limpiar lock — licencia cerrada, ya no necesita rastreo
    }

    // ── desbloquear ─────────────────────────────────────────────────────────
    else if(tipo === 'desbloquear') {
      if(!supervisor) {
        return res.status(403).json({ ok: false, error: 'Solo un supervisor puede desbloquear.' });
      }
      if(meta.estado !== 'cerrado') {
        return res.status(400).json({ ok: false, error: 'La licencia no está cerrada.' });
      }
      meta.estado      = 'activo';
      meta.fechaCierre = null;
      cdgBitacora(meta, usuario, 'licencia_desbloqueada', 'supervisor');
    }

    else {
      return res.status(400).json({ ok: false, error: 'Tipo de acción desconocido: ' + tipo });
    }

    // Para 'cerrar': meta ya fue guardado con lock early (meta.version ya incrementado).
    // El bloque compartido haría un segundo increment y save redundante — lo saltamos.
    if(tipo !== 'cerrar') {
      // FIX (lun 1-jun-2026, v19): re-leer meta fresco antes de guardar.
      // Escenario sin este fix: guardar_progreso lee meta (activo), cierre guarda
      // meta cerrado + crea Hamilton, guardar_progreso termina y pisa con meta stale
      // → licencia reabierta en Supabase con Hamilton ya creado.
      // EXCEPCIÓN: 'desbloquear' tiene que guardar estado='activo' sobre una licencia
      // cerrada — es exactamente su propósito, no debe ser bloqueado por esta guardia.
      var metaFresh = await cdgGetMeta(licenciaId);
      if(metaFresh && metaFresh.estado === 'cerrado' && tipo !== 'desbloquear') {
        return res.status(409).json({ ok: false, error: 'La licencia fue cerrada mientras procesabas esta acción. No se guardaron cambios.' });
      }
      meta.version = (meta.version || 0) + 1;
      await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG accion save');
    }
    console.log('CDG v2 accion:', tipo, licenciaId, 'por:', usuario);
    res.json({ ok: true, meta: meta });

  } catch(e) {
    console.log('CDG v2 accion FAILED:', tipo, licenciaId, e.message);
    if(tipo === 'cerrar') cdgLockClear(licenciaId);
    res.status(500).json({ ok: false, error: 'No se pudo ejecutar la acción. Reintentá. (' + e.message + ')' });
  } finally {
    if(lockAdquirido) cdgLockRelease(licenciaId);
  }
});

// ══════════════════════════════════════════════════════════════════════════
// FIN CDG v2 — Los endpoints /api/cdg/* originales (v1 legacy) se conservan
// sin modificación arriba de este bloque para rollback si fuera necesario.
// ══════════════════════════════════════════════════════════════════════════

// ── Start ─────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
// Always start server — even if Supabase fails
app.listen(PORT, () => console.log('Conteo app on port ' + PORT));
// Load state after server is up
loadState().catch(e => console.log('State load failed (non-fatal):', e.message));

// ══════════════════════════════════════════════════════════════════════════
// CDG v2 — MANIFIESTO / VALIDACIÓN WMS
// FIX (lun 1-jun-2026, server v20): endpoints WMS para comparación CDG vs WMS.
// Tabla: cdg_wms (una fila por licencia, UPSERT al recargar).
// Routing: declarados ANTES de /:id para evitar captura por Express.
// ══════════════════════════════════════════════════════════════════════════

// ── Helpers WMS ────────────────────────────────────────────────────────────

async function cdgGetWms(licenciaId) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var rows = await supabase('GET', 'cdg_wms', null,
    '?licencia_id=eq.' + encodeURIComponent(cdgNormId(licenciaId)) + '&limit=1');
  return (Array.isArray(rows) && rows.length > 0) ? rows[0] : null;
}

async function cdgUpsertWms(licenciaId, data) {
  // UPSERT por licencia_id. Reemplaza completamente si ya existe.
  await supabase('POST', 'cdg_wms', data, '?on_conflict=licencia_id');
}


// Helper: normalizar valor de columna Número del WMS.
// El Excel puede traer "U25-161959 · CDG" o "U25-161959 - CDG".
// Se extrae solo la parte izquierda del separador para comparar con licenciaId.
function cdgNormLicWMS(v) {
  var s = String(v || '').trim().toUpperCase();
  // Separador "·" (U+00B7, punto medio) con o sin espacios
  var dotIdx = s.indexOf('·');
  if(dotIdx >= 0) s = s.substring(0, dotIdx);
  // Separador " - " con espacios
  var dashIdx = s.indexOf(' - ');
  if(dashIdx >= 0) s = s.substring(0, dashIdx);
  return s.trim();
}

// Parseo robusto de cantidad: soporta 1234, "1,234", "1,234.00", espacios.
// Retorna null si no es numérico o <= 0.
function cdgParseQty(val) {
  if(val === null || val === undefined || val === '') return null;
  var s = String(val).trim().replace(/,/g, '');
  var n = Number(s);
  if(isNaN(n) || n <= 0) return null;
  return n;
}

// Normaliza un encabezado de columna para detección flexible:
// "Cant." → "cant", "# Ingreso" → "ingreso", "Número" → "numero"
function cdgNormHdr(h) {
  return String(h || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quitar tildes
    .replace(/[^a-z0-9]/g, '');                        // solo alfanumérico
}

// FIX (mar 2-jun-2026, server v20): rutas fijas /bulk y /sincronizar-hamilton
// DEBEN ir ANTES de la ruta paramétrica /:id. Express hace match en orden.
// Si /:id va primero, captura "bulk" como parámetro :id.

// ══════════════════════════════════════════════════════════════════════════
// CDG WMS BULK — POST /api/cdg/wms/bulk
// FIX (lun 1-jun-2026, server v20): carga masiva de WMS para licencias antiguas.
// Recibe el mismo Excel WMS (hoja Traslados) y procesa TODAS las licencias.
// No requiere que la licencia exista como cdg_meta_* v2; sirve para CDG clásico.
// No crea conteos, no modifica cdg_lineas, no toca Hamilton.
// ══════════════════════════════════════════════════════════════════════════
app.post('/api/cdg/wms/bulk', upload.single('file'), async (req, res) => {
  var usuario = String(req.body && req.body.usuario ? req.body.usuario : '').trim();
  if(!usuario)  return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    // ── 1. Parsear Excel ─────────────────────────────────────────────────
    var wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: 'buffer', raw: false, cellDates: true });
    } catch(e) {
      return res.status(400).json({ ok: false, error: 'El archivo no es un Excel válido. (' + e.message + ')' });
    }

    var advertencias = [];
    var sheetName = wb.SheetNames.find(function(n){ return n.trim().toLowerCase() === 'traslados'; });
    if(!sheetName) {
      sheetName = wb.SheetNames[0];
      advertencias.push('Hoja "Traslados" no encontrada; se usó la primera hoja: "' + sheetName + '".');
    }
    var ws   = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if(rows.length < 2) return res.status(400).json({ ok: false, error: 'El archivo está vacío o solo tiene encabezado.' });

    // ── 2. Detectar columnas (reutiliza la misma lógica del endpoint individual) ─
    var hdr    = rows[0].map(cdgNormHdr);
    var cNum   = hdr.indexOf('numero');
    var cSku   = hdr.indexOf('sku');
    var cNombre = hdr.indexOf('nombre');
    if(cNombre < 0) cNombre = hdr.indexOf('descripcion');
    var cCant  = hdr.indexOf('cant');
    var cUnis  = hdr.indexOf('unidades');
    var cFecha = hdr.indexOf('fecha');
    var cOrig  = hdr.indexOf('origen');
    var cDest  = hdr.indexOf('destino');
    var cTipo  = hdr.indexOf('tipo');
    var cStatus = hdr.indexOf('status');
    var cLineas = hdr.indexOf('lineas');

    var faltantes = [];
    if(cNum  < 0) faltantes.push('"Número"');
    if(cSku  < 0) faltantes.push('"SKU"');
    if(cCant < 0 && cUnis < 0) faltantes.push('"Cant." o "Unidades"');
    if(faltantes.length) {
      return res.status(400).json({
        ok: false,
        error: 'Columnas requeridas no encontradas: ' + faltantes.join(', ') + '. Encabezados: ' + rows[0].join(', ')
      });
    }

    // ── 3. Agrupar filas por licencia ─────────────────────────────────────
    var porLicencia = {};  // licNorm → { skuMap:{}, encabezado, filaCount }
    rows.slice(1).forEach(function(r) {
      var licNorm = cdgNormLicWMS(r[cNum]);
      if(!licNorm) return;
      var sku  = String(r[cSku]  || '').trim();
      if(!sku) return;
      var qty  = cdgParseQty(cCant >= 0 ? r[cCant] : null);
      if(qty === null) qty = cdgParseQty(cUnis >= 0 ? r[cUnis] : null);
      if(qty === null) return;
      var desc = String(cNombre >= 0 ? (r[cNombre] || '') : '').trim();

      if(!porLicencia[licNorm]) {
        porLicencia[licNorm] = {
          skuMap: {},
          skuOrder: [],
          encabezado: {
            fecha:     cFecha  >= 0 ? (r[cFecha]  || null) : null,
            origen:    cOrig   >= 0 ? (r[cOrig]   || null) : null,
            destino:   cDest   >= 0 ? (r[cDest]   || null) : null,
            tipo:      cTipo   >= 0 ? (r[cTipo]   || null) : null,
            status:    cStatus >= 0 ? (r[cStatus] || null) : null,
            lineasWMS: cLineas >= 0 ? Number(r[cLineas] || 0) : 0
          }
        };
      }
      var sm = porLicencia[licNorm].skuMap;
      if(sm[sku]) {
        sm[sku].cantidad += qty;
        if(!sm[sku].descripcion && desc) sm[sku].descripcion = desc;
      } else {
        sm[sku] = { descripcion: desc, cantidad: qty };
        porLicencia[licNorm].skuOrder.push(sku);
      }
    });

    var licencias = Object.keys(porLicencia);
    if(!licencias.length) {
      return res.status(400).json({ ok: false, error: 'No se encontraron filas válidas en el archivo.' });
    }

    // ── 4. Enriquecer descripciones vacías con sku_catalog (batch por licencia) ──
    // Recopilar todos los SKUs sin descripción en un solo set
    var skusSinDesc = [];
    licencias.forEach(function(lic) {
      var sm = porLicencia[lic].skuMap;
      Object.keys(sm).forEach(function(sku){ if(!sm[sku].descripcion) skusSinDesc.push(sku); });
    });
    var descMap = {};
    if(skusSinDesc.length > 0 && SUPABASE_URL && SUPABASE_KEY) {
      var unique = skusSinDesc.filter(function(v, i, a){ return a.indexOf(v) === i; });
      var chunkSize = 200;
      for(var ci = 0; ci < unique.length; ci += chunkSize) {
        try {
          var chunk = unique.slice(ci, ci + chunkSize);
          var cat = await supabase('GET', 'sku_catalog', null,
            '?sku=in.(' + chunk.map(function(s){ return encodeURIComponent(s); }).join(',') + ')&select=sku,descripcion');
          if(Array.isArray(cat)) cat.forEach(function(row){ if(row.descripcion) descMap[row.sku] = row.descripcion; });
        } catch(e) { advertencias.push('Enriquecimiento parcial sku_catalog: ' + e.message); }
      }
    }

    // ── 5. Upsert por licencia ─────────────────────────────────────────────
    var now = new Date().toISOString();
    var nombreArchivo = req.file.originalname || 'bulk.xlsx';
    var licenciasActualizadas = [];
    var licenciasSinConteoCDG = [];
    var errores = [];
    var totalSkus = 0;

    for(var li = 0; li < licencias.length; li++) {
      var lic = licencias[li];
      var data = porLicencia[lic];
      var skusArray = data.skuOrder.map(function(sku) {
        var entry = data.skuMap[sku];
        var desc  = entry.descripcion || descMap[sku] || '';
        return { sku: sku, descripcion: desc, cantidad: entry.cantidad };
      });
      totalSkus += skusArray.length;

      // Verificar si existe como v2 (meta) — solo informativo, no bloquea
      var metaV2 = null;
      try { metaV2 = await cdgGetMeta(lic); } catch(e) { /* no existe como v2 */ }
      if(!metaV2 && !state.cdg[lic] && !state.teorico[lic]) {
        licenciasSinConteoCDG.push(lic);
      }

      try {
        await cdgUpsertWms(lic, {
          licencia_id:    lic,
          skus:           JSON.stringify(skusArray),
          encabezado:     JSON.stringify(data.encabezado),
          cargado_por:    usuario,
          ts_carga:       now,
          nombre_archivo: nombreArchivo,
          total_skus:     skusArray.length,
          total_unidades: skusArray.reduce(function(a, s){ return a + s.cantidad; }, 0)
        });
        licenciasActualizadas.push(lic);
      } catch(upsertErr) {
        errores.push({ licencia: lic, error: upsertErr.message });
      }
    }

    console.log('CDG WMS bulk:', usuario, '—', licenciasActualizadas.length, 'licencias,', totalSkus, 'SKUs');
    res.json({
      ok:                      true,
      totalLicenciasProcesadas: licencias.length,
      totalSkus:               totalSkus,
      licenciasActualizadas:   licenciasActualizadas,
      licenciasSinConteoCDG:  licenciasSinConteoCDG,
      errores:                 errores,
      advertencias:            advertencias.length ? advertencias : undefined
    });

  } catch(e) {
    console.log('CDG WMS bulk FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'Error en bulk WMS: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════
// CDG WMS SINCRONIZAR HAMILTON — POST /api/cdg/wms/sincronizar-hamilton
// FIX (mar 2-jun-2026, server v20 rev2):
//   - Consolida por SKU antes de cruzar con WMS (fix bug líneas repetidas).
//     CDG clásico puede tener N filas del mismo SKU; Hamilton necesita una sola.
//   - Fuente de consolidación: (1) cdg_lineas v2, (2) state.cdg[lic].items v1,
//     (3) teo.items existente como fallback.
//   - Remapea fisico existente al nuevo alineamiento por SKU consolidado.
//   - snapshot/rollback de teorico+fisico+version+historial.
//   - Busca por licId directo y por fallback cdgRef.
//   - Recalcula siempre con el WMS vigente.
//   - No crea conteos nuevos. No toca cdg_lineas. No toca Hamilton general.
// ══════════════════════════════════════════════════════════════════════════
app.post('/api/cdg/wms/sincronizar-hamilton', async (req, res) => {
  var usuario = String((req.body && req.body.usuario) || '').trim();
  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });

  try {
    // FIX (mar 2-jun-2026, server v20 rev): procesar cdg_wms página por página.
    // NO acumular en wmsAllRows — cada página se procesa y descarta para no
    // retener skus jsonb completos de 200 licencias en memoria simultáneamente.
    var PAGE = 200;
    var totalProcesadas = 0;
    var huboWms = false;

    // ── Snapshot ANTES de mutar — para rollback si el save falla ──────────
    var snapTeorico   = state.teorico ? JSON.parse(JSON.stringify(state.teorico)) : {};
    var snapFisico    = state.fisico   ? JSON.parse(JSON.stringify(state.fisico))  : {};
    var snapVersion   = state.version;
    var snapHistorial = state.historial ? state.historial.length : 0;

    var actualizadas   = [];
    var sinConteo      = [];
    var advertencias   = [];

    for(var pg = 0; pg < 10; pg++) {   // máx 10 páginas = 2000 licencias
      var pageRows = await supabase('GET', 'cdg_wms', null,
        '?select=licencia_id,skus&limit='+PAGE+'&offset='+(pg*PAGE));
      if(!Array.isArray(pageRows) || !pageRows.length) break;
      huboWms = true;
      totalProcesadas += pageRows.length;

      for(var wi = 0; wi < pageRows.length; wi++) {
      var wmsRow = pageRows[wi];
      var lic    = String(wmsRow.licencia_id || '').trim().toUpperCase();
      if(!lic) continue;

      // Buscar entrada Hamilton: (a) directo, (b) fallback por cdgRef
      var teoKey = null;
      if(state.teorico && state.teorico[lic] &&
         (state.teorico[lic].fromCDG || state.teorico[lic].cdgValidado)) {
        teoKey = lic;
      } else {
        var keys = state.teorico ? Object.keys(state.teorico) : [];
        for(var ki = 0; ki < keys.length; ki++) {
          var k = keys[ki];
          var t = state.teorico[k];
          if(t && (t.fromCDG || t.cdgValidado) && String(t.cdgRef || '').trim().toUpperCase() === lic) {
            teoKey = k; break;
          }
        }
      }
      if(!teoKey) { sinConteo.push(lic); continue; }

      var teo = state.teorico[teoKey];

      // ── Capturar items y fisico ANTES de cualquier mutación ──────────────
      // CRÍTICO: viejoItems debe leerse ANTES de modificar teo.items.
      // Si se lee después de la consolidación, ya no tiene las N filas originales
      // del mismo SKU y el remap de fisico queda incorrecto (suma solo la primera fila).
      // FIX (mar 2-jun-2026, server v20 rev3): mover captura aquí, antes del consolidado.
      var viejoFisico = state.fisico && state.fisico[teoKey]
        ? (Array.isArray(state.fisico[teoKey]) ? state.fisico[teoKey].slice() : null)
        : null;
      var viejoItems = (Array.isArray(teo.items) ? teo.items : [])
        .filter(function(it){ return !it._soloWMS; });
      try {
        skusWMS = typeof wmsRow.skus === 'string' ? JSON.parse(wmsRow.skus) : (wmsRow.skus || []);
      } catch(e) { skusWMS = []; }
      if(!skusWMS.length) continue;

      // Mapa WMS: SKU_UPPER → { cantidad, descripcion }
      var wmsMap = {};
      var wmsDescMap = {};
      skusWMS.forEach(function(s){
        var sk = String(s.sku || '').trim().toUpperCase();
        wmsMap[sk] = (wmsMap[sk] || 0) + (Number(s.cantidad) || 0);
        if(!wmsDescMap[sk] && s.descripcion) wmsDescMap[sk] = s.descripcion;
      });

      // ── Construir base consolidada por SKU ─────────────────────────────
      // Prioridad:
      //   1. CDG v2: leer cdg_lineas vigentes y agrupar por SKU.
      //   2. CDG v1: leer state.cdg[lic].items y agrupar por SKU.
      //   3. Fallback: items existentes en teo (también agrupar — puede tener dupes).
      // En todos los casos: UNA fila por SKU en la base consolidada.
      // { skUpper: { sku, desc, qty, costo, orden } }
      var cdgConsolidado = {};  // SK_UPPER → { sku, desc, qty, costo, orden }
      var ordenSku = [];        // orden de primera aparición

      // Intento 1: CDG v2 desde cdg_lineas
      var metaV2 = null;
      try { metaV2 = await cdgGetMeta(lic); } catch(e) {}
      if(metaV2) {
        var lineasV2 = await cdgGetLineas(lic, null); // carga completa, solo activas
        if(Array.isArray(lineasV2) && lineasV2.length) {
          lineasV2.forEach(function(l){
            var sk = String(l.sku || '').trim().toUpperCase();
            if(!sk) return;
            var qty = Number(l.cantidad) || 0;
            if(cdgConsolidado[sk]) {
              cdgConsolidado[sk].qty += qty;
              if(!cdgConsolidado[sk].desc && l.descripcion) cdgConsolidado[sk].desc = l.descripcion;
              if(cdgConsolidado[sk].costo == null && l.costo_unit != null) cdgConsolidado[sk].costo = l.costo_unit;
            } else {
              cdgConsolidado[sk] = { sku: l.sku, desc: l.descripcion||'', qty: qty, costo: l.costo_unit||null, orden: ordenSku.length };
              ordenSku.push(sk);
            }
          });
        }
      }

      // Intento 2: CDG v1 desde state.cdg
      if(!ordenSku.length && state.cdg && state.cdg[lic]) {
        var itemsV1 = state.cdg[lic].items || [];
        itemsV1.forEach(function(it){
          var sk = String(it.sku || '').trim().toUpperCase();
          if(!sk) return;
          var qty = Number(it.qty) || Number(it.cantidad) || 0;
          if(cdgConsolidado[sk]) {
            cdgConsolidado[sk].qty += qty;
            if(!cdgConsolidado[sk].desc && it.desc) cdgConsolidado[sk].desc = it.desc;
            if(cdgConsolidado[sk].costo == null && it.costo != null) cdgConsolidado[sk].costo = it.costo;
          } else {
            cdgConsolidado[sk] = { sku: it.sku, desc: it.desc||'', qty: qty, costo: it.costo||null, orden: ordenSku.length };
            ordenSku.push(sk);
          }
        });
      }

      // Fallback 3: items existentes en teo (sin _soloWMS), también consolidando
      if(!ordenSku.length) {
        var itemsFallback = (Array.isArray(teo.items) ? teo.items : [])
          .filter(function(it){ return !it._soloWMS; });
        itemsFallback.forEach(function(it){
          var sk = String(it.sku || '').trim().toUpperCase();
          if(!sk) return;
          var qty = Number(it.qty) || Number(it.cantidad) || 0;
          if(cdgConsolidado[sk]) {
            cdgConsolidado[sk].qty += qty;
            if(!cdgConsolidado[sk].desc && it.desc) cdgConsolidado[sk].desc = it.desc;
          } else {
            cdgConsolidado[sk] = { sku: it.sku, desc: it.desc||'', qty: qty, costo: null, orden: ordenSku.length };
            ordenSku.push(sk);
          }
        });
      }

      // ── Remap fisico existente al nuevo índice consolidado por SKU ────────
      // fisMapeado usa viejoItems (capturado antes de mutar) y viejoFisico.
      // fisMapeado: SK_UPPER → { fisico, daniado, quien, ts, cobertura }
      var fisMapeado = {};
      if(Array.isArray(viejoFisico) && viejoFisico.length && viejoItems.length) {
        viejoItems.forEach(function(it, idx){
          var sv = viejoFisico[idx];
          if(!sv || sv.fisico === null || sv.fisico === undefined) return;
          var sk = String(it.sku || '').trim().toUpperCase();
          var fv = Number(sv.fisico);
          var dv = Number(sv.daniado || 0);
          if(fisMapeado[sk]) {
            fisMapeado[sk].fisico  += fv;
            fisMapeado[sk].daniado += dv;
            // Conservar último quien/ts no vacío
            if(sv.quien) fisMapeado[sk].quien = sv.quien;
            if(sv.ts)    fisMapeado[sk].ts    = sv.ts;
          } else {
            fisMapeado[sk] = {
              fisico:   fv,
              daniado:  dv,
              quien:    sv.quien    || '',
              ts:       sv.ts       || '',
              cobertura: sv.cobertura || 'En revisión',
              calcExpr: sv.calcExpr || null
            };
          }
        });
      }

      // ── Cruzar consolidado con WMS ────────────────────────────────────────
      var skusCDG = {};
      var nuevosItems = ordenSku.map(function(sk){
        skusCDG[sk] = true;
        var base = cdgConsolidado[sk];
        var desc = base.desc || wmsDescMap[sk] || '';
        var item = {
          sku:  base.sku,
          desc: desc,
          qty:  base.qty   // Validado CDG consolidado
        };
        if(base.costo != null) item.costo = base.costo;
        if(wmsMap[sk] !== undefined) {
          item.teoricoWMS = wmsMap[sk];
        }
        if(base.sku && base.sku.trim().toUpperCase() !== sk) {
          // Normalizar sku al original del CDG
        }
        item.raw = { origen: 'CDG', status: 'CDG Validado' };
        return item;
      });

      // SKUs solo-WMS (en WMS pero no en CDG)
      Object.keys(wmsMap).forEach(function(sk){
        if(skusCDG[sk]) return;
        nuevosItems.push({
          sku: sk, desc: wmsDescMap[sk] || '', qty: 0,
          teoricoWMS: wmsMap[sk],
          raw: { origen: 'CDG', status: 'CDG Validado' },
          _soloWMS: true
        });
      });

      // ── Actualizar teorico ────────────────────────────────────────────────
      teo.items = nuevosItems;

      // ── Actualizar fisico si hay remap disponible ─────────────────────────
      var tieneFisicoRemap = Object.keys(fisMapeado).length > 0;
      if(tieneFisicoRemap) {
        // Construir nuevo array fisico alineado a nuevosItems (excluye _soloWMS)
        var nuevoFisico = nuevosItems.map(function(item){
          if(item._soloWMS) return null; // SKU solo-WMS: sin físico
          var sk = String(item.sku || '').trim().toUpperCase();
          var fm = fisMapeado[sk];
          if(!fm) return null; // sin físico registrado para este SKU
          return {
            fisico:    fm.fisico,
            daniado:   fm.daniado,
            quien:     fm.quien,
            ts:        fm.ts,
            cobertura: fm.cobertura,
            calcExpr:  fm.calcExpr,
            lastAt:    Date.now()
          };
        });
        state.fisico[teoKey] = nuevoFisico;
      }
      // Si no hay fisico previo con datos: no tocar state.fisico[teoKey]

      actualizadas.push(lic + (teoKey !== lic ? '→'+teoKey : '')
        + ' (' + nuevosItems.filter(function(i){ return !i._soloWMS; }).length + ' SKUs consolidados)');
      } // fin for wi (pageRows)

      if(pageRows.length < PAGE) break; // última página
    } // fin for pg (páginas)

    if(!huboWms) {
      return res.json({ ok: true, actualizadas: [], sinConteo: [], msg: 'No hay WMS cargados.' });
    }

    // ── Persistir ────────────────────────────────────────────────────────
    if(actualizadas.length > 0) {
      state.version++;
      try {
        await withTimeout(
          dbSet('daily_state', buildDailyStatePayload()),
          15000, 'CDG WMS sincronizar-hamilton save'
        );
        console.log('CDG WMS sincronizar-hamilton OK:', actualizadas.length, 'licencias,', usuario);
      } catch(saveErr) {
        // Rollback completo: restaurar teorico, fisico, version e historial
        state.teorico = snapTeorico;
        state.fisico  = snapFisico;
        state.version = snapVersion;
        if(state.historial) state.historial.length = snapHistorial;
        console.log('CDG WMS sincronizar-hamilton ROLLBACK:', saveErr.message);
        return res.status(500).json({
          ok: false,
          error: 'No se guardó. Los datos en memoria fueron restaurados. Reintentá. (' + saveErr.message + ')'
        });
      }
    }

    res.json({
      ok:              true,
      actualizadas:    actualizadas,
      sinConteo:       sinConteo,
      advertencias:    advertencias.length ? advertencias : undefined,
      totalProcesadas: totalProcesadas
    });

  } catch(e) {
    console.log('CDG WMS sincronizar-hamilton FAILED:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ── cdgCargarWmsDesdeExcel ────────────────────────────────────────────────
// Helper que encapsula todo el parseo y guardado del WMS desde Excel.
// FIX (mar 2-jun-2026, server v20): extraído para reutilizar en el endpoint
// general POST /api/cdg/wms/:id (sin cdg_meta) y en POST /api/cdg/v2/:id/wms.
// Parámetros:
//   licenciaId: ya normalizado con cdgNormId
//   req: request de Express (req.file, req.body.usuario)
//   meta: cdg_meta ya cargado (o null si no existe)
//   res: response de Express
async function cdgCargarWmsDesdeExcel(licenciaId, req, meta, res) {
  var usuario = String(req.body && req.body.usuario ? req.body.usuario : '').trim();
  if(!usuario)  return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    // ── Parsear Excel ────────────────────────────────────────────────────
    var wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: 'buffer', raw: false, cellDates: true });
    } catch(e) {
      return res.status(400).json({ ok: false, error: 'El archivo no es un Excel válido. (' + e.message + ')' });
    }

    var advertencia = null;
    var sheetName = wb.SheetNames.find(function(n) { return n.trim().toLowerCase() === 'traslados'; });
    if(!sheetName) {
      sheetName = wb.SheetNames[0];
      advertencia = 'Hoja "Traslados" no encontrada. Se usó la primera hoja: "' + sheetName + '". Verificá que estás cargando el archivo WMS correcto.';
    }
    var ws   = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if(rows.length < 2) return res.status(400).json({ ok: false, error: 'El archivo está vacío o solo tiene encabezado.' });

    // ── Detectar columnas ────────────────────────────────────────────────
    var hdr     = rows[0].map(cdgNormHdr);
    var cNum    = hdr.indexOf('numero');
    var cSku    = hdr.indexOf('sku');
    var cNombre = hdr.indexOf('nombre');
    if(cNombre < 0) cNombre = hdr.indexOf('descripcion');
    var cCant   = hdr.indexOf('cant');
    var cUnis   = hdr.indexOf('unidades');
    var cFecha  = hdr.indexOf('fecha');
    var cOrigen = hdr.indexOf('origen');
    var cDest   = hdr.indexOf('destino');
    var cTipo   = hdr.indexOf('tipo');
    var cStatus = hdr.indexOf('status');
    var cLineas = hdr.indexOf('lineas');

    var faltantes = [];
    if(cNum  < 0) faltantes.push('"Número"');
    if(cSku  < 0) faltantes.push('"SKU"');
    if(cCant < 0 && cUnis < 0) faltantes.push('"Cant." o "Unidades"');
    if(faltantes.length > 0) {
      return res.status(400).json({
        ok: false,
        error: 'Columnas requeridas no encontradas: ' + faltantes.join(', ') + '. '
             + 'Encabezados detectados: ' + rows[0].join(', ') + '. '
             + 'Verificá que estás cargando el archivo WMS correcto (hoja "Traslados" del export WMS).'
      });
    }

    // ── Filtrar por licencia ─────────────────────────────────────────────
    var licNorm = cdgNormLicWMS(licenciaId);
    var filasFiltradas = rows.slice(1).filter(function(r) {
      return cdgNormLicWMS(r[cNum]) === licNorm;
    });
    if(filasFiltradas.length === 0) {
      var licEncontradas = [];
      var vistas = {};
      for(var i = 1; i < rows.length && licEncontradas.length < 10; i++) {
        var v = cdgNormLicWMS(rows[i][cNum]);
        if(v && !vistas[v]) { licEncontradas.push(v); vistas[v] = true; }
      }
      return res.status(400).json({
        ok: false,
        error: 'El archivo no contiene filas para la licencia "' + licenciaId + '". '
             + 'Licencias encontradas: ' + licEncontradas.join(', ')
             + (licEncontradas.length === 10 ? ' (y más…)' : '') + '.'
      });
    }

    // ── Extraer encabezado ───────────────────────────────────────────────
    var f0 = filasFiltradas[0];
    var encabezado = {
      fecha:     f0[cFecha]  || null,
      origen:    cOrigen  >= 0 ? (f0[cOrigen]  || null) : null,
      destino:   cDest    >= 0 ? (f0[cDest]    || null) : null,
      tipo:      cTipo    >= 0 ? (f0[cTipo]    || null) : null,
      status:    cStatus  >= 0 ? (f0[cStatus]  || null) : null,
      lineasWMS: cLineas  >= 0 ? Number(f0[cLineas] || 0) : 0
    };

    // ── Procesar filas: parsear + deduplicar ─────────────────────────────
    var skuMap = {}, skuOrder = [];
    filasFiltradas.forEach(function(r) {
      var sku = String(r[cSku] || '').trim();
      if(!sku) return;
      var qty = cdgParseQty(cCant >= 0 ? r[cCant] : null);
      if(qty === null) qty = cdgParseQty(cUnis >= 0 ? r[cUnis] : null);
      if(qty === null) return;
      var desc = String(cNombre >= 0 ? (r[cNombre] || '') : '').trim();
      if(skuMap[sku]) {
        skuMap[sku].cantidad += qty;
        if(!skuMap[sku].descripcion && desc) skuMap[sku].descripcion = desc;
      } else {
        skuMap[sku] = { descripcion: desc, cantidad: qty };
        skuOrder.push(sku);
      }
    });

    // ── Enriquecer descripciones vacías desde sku_catalog ────────────────
    var skusSinDesc = skuOrder.filter(function(s) { return !skuMap[s].descripcion; });
    if(skusSinDesc.length > 0 && SUPABASE_URL && SUPABASE_KEY) {
      var enrichWarnings = [];
      for(var ci = 0; ci < skusSinDesc.length; ci += 200) {
        var chunk = skusSinDesc.slice(ci, ci + 200);
        try {
          var catalogRows = await supabase('GET', 'sku_catalog', null,
            '?sku=in.(' + chunk.map(function(s){ return encodeURIComponent(s); }).join(',') + ')&select=sku,descripcion');
          if(Array.isArray(catalogRows)) {
            catalogRows.forEach(function(row) {
              if(skuMap[row.sku] && !skuMap[row.sku].descripcion && row.descripcion)
                skuMap[row.sku].descripcion = row.descripcion;
            });
          }
        } catch(enrichErr) {
          enrichWarnings.push('sku_catalog chunk ' + ci + ': ' + enrichErr.message);
        }
      }
      if(enrichWarnings.length) advertencia = (advertencia ? advertencia + ' | ' : '') + enrichWarnings.join(' | ');
    }

    // ── Construir array final + UPSERT ───────────────────────────────────
    var skusArray = skuOrder.map(function(sku) {
      return { sku: sku, descripcion: skuMap[sku].descripcion || '', cantidad: skuMap[sku].cantidad };
    });
    var totalSkus     = skusArray.length;
    var totalUnidades = skusArray.reduce(function(a, s){ return a + s.cantidad; }, 0);
    var now           = new Date().toISOString();
    var nombreArchivo = req.file.originalname || 'archivo.xlsx';

    await withTimeout(cdgUpsertWms(licenciaId, {
      licencia_id:    licenciaId,
      skus:           JSON.stringify(skusArray),
      encabezado:     JSON.stringify(encabezado),
      cargado_por:    usuario,
      ts_carga:       now,
      nombre_archivo: nombreArchivo,
      total_skus:     totalSkus,
      total_unidades: totalUnidades
    }), 15000, 'CDG WMS upsert');

    // ── Actualizar meta.wms si existe cdg_meta ────────────────────────────
    // Si meta es null (v1 o solo-WMS), no intentar cdgSaveMeta.
    var metaWarning = null;
    if(meta) {
      meta.wms = { cargadoPor: usuario, tsCarga: now, nombreArchivo, totalSkus, totalUnidades };
      cdgBitacora(meta, usuario, 'wms_cargado', nombreArchivo + ' — ' + totalSkus + ' SKUs, ' + totalUnidades + ' unidades');
      try {
        await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG WMS meta save');
      } catch(metaErr) {
        try {
          await withTimeout(cdgSaveMeta(licenciaId, meta), 10000, 'CDG WMS meta save retry');
        } catch(retryErr) {
          metaWarning = 'WMS guardado. No se pudo actualizar caché de metadata (' + retryErr.message + '). El WMS estará disponible igual.';
          console.log('CDG WMS meta save FAILED after retry:', retryErr.message);
        }
      }
    }

    var respAdvertencia = [advertencia, metaWarning].filter(Boolean).join(' | ') || undefined;
    console.log('CDG WMS cargado:', licenciaId, 'por', usuario, '—', totalSkus, 'SKUs,', totalUnidades, 'unidades');
    res.json({
      ok: true, totalSkus, totalUnidades, tsCarga: now,
      ...(respAdvertencia ? { advertencia: respAdvertencia } : {})
    });

  } catch(e) {
    console.log('CDG WMS upload FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'Error al procesar WMS: ' + e.message });
  }
}

// ── POST /api/cdg/wms/:id ─────────────────────────────────────────────────
// Endpoint general WMS: NO exige cdg_meta v2.
// FIX (mar 2-jun-2026, server v20): permite cargar WMS para v1, v2 y solo-WMS.
// Si existe cdg_meta y estado === 'cerrado' → rechaza.
// Si existe cdg_meta abierta → guarda WMS + actualiza meta.wms.
// Si NO existe cdg_meta → guarda solo en cdg_wms (sin meta).
app.post('/api/cdg/wms/:id', upload.single('file'), async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    var meta = null;
    try { meta = await cdgGetMeta(licenciaId); } catch(e) { /* no existe como v2 */ }
    if(meta && meta.estado === 'cerrado') {
      return res.status(400).json({ ok: false, error: 'La licencia está cerrada. No se puede cargar WMS.' });
    }
    await cdgCargarWmsDesdeExcel(licenciaId, req, meta, res);
  } catch(e) {
    console.log('CDG WMS POST general FAILED:', e.message);
    if(!res.headersSent) res.status(500).json({ ok: false, error: e.message });
  }
});

// ── POST /api/cdg/v2/:id/wms ─────────────────────────────────────────────
// Mantiene compatibilidad con el cliente existente de captura v2.
// Ahora delega a cdgCargarWmsDesdeExcel (misma lógica, sin código duplicado).
// FIX (mar 2-jun-2026, server v20): ya no duplica la lógica de parseo.
app.post('/api/cdg/v2/:id/wms', upload.single('file'), async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia v2 no encontrada: ' + licenciaId });
    if(meta.estado === 'cerrado') {
      return res.status(400).json({ ok: false, error: 'La licencia está cerrada. No se puede cargar WMS.' });
    }
    await cdgCargarWmsDesdeExcel(licenciaId, req, meta, res);
  } catch(e) {
    console.log('CDG v2 WMS upload FAILED:', e.message);
    if(!res.headersSent) res.status(500).json({ ok: false, error: e.message });
  }
});

// ── cdgResponderWms: helper compartido para devolver WMS de cdg_wms ─────────
// No exige cdg_meta v2. Funciona para v1, v2 y licencias solo-WMS.
async function cdgResponderWms(licenciaId, res) {
  var wms = await cdgGetWms(licenciaId);
  if(!wms) return res.json({ ok: true, skus: [], wmsMeta: null, encabezado: null });
  var skus      = typeof wms.skus       === 'string' ? JSON.parse(wms.skus)       : (wms.skus       || []);
  var encabezado = typeof wms.encabezado === 'string' ? JSON.parse(wms.encabezado) : (wms.encabezado || null);
  res.json({
    ok: true, skus: skus, encabezado: encabezado,
    wmsMeta: {
      cargadoPor:    wms.cargado_por,
      tsCarga:       wms.ts_carga,
      nombreArchivo: wms.nombre_archivo,
      totalSkus:     wms.total_skus,
      totalUnidades: wms.total_unidades
    }
  });
}

// ── GET /api/cdg/wms/:id ───────────────────────────────────────────────────
// Endpoint general WMS: no exige cdg_meta v2.
// Sirve para v1 clásico, v2 colaborativo y licencias "solo WMS".
// FIX (mar 2-jun-2026, server v20): nuevo endpoint sin restricción de v2.
app.get('/api/cdg/wms/:id', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    await cdgResponderWms(licenciaId, res);
  } catch(e) {
    console.log('CDG WMS GET FAILED:', licenciaId, e.message);
    res.status(500).json({ ok: false, error: 'Error al obtener WMS: ' + e.message });
  }
});

// ── GET /api/cdg/v2/:id/wms ────────────────────────────────────────────────
// Mantiene compatibilidad con el cliente existente de captura v2.
// Delega a cdgResponderWms (ya no exige cdg_meta).
// FIX (mar 2-jun-2026, server v20): eliminada la guardia de cdg_meta para
// no bloquear v1/solo-WMS. Si el caller lo necesita solo para v2, funciona igual.
app.get('/api/cdg/v2/:id/wms', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    await cdgResponderWms(licenciaId, res);
  } catch(e) {
    console.log('CDG v2 WMS GET FAILED:', licenciaId, e.message);
    res.status(500).json({ ok: false, error: 'Error al obtener WMS: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════════════════
// BOD MODULE START — Bodega CDG / Armado de Tarimas — Fase 1
// FIX (mar 2-jun-2026, server v20): módulo completamente aislado.
// Flag: BOD_ENABLED=true en variables de entorno de Render para activar.
// Con BOD_ENABLED=false (default), todos los endpoints responden 503.
// No toca: Hamilton, CDG, app_state, cdg_lineas, cdg_wms, sku_catalog,
//          teorico, fisico, endpoints /api/cdg/*, /api/state.
// ══════════════════════════════════════════════════════════════════════════

var BOD_ENABLED = process.env.BOD_ENABLED === 'true';

// Middleware guard — aplica a todos los endpoints /api/bod/*
function bodGuard(req, res, next) {
  if(!BOD_ENABLED) return res.status(503).json({ ok: false, error: 'Módulo bodega no habilitado.' });
  next();
}

// Helper: normalizar nombre de tarima → "A1", "B12", etc.
// "a 1" → "A1", "A-1" → "A1", "a1" → "A1"
function bodNormTarima(v) {
  if(!v) return '';
  return String(v).trim().toUpperCase().replace(/[\s\-_]+/g, '');
}

// Helper: leer bdg_barra_sku desde Supabase
async function bodGetBarra(barra) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var rows = await supabase('GET', 'bod_barra_sku', null,
    '?barra=eq.' + encodeURIComponent(String(barra).trim()) + '&limit=1');
  return (Array.isArray(rows) && rows.length) ? rows[0] : null;
}

// ── GET /api/bod/status ────────────────────────────────────────────────────
// Permite al cliente saber si el módulo está habilitado.
app.get('/api/bod/status', function(req, res) {
  res.json({ ok: true, enabled: BOD_ENABLED });
});

// ── POST /api/bod/sesion ───────────────────────────────────────────────────
// Crear o devolver sesión existente (idempotente por licencia_id+fecha+tipo).
app.post('/api/bod/sesion', bodGuard, async (req, res) => {
  try {
    var { licencia_id, fecha_trabajo, tipo, creado_por, notas } = req.body || {};
    licencia_id  = String(licencia_id  || '').trim().toUpperCase();
    fecha_trabajo = String(fecha_trabajo || '').trim();
    tipo          = String(tipo          || 'recoleccion').trim();
    creado_por    = String(creado_por    || '').trim();
    if(!licencia_id)   return res.status(400).json({ ok:false, error:'falta licencia_id' });
    if(!fecha_trabajo) return res.status(400).json({ ok:false, error:'falta fecha_trabajo (YYYY-MM-DD)' });
    if(!creado_por)    return res.status(400).json({ ok:false, error:'falta creado_por' });

    // Buscar existente
    var existing = await supabase('GET', 'bod_sesiones', null,
      '?licencia_id=eq.' + encodeURIComponent(licencia_id)
      + '&fecha_trabajo=eq.' + encodeURIComponent(fecha_trabajo)
      + '&tipo=eq.'          + encodeURIComponent(tipo)
      + '&limit=1');
    if(Array.isArray(existing) && existing.length) {
      return res.json({ ok:true, sesion: existing[0], created: false });
    }

    // Crear nueva
    var newId = 'bod-' + Date.now() + '-' + Math.random().toString(36).slice(2,7);
    var sesion = {
      id: newId, licencia_id, tipo, estado: 'abierta',
      creado_por, modificado_por: creado_por,
      fecha_trabajo, ts_creacion: new Date().toISOString(),
      notas: notas || null
    };
    var created = await supabase('POST', 'bod_sesiones', sesion, '');
    var row = Array.isArray(created) ? created[0] : sesion;
    res.json({ ok:true, sesion: row, created: true });
  } catch(e) {
    console.log('BOD sesion POST FAILED:', e.message);
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion?licencia_id=X&fecha=Y ─────────────────────────────
// Buscar sesión por licencia + fecha en lugar del ID interno.
// Devuelve el mismo formato que GET /api/bod/sesion/:id.
app.get('/api/bod/sesion', bodGuard, async (req, res) => {
  try {
    var lic   = String(req.query.licencia_id || '').trim().toUpperCase();
    var fecha = String(req.query.fecha       || '').trim();
    var tipo  = String(req.query.tipo        || 'recoleccion').trim();
    if(!lic || !fecha) return res.status(400).json({ ok:false, error:'falta licencia_id o fecha' });

    var sesRows = await supabase('GET', 'bod_sesiones', null,
      '?licencia_id=eq.' + encodeURIComponent(lic)
      + '&fecha_trabajo=eq.' + encodeURIComponent(fecha)
      + '&tipo=eq.' + encodeURIComponent(tipo)
      + '&order=ts_creacion.desc&limit=1');
    if(!Array.isArray(sesRows) || !sesRows.length) {
      return res.status(404).json({ ok:false, error:'No se encontró sesión para la licencia "'+lic+'" en la fecha '+fecha+'.' });
    }
    var sesion = sesRows[0];
    var lineas = await supabase('GET', 'bod_lineas', null,
      '?sesion_id=eq.' + encodeURIComponent(sesion.id) + '&eliminada=eq.false&order=ts_captura.asc');
    var lineArr = Array.isArray(lineas) ? lineas : [];
    var porTarima = {};
    lineArr.forEach(function(l) {
      var t = l.tarima || '?';
      if(!porTarima[t]) porTarima[t] = { tarima:t, lineas:0, unidades:0 };
      porTarima[t].lineas++;
      porTarima[t].unidades += Number(l.cantidad)||0;
    });
    res.json({ ok:true, sesion, lineas: lineArr, resumenTarimas: Object.values(porTarima) });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion/:id ────────────────────────────────────────────────
// Devolver sesión + líneas vigentes + resumen por tarima.
app.get('/api/bod/sesion/:id', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var rows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(rows) || !rows.length) return res.status(404).json({ ok:false, error:'Sesión no encontrada' });
    var sesion = rows[0];
    var lineas = await supabase('GET', 'bod_lineas', null,
      '?sesion_id=eq.' + encodeURIComponent(sesId) + '&eliminada=eq.false&order=ts_captura.asc');
    var lineArr = Array.isArray(lineas) ? lineas : [];
    // Resumen por tarima
    var porTarima = {};
    lineArr.forEach(function(l) {
      var t = l.tarima || '?';
      if(!porTarima[t]) porTarima[t] = { tarima:t, lineas:0, unidades:0 };
      porTarima[t].lineas++;
      porTarima[t].unidades += Number(l.cantidad)||0;
    });
    res.json({ ok:true, sesion, lineas: lineArr, resumenTarimas: Object.values(porTarima) });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/sesion/:id/linea ─────────────────────────────────────────
app.post('/api/bod/sesion/:id/linea', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var { tarima, barra, sku, descripcion, cantidad, operador } = req.body || {};
    tarima     = bodNormTarima(tarima);
    barra      = String(barra      || '').trim() || null;
    sku        = String(sku        || '').trim();
    descripcion = String(descripcion || '').trim();
    cantidad   = Number(cantidad);
    operador   = String(operador   || '').trim();
    if(!tarima)   return res.status(400).json({ ok:false, error:'falta tarima' });
    if(!sku)      return res.status(400).json({ ok:false, error:'falta sku' });
    if(!operador) return res.status(400).json({ ok:false, error:'falta operador' });
    if(!(cantidad > 0)) return res.status(400).json({ ok:false, error:'cantidad debe ser mayor a 0' });

    // Verificar sesión existe y está abierta
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'Sesión no encontrada' });
    if(sesRows[0].estado === 'cerrada') return res.status(400).json({ ok:false, error:'La sesión está cerrada.' });

    // Completar descripción desde sku_catalog si vacía
    var advertencia = null;
    if(!descripcion && SUPABASE_URL && SUPABASE_KEY) {
      try {
        var catRows = await supabase('GET', 'sku_catalog', null,
          '?sku=eq.' + encodeURIComponent(sku) + '&select=sku,descripcion&limit=1');
        if(Array.isArray(catRows) && catRows.length && catRows[0].descripcion)
          descripcion = catRows[0].descripcion;
      } catch(e) { advertencia = 'No se pudo completar descripción desde catálogo.'; }
    }

    var now   = new Date().toISOString();
    var linId = 'bl-' + Date.now() + '-' + Math.random().toString(36).slice(2,7);
    var linea = {
      id: linId, sesion_id: sesId, licencia_id: sesRows[0].licencia_id,
      tarima, barra, sku, descripcion, cantidad, operador,
      ts_captura: now, ts_modif: now, eliminada: false,
      auditado: false, auditado_por: null, ts_auditado: null, cantidad_audit: null
    };
    var created = await supabase('POST', 'bod_lineas', linea, '');
    var row = Array.isArray(created) ? created[0] : linea;
    res.json({ ok:true, linea: row, ...(advertencia ? { advertencia } : {}) });
  } catch(e) {
    console.log('BOD linea POST FAILED:', e.message);
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── PATCH /api/bod/linea/:lineaId ─────────────────────────────────────────
app.patch('/api/bod/linea/:lineaId', bodGuard, async (req, res) => {
  try {
    var linId   = String(req.params.lineaId).trim();
    var { cantidad, sku, descripcion, tarima, usuario } = req.body || {};
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Línea no encontrada' });
    var lin = rows[0];
    // Permisos: mismo operador o supervisor (caller indica con supervisor:true)
    var esSup = req.body && req.body.supervisor === true;
    if(lin.operador !== usuario && !esSup)
      return res.status(403).json({ ok:false, error:'Solo el operador original o supervisor puede editar.' });
    var patch = { ts_modif: new Date().toISOString() };
    if(cantidad !== undefined) {
      if(!(Number(cantidad) > 0)) return res.status(400).json({ ok:false, error:'cantidad debe ser mayor a 0' });
      patch.cantidad = Number(cantidad);
    }
    if(sku         !== undefined) patch.sku         = String(sku).trim();
    if(descripcion !== undefined) patch.descripcion = String(descripcion).trim();
    if(tarima      !== undefined) patch.tarima      = bodNormTarima(tarima);
    await supabase('PATCH', 'bod_lineas', patch, '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true, patch });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── DELETE /api/bod/linea/:lineaId ────────────────────────────────────────
app.delete('/api/bod/linea/:lineaId', bodGuard, async (req, res) => {
  try {
    var linId   = String(req.params.lineaId).trim();
    var { usuario, supervisor } = req.body || {};
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Línea no encontrada' });
    var lin = rows[0];
    var esSup = supervisor === true;
    if(lin.operador !== usuario && !esSup)
      return res.status(403).json({ ok:false, error:'Solo el operador original o supervisor puede eliminar.' });
    await supabase('PATCH', 'bod_lineas', { eliminada:true, ts_modif:new Date().toISOString() },
      '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion/:id/lineas ────────────────────────────────────────
app.get('/api/bod/sesion/:id/lineas', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var desde  = req.query.desde || null;
    var inclEl = req.query.incluir_eliminadas === 'true';
    var q = '?sesion_id=eq.' + encodeURIComponent(sesId) + '&order=ts_captura.asc';
    if(!inclEl) q += '&eliminada=eq.false';
    if(desde)   q += '&ts_modif=gte.' + encodeURIComponent(desde);
    var lineas = await supabase('GET', 'bod_lineas', null, q);
    res.json({ ok:true, lineas: Array.isArray(lineas) ? lineas : [] });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/barra/:barra ────────────────────────────────────────────
app.get('/api/bod/barra/:barra', bodGuard, async (req, res) => {
  try {
    var barra = String(req.params.barra).trim();
    var row   = await bodGetBarra(barra);
    if(!row) return res.json({ ok:true, encontrado:false });
    // Completar descripción desde sku_catalog si vacía
    if(!row.descripcion && SUPABASE_URL && SUPABASE_KEY) {
      try {
        var cat = await supabase('GET', 'sku_catalog', null,
          '?sku=eq.' + encodeURIComponent(row.sku) + '&select=sku,descripcion&limit=1');
        if(Array.isArray(cat) && cat.length && cat[0].descripcion)
          row = Object.assign({}, row, { descripcion: cat[0].descripcion });
      } catch(e) { /* no bloquear */ }
    }
    res.json({ ok:true, encontrado:true, barra: row.barra, sku: row.sku, descripcion: row.descripcion || '' });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/barra-sku/upload ────────────────────────────────────────
app.post('/api/bod/barra-sku/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    if(!req.file) return res.status(400).json({ ok:false, error:'falta archivo' });

    // ── Parsear CSV o XLSX con XLSX (maneja ambos formatos) ───────────────
    var wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });
    if(rows.length < 2) return res.status(400).json({ ok:false, error:'Archivo vacío o solo encabezado.' });

    // ── Detectar columnas ─────────────────────────────────────────────────
    var hdr   = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var cBarra = hdr.findIndex(function(h){ return h==='barra'||h.includes('barcode')||h.includes('codigo'); });
    var cSku   = hdr.findIndex(function(h){ return h==='sku'||h.includes('articulo')||h.includes('artículo'); });
    var cDesc  = hdr.findIndex(function(h){ return h.includes('desc')||h.includes('nombre'); });
    if(cBarra < 0 || cSku < 0) {
      return res.status(400).json({
        ok:false,
        error:'No se encontraron columnas Barra y SKU. Encabezados detectados: '+rows[0].join(', ')
      });
    }

    // ── Deduplicar en memoria ─────────────────────────────────────────────
    // Reglas: ignorar barra vacía, sku vacío, barra con < 6 caracteres.
    // Si barra aparece duplicada, conservar la última fila.
    var dedup = {};        // barra → { barra, sku, descripcion }
    var omitidas = 0;
    var now = new Date().toISOString();

    for(var ri = 1; ri < rows.length; ri++) {
      var r    = rows[ri];
      var bar  = String(r[cBarra]||'').trim();
      var sku  = String(r[cSku]  ||'').trim();
      var desc = cDesc >= 0 ? String(r[cDesc]||'').trim() : '';
      if(!bar || !sku || bar.length < 6) { omitidas++; continue; }
      dedup[bar] = { barra:bar, sku:sku, descripcion:desc, ts_actualizacion:now };
    }

    var unicas  = Object.keys(dedup).length;
    var registros = Object.values(dedup);

    if(!unicas) {
      return res.status(400).json({
        ok:false,
        error:'No se encontraron filas válidas. Revisá que las columnas barra y sku tengan datos.',
        omitidas: omitidas
      });
    }

    // ── Upsert por chunks de 500 ──────────────────────────────────────────
    // PostgREST acepta arrays en el body + on_conflict=barra para upsert masivo.
    var CHUNK = 500;
    var procesadas = 0;
    var errores    = [];
    var chunks     = 0;
    var startedAt    = Date.now();
    var MAX_UPLOAD_MS = 22000;

    for(var ci = 0; ci < registros.length; ci += CHUNK) {
      if(Date.now() - startedAt > MAX_UPLOAD_MS) {
        errores.push({
          chunk: chunks + 1,
          error: 'Upload detenido antes del timeout HTTP. Reintentar para completar chunks pendientes.'
        });
        break;
      }
      var batch = registros.slice(ci, ci + CHUNK);
      chunks++;
      try {
        await withTimeout(
          supabase('POST', 'bod_barra_sku', batch, '?on_conflict=barra'),
          8000,
          'barra-sku chunk '+chunks
        );
        procesadas += batch.length;
      } catch(e) {
        console.log('BOD barra-sku chunk', chunks, 'falló:', e.message);
        errores.push({ chunk: chunks, error: e.message });
      }
    }

    console.log('BOD barra-sku upload OK: unicas='+unicas+' procesadas='+procesadas+' chunks='+chunks);
    res.json({
      ok:        true,
      procesadas: procesadas,
      unicas:     unicas,
      omitidas:   omitidas,
      chunks:     chunks,
      ...(errores.length ? { errores: errores } : {})
    });

  } catch(e) {
    console.log('BOD barra-sku upload FAILED:', e.message);
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/linea/:lineaId/auditar ─────────────────────────────────
app.post('/api/bod/linea/:lineaId/auditar', bodGuard, async (req, res) => {
  try {
    var linId = String(req.params.lineaId).trim();
    var { usuario, supervisor, auditor, cantidad_audit } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden auditar.' });
    if(cantidad_audit === undefined || cantidad_audit === null || !(Number(cantidad_audit) > 0))
      return res.status(400).json({ ok:false, error:'cantidad_audit debe ser mayor a 0.' });
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Línea no encontrada' });
    var now = new Date().toISOString();
    var patch = { auditado:true, auditado_por:usuario, ts_auditado:now, ts_modif:now,
                  cantidad_audit: Number(cantidad_audit) };
    await supabase('PATCH', 'bod_lineas', patch, '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/linea/:lineaId/unauditar ────────────────────────────────
// Revierte una línea auditada a estado pendiente.
app.post('/api/bod/linea/:lineaId/unauditar', bodGuard, async (req, res) => {
  try {
    var linId = String(req.params.lineaId).trim();
    var { usuario, supervisor, auditor } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden anular auditorías.' });
    var now = new Date().toISOString();
    await supabase('PATCH', 'bod_lineas',
      { auditado:false, auditado_por:null, ts_auditado:null, cantidad_audit:null, ts_modif:now },
      '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/sesion/:id/cerrar ──────────────────────────────────────
app.post('/api/bod/sesion/:id/cerrar', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var { usuario, supervisor, auditor } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden cerrar sesiones.' });
    var rows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Sesión no encontrada' });
    await supabase('PATCH', 'bod_sesiones',
      { estado:'cerrada', ts_cierre:new Date().toISOString(), modificado_por:usuario },
      '?id=eq.' + encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion/:id/furgones ─────────────────────────────────────
app.get('/api/bod/sesion/:id/furgones', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var rows = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=furgon.asc,tarima.asc');
    res.json({ ok:true, asignaciones: Array.isArray(rows) ? rows : [] });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── POST /api/bod/sesion/:id/furgon ──────────────────────────────────────
app.post('/api/bod/sesion/:id/furgon', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var { tarimas, furgon, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden asignar furgones.' });
    if(!furgon)    return res.status(400).json({ ok:false, error:'falta furgon' });
    if(!Array.isArray(tarimas)||!tarimas.length) return res.status(400).json({ ok:false, error:'tarimas debe ser array no vacío' });
    var normadas = tarimas.map(function(t){ return bodNormTarima(t); }).filter(Boolean);
    if(!normadas.length) return res.status(400).json({ ok:false, error:'ninguna tarima válida' });
    furgon = String(furgon).trim();

    // Leer asignaciones existentes de la sesión
    var existing = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId));
    var asigMap = {};
    if(Array.isArray(existing)) existing.forEach(function(r){ asigMap[r.tarima] = r.furgon; });

    // Detectar bloqueadas (asignadas a otro furgón distinto)
    var bloqueadas = normadas.filter(function(t){ return asigMap[t] && asigMap[t] !== furgon; });
    if(bloqueadas.length) return res.status(409).json({ ok:false, error:'Algunas tarimas ya están asignadas.', bloqueadas: bloqueadas });

    // Insertar solo las nuevas
    var nuevas = normadas.filter(function(t){ return !asigMap[t]; });
    var now = new Date().toISOString();
    for(var i=0; i<nuevas.length; i++) {
      await supabase('POST', 'bod_tarima_furgon',
        {
          id: 'btf-' + Date.now() + '-' + i + '-' + Math.random().toString(36).slice(2,7),
          sesion_id: sesId,
          tarima: nuevas[i],
          furgon: furgon,
          asignado_por: usuario || '',
          ts_asig: now
        },
        '?on_conflict=sesion_id,tarima');
    }
    res.json({ ok:true, asignadas: normadas.length, nuevas: nuevas.length });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── DELETE /api/bod/sesion/:id/furgon/:tarima ─────────────────────────────
app.delete('/api/bod/sesion/:id/furgon/:tarima', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var tarima  = bodNormTarima(req.params.tarima);
    var { usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden quitar asignaciones.' });
    await supabase('DELETE', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&tarima=eq.'+encodeURIComponent(tarima));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion/:id/manifiesto?furgon=N ──────────────────────────
// Igual que antes + agrega estado_cierre desde bod_furgon_cierres
app.get('/api/bod/sesion/:id/manifiesto', bodGuard, async (req, res) => {
  try {
    var sesId        = String(req.params.id).trim();
    var filtroFurgon = req.query.furgon ? String(req.query.furgon).trim() : null;

    var [lineas, asignaciones, cierres] = await Promise.all([
      supabase('GET', 'bod_lineas', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&eliminada=eq.false&order=ts_captura.asc'),
      supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)),
      supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)).catch(function(){ return []; })
    ]);
    lineas       = Array.isArray(lineas)       ? lineas       : [];
    asignaciones = Array.isArray(asignaciones) ? asignaciones : [];
    cierres      = Array.isArray(cierres)      ? cierres      : [];

    // Mapa tarima → furgon y cierre → datos
    var tarimaFurgon = {};
    asignaciones.forEach(function(a){ tarimaFurgon[a.tarima] = a.furgon; });
    var cierreMap = {};
    cierres.forEach(function(c){ cierreMap[c.furgon] = c; });

    var tarimasSinAsignar = [];
    var tarimasConLineas = {};
    lineas.forEach(function(l){ tarimasConLineas[l.tarima] = true; });
    Object.keys(tarimasConLineas).forEach(function(t){
      if(!tarimaFurgon[t]) tarimasSinAsignar.push(t);
    });

    var porFurgon = {};
    var tarimasFurgon = {};

    lineas.forEach(function(l){
      var furgon = tarimaFurgon[l.tarima];
      if(!furgon) return;
      if(filtroFurgon && furgon !== filtroFurgon) return;
      if(!porFurgon[furgon]) porFurgon[furgon] = {};
      if(!tarimasFurgon[furgon]) tarimasFurgon[furgon] = {};
      tarimasFurgon[furgon][l.tarima] = true;
      var sku  = l.sku || '?';
      var cant = (l.auditado && l.cantidad_audit != null)
        ? Number(l.cantidad_audit) : Number(l.cantidad);
      var desc = l.descripcion || '';
      if(!porFurgon[furgon][sku]) {
        porFurgon[furgon][sku] = { sku:sku, descripcion:desc, unidades:0, lineasCount:0, pendientes:0 };
      }
      var entry = porFurgon[furgon][sku];
      entry.unidades    += cant;
      entry.lineasCount += 1;
      if(!l.auditado) entry.pendientes += 1;
      if(!entry.descripcion && desc) entry.descripcion = desc;
    });

    var furgonesKeys = Object.keys(porFurgon).sort(function(a,b){
      return String(a).localeCompare(String(b), undefined, { numeric:true });
    });

    var furgonesRes = furgonesKeys.map(function(furgon){
      var skus = Object.values(porFurgon[furgon]);
      var totalUnidades   = skus.reduce(function(s,x){ return s+x.unidades; }, 0);
      var totalPendientes = skus.reduce(function(s,x){ return s+x.pendientes; }, 0);
      var cierre = cierreMap[furgon] || null;
      return {
        furgon:           furgon,
        tarimas:          Object.keys(tarimasFurgon[furgon]).sort(),
        skus:             skus,
        total_unidades:   totalUnidades,
        total_skus:       skus.length,
        tiene_pendientes: totalPendientes > 0,
        finalizado:       !!cierre,
        licencia_hija:    cierre ? cierre.licencia_hija  : null,
        destino_tr999:    cierre ? cierre.destino_tr999  : null,
        ts_cierre:        cierre ? cierre.ts_cierre      : null,
        cerrado_por:      cierre ? cierre.cerrado_por    : null
      };
    });

    res.json({ ok:true, sesion_id:sesId, furgones:furgonesRes, tarimas_sin_asignar:tarimasSinAsignar });
  } catch(e) {
    console.log('BOD manifiesto FAILED:', e.message);
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── POST /api/bod/sesion/:id/furgon/:furgon/finalizar ─────────────────────
// Registra el cierre de un furgón: licencia hija + destino TR999.
// Persiste en bod_furgon_cierres. Idempotente: error 409 si ya cerrado.
app.post('/api/bod/sesion/:id/furgon/:furgon/finalizar', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var furgon  = String(req.params.furgon).trim();
    var { licencia_hija, destino_tr999, observacion, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden finalizar furgones.' });

    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    destino_tr999 = String(destino_tr999 || '').trim().toUpperCase();
    observacion   = String(observacion   || '').trim();
    usuario       = String(usuario       || '').trim();

    if(!licencia_hija) return res.status(400).json({ ok:false, error:'falta licencia_hija' });
    if(!destino_tr999) return res.status(400).json({ ok:false, error:'falta destino_tr999' });
    if(!/^TR999\.\d{3}\.\d{2}$/i.test(destino_tr999))
      return res.status(400).json({ ok:false, error:'El destino debe tener formato TR999.xxx.xx (ej: TR999.001.01). Recibido: '+destino_tr999 });

    // Verificar sesión existe
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'Sesión no encontrada' });
    var sesion = sesRows[0];

    // Verificar que no esté ya finalizado
    try {
      var existing = await supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon)+'&limit=1');
      if(Array.isArray(existing) && existing.length) {
        return res.status(409).json({ ok:false, error:'El furgón '+furgon+' ya fue finalizado con licencia hija '+existing[0].licencia_hija+'.' });
      }
    } catch(e) { /* tabla puede no existir aún — continuar */ }

    // Leer tarimas y líneas
    var asig = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon));
    asig = Array.isArray(asig) ? asig : [];
    if(!asig.length) return res.status(400).json({ ok:false, error:'El furgón '+furgon+' no tiene tarimas asignadas.' });

    var tarimas = asig.map(function(a){ return a.tarima; });
    var tarimasQ = '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&tarima=in.('+tarimas.map(encodeURIComponent).join(',')+')'
      +'&eliminada=eq.false';
    var lineas = await supabase('GET', 'bod_lineas', null, tarimasQ);
    lineas = Array.isArray(lineas) ? lineas : [];
    if(!lineas.length) return res.status(400).json({ ok:false, error:'El furgón '+furgon+' no tiene líneas capturadas.' });

    // Agrupar por SKU
    var skuMap = {};
    lineas.forEach(function(l){
      var cant = (l.auditado && l.cantidad_audit != null) ? Number(l.cantidad_audit) : Number(l.cantidad);
      if(!skuMap[l.sku]) skuMap[l.sku] = { sku:l.sku, descripcion:l.descripcion||'', unidades:0 };
      skuMap[l.sku].unidades += cant;
    });
    var resumenSkus = Object.values(skuMap);
    var totalUnidades = resumenSkus.reduce(function(s,x){ return s+x.unidades; }, 0);

    // Persistir el cierre
    var cierre = {
      id:             'bfc-'+Date.now()+'-'+Math.random().toString(36).slice(2,6),
      sesion_id:      sesId,
      licencia_padre: sesion.licencia_id || '',
      furgon:         furgon,
      licencia_hija:  licencia_hija,
      destino_tr999:  destino_tr999,
      tarimas:        tarimas,
      resumen_skus:   resumenSkus,
      unidades_total: totalUnidades,
      observacion:    observacion,
      cerrado_por:    usuario,
      ts_cierre:      new Date().toISOString()
    };
    await supabase('POST', 'bod_furgon_cierres', cierre, '');

    res.json({
      ok:            true,
      furgon:        furgon,
      licencia_hija: licencia_hija,
      destino_tr999: destino_tr999,
      tarimas:       tarimas,
      skus:          resumenSkus,
      unidades_total: totalUnidades
    });
  } catch(e) {
    console.log('BOD furgon finalizar FAILED:', e.message);
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ══════════════════════════════════════════════════════════════════════════
// BOD MODULE END
// ══════════════════════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════════════════════
// BOD WMS / TRAZABILIDAD — POST /api/bod/wms/upload
// Parsea Excel WMS con columnas: Número, Fecha, Status, SKU, Nombre,
// Unidades, Origen, Destino. Clasifica movimientos y guarda en bod_wms_movimientos.
// FIX (mar 2-jun-2026, server v20): módulo de trazabilidad BOD Fase 2.
// ══════════════════════════════════════════════════════════════════════════
app.post('/api/bod/wms/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    if(!req.file) return res.status(400).json({ ok:false, error:'falta archivo' });
    var usuario          = String((req.body && req.body.usuario)            || '').trim();
    if(!usuario) return res.status(400).json({ ok:false, error:'falta usuario' });
    var licenciaId       = String((req.body && req.body.licencia_id)        || '').trim().toUpperCase();
    var tipoLicencia     = String((req.body && req.body.tipo_licencia)      || 'inicial').trim();
    var licenciaPadre    = String((req.body && req.body.licencia_padre)     || '').trim().toUpperCase();
    var furgonRelacionado= String((req.body && req.body.furgon_relacionado) || '').trim();

    // Validaciones extra para licencias hija
    if(tipoLicencia === 'hija_salida') {
      if(!licenciaId)        return res.status(400).json({ ok:false, error:'falta licencia_id (licencia hija)' });
      if(!licenciaPadre)     return res.status(400).json({ ok:false, error:'falta licencia_padre' });
      if(!furgonRelacionado) return res.status(400).json({ ok:false, error:'falta furgon_relacionado' });
    }

    var wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });
    if(rows.length < 2) return res.status(400).json({ ok:false, error:'Archivo vacío.' });

    var hdr = rows[0].map(function(h){ return String(h).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); });
    var cNum     = hdr.findIndex(function(h){ return h==='numero'||h.includes('numer'); });
    var cFecha   = hdr.findIndex(function(h){ return h==='fecha'; });
    var cStatus  = hdr.findIndex(function(h){ return h==='status'||h==='estado'; });
    var cSku     = hdr.findIndex(function(h){ return h==='sku'||h==='codigo'; });
    var cNombre  = hdr.findIndex(function(h){ return h.includes('nombre')||h.includes('descri'); });
    var cUnis    = hdr.findIndex(function(h){ return h==='unidades'||h==='cantidad'||h==='cant'; });
    var cOrigen  = hdr.findIndex(function(h){ return h==='origen'; });
    var cDestino = hdr.findIndex(function(h){ return h==='destino'; });

    if(cSku < 0 || cUnis < 0 || cOrigen < 0 || cDestino < 0) {
      return res.status(400).json({ ok:false, error:'Faltan columnas: SKU, Unidades, Origen, Destino. Encontrados: '+rows[0].join(', ') });
    }

    var now = new Date().toISOString();
    var nomArchivo = req.file.originalname || 'wms.xlsx';

    // Escanear TODAS las filas y construir set de licencias encontradas
    var licenciasEnArchivo = {};
    if(cNum >= 0) {
      rows.slice(1).forEach(function(r){
        var lic = String(r[cNum]||'').trim().toUpperCase();
        if(lic) licenciasEnArchivo[lic] = true;
      });
    }
    var licsDistintas = Object.keys(licenciasEnArchivo).filter(function(l){
      var norm = l.split(/[\s·\-·]+/)[0];
      var esp  = licenciaId.split(/[\s·\-·]+/)[0];
      return norm !== esp;
    });
    if(licsDistintas.length > 0) {
      return res.status(400).json({
        ok:false,
        error:'El archivo contiene licencias: '+Object.keys(licenciasEnArchivo).join(', ')+
              '. La sesión activa es "'+licenciaId+'". Verificá el archivo o la sesión.'
      });
    }

    // Modo reemplazo selectivo según tipo_licencia
    try {
      if(tipoLicencia === 'hija_salida') {
        // Reemplazar solo movimientos de esta licencia hija + furgón
        await supabase('DELETE', 'bod_wms_movimientos', null,
          '?licencia_id=eq.'+encodeURIComponent(licenciaId)
          +'&furgon_relacionado=eq.'+encodeURIComponent(furgonRelacionado)
          +'&tipo_licencia=eq.hija_salida');
      } else {
        // Reemplazar movimientos iniciales de esta licencia
        await supabase('DELETE', 'bod_wms_movimientos', null,
          '?licencia_id=eq.'+encodeURIComponent(licenciaId)+'&tipo_licencia=eq.inicial');
        // Fallback: si la tabla no tiene la col aún, borrar todos los de esta licencia
      }
    } catch(e) {
      // Fallback: borrar todos los de esa licencia
      try {
        await supabase('DELETE', 'bod_wms_movimientos', null,
          '?licencia_id=eq.'+encodeURIComponent(licenciaId));
      } catch(e2) { console.log('BOD WMS delete previo warn:', e2.message); }
    }

    var movimientos = [];
    var procesadas = 0, omitidas = 0;
    var cntEntradas = 0, cntSalidas = 0, cntOtros = 0;

    rows.slice(1).forEach(function(r){
      var sku     = String(r[cSku]     || '').trim();
      var origen  = String(cOrigen  >= 0 ? (r[cOrigen]  || '') : '').trim().toUpperCase();
      var destino = String(cDestino >= 0 ? (r[cDestino] || '') : '').trim().toUpperCase();
      if(!sku){ omitidas++; return; }
      var unis = Number(r[cUnis]) || 0;
      if(unis <= 0){ omitidas++; return; }

      var tipo;
      if(destino === '952.006.01')                                    tipo = 'entrada_bolson';
      else if(origen === '952.006.01' && /^TR999\./.test(destino))    tipo = 'salida_tr999';
      else                                                             tipo = 'otro';

      if(tipo === 'entrada_bolson') cntEntradas++;
      else if(tipo === 'salida_tr999') cntSalidas++;
      else cntOtros++;

      movimientos.push({
        id:                'bwm-'+Date.now()+'-'+procesadas+'-'+Math.random().toString(36).slice(2,5),
        licencia_id:       licenciaId,
        licencia_padre:    licenciaPadre,
        furgon_relacionado: furgonRelacionado,
        tipo_licencia:     tipoLicencia,
        fecha:             cFecha  >= 0 ? String(r[cFecha]  || '').trim() : '',
        status:            cStatus >= 0 ? String(r[cStatus] || '').trim() : '',
        sku:               sku,
        descripcion:       cNombre >= 0 ? String(r[cNombre] || '').trim() : '',
        unidades:          unis,
        origen:            origen,
        destino:           destino,
        tipo_movimiento:   tipo,
        archivo_origen:    nomArchivo,
        cargado_por:       usuario,
        ts_carga:          now
      });
      procesadas++;
    });

    if(!movimientos.length) return res.status(400).json({ ok:false, error:'Sin filas válidas.' });

    // Insertar en chunks de 200
    var errores = [];
    for(var ci = 0; ci < movimientos.length; ci += 200) {
      var chunk = movimientos.slice(ci, ci + 200);
      try {
        await supabase('POST', 'bod_wms_movimientos', chunk, '?on_conflict=id');
      } catch(e) { errores.push({ desde:ci, error:e.message }); }
    }

    var advertencia;
    if(tipoLicencia === 'hija_salida' && cntSalidas === 0) {
      advertencia = 'La licencia hija fue cargada, pero no contiene salidas desde 952.006.01 hacia TR999.';
    } else if(tipoLicencia !== 'hija_salida' && cntEntradas === 0) {
      advertencia = 'El archivo fue cargado, pero no contiene movimientos con destino 952.006.01. WMS vs App mostrará cero teórico.';
    }

    res.json({
      ok:true, procesadas, omitidas,
      entradas_bolson: cntEntradas,
      salidas_tr999:   cntSalidas,
      otros:           cntOtros,
      tipo_licencia:   tipoLicencia,
      advertencia:     advertencia,
      errores:         errores.length ? errores : undefined
    });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── Helpers para construir filtro de licencias (punto 5) ─────────────────
// Acepta ?licencia_id=X | ?fecha=YYYY-MM-DD | ?licencias=A,B,C
function bodBuildMovQuery(req) {
  var licId   = String(req.query.licencia_id || '').trim().toUpperCase();
  var fecha   = String(req.query.fecha       || '').trim();
  var licsRaw = String(req.query.licencias   || '').trim();
  var lics    = licsRaw ? licsRaw.split(',').map(function(l){ return l.trim().toUpperCase(); }).filter(Boolean) : [];
  if(licId && !lics.includes(licId)) lics.push(licId);

  if(lics.length === 1) return { q:'?licencia_id=eq.'+encodeURIComponent(lics[0])+'&limit=10000', lics:lics };
  if(lics.length > 1)  return { q:'?licencia_id=in.('+lics.map(encodeURIComponent).join(',')+')'+'&limit=10000', lics:lics };
  if(fecha)            return { q:'?fecha=eq.'+encodeURIComponent(fecha)+'&limit=10000', lics:[] };
  return { q:'?limit=5000', lics:[] };
}

// ── GET /api/bod/reportes/wms-vs-app ─────────────────────────────────────
// Acepta: licencia_id | fecha | licencias (CSV)
app.get('/api/bod/reportes/wms-vs-app', bodGuard, async (req, res) => {
  try {
    var f = bodBuildMovQuery(req);

    // app: bod_lineas por licencia_id
    var lineasQ = f.lics.length === 1
      ? '?licencia_id=eq.'+encodeURIComponent(f.lics[0])+'&eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id'
      : f.lics.length > 1
      ? '?licencia_id=in.('+f.lics.map(encodeURIComponent).join(',')+')'+'&eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id&limit=10000'
      : '?eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id&limit=10000';

    var [wmsRows, lineas] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, f.q+'&tipo_movimiento=eq.entrada_bolson'),
      supabase('GET', 'bod_lineas', null, lineasQ)
    ]);
    wmsRows = Array.isArray(wmsRows) ? wmsRows : [];
    lineas  = Array.isArray(lineas)  ? lineas  : [];

    // Agregar por licencia+sku para WMS
    var wmsMap = {}, wmsDesc = {}, wmsLic = {};
    wmsRows.forEach(function(r){
      var k = r.licencia_id+'|'+r.sku;
      wmsMap[k]  = (wmsMap[k]||0) + Number(r.unidades);
      if(!wmsDesc[k] && r.descripcion) wmsDesc[k] = r.descripcion;
      if(!wmsLic[k]) wmsLic[k] = r.licencia_id;
    });
    // Agregar por licencia+sku para APP
    var appMap = {}, appDesc = {}, appLic = {};
    lineas.forEach(function(l){
      var k = l.licencia_id+'|'+l.sku;
      appMap[k]  = (appMap[k]||0) + Number(l.cantidad);
      if(!appDesc[k] && l.descripcion) appDesc[k] = l.descripcion;
      if(!appLic[k]) appLic[k] = l.licencia_id;
    });

    var keys = Object.keys(Object.assign({}, wmsMap, appMap));
    var result = keys.map(function(k){
      var wms = wmsMap[k]||0, app = appMap[k]||0;
      var dif = app - wms;
      var estado;
      if     (wms > 0 && app === wms)       estado = 'Cuadrado';
      else if(wms > 0 && app === 0)         estado = 'Pendiente de entarimar';
      else if(wms > 0 && app > 0 && wms > app) estado = 'Diferencia: faltante APP';
      else if(wms > 0 && app > 0 && app > wms) estado = 'Diferencia: excedente APP';
      else if(wms === 0 && app > 0)         estado = 'Entarimado sin WMS';
      else                                   estado = 'Sin datos';
      var parts = k.split('|');
      var lic = wmsLic[k] || appLic[k] || parts[0];
      var sku = parts.slice(1).join('|');
      return { licencia_id:lic, sku, descripcion:wmsDesc[k]||appDesc[k]||'',
               cantidad_teorica_wms:wms, cantidad_app:app, diferencia:dif, estado };
    });
    res.json({ ok:true, skus:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/reportes/flujo-bolson ───────────────────────────────────
// Acepta: licencia_id=X | licencias=X,Y,Z
// Entradas: tipo_movimiento=entrada_bolson de las licencias indicadas
// Salidas: tipo_movimiento=salida_tr999 donde licencia_padre está en las licencias indicadas
app.get('/api/bod/reportes/flujo-bolson', bodGuard, async (req, res) => {
  try {
    var licId   = String(req.query.licencia_id || '').trim().toUpperCase();
    var licsRaw = String(req.query.licencias   || '').trim();
    var lics    = licsRaw ? licsRaw.split(',').map(function(l){ return l.trim().toUpperCase(); }).filter(Boolean) : [];
    if(licId && !lics.includes(licId)) lics.push(licId);

    // Construir queries para entradas y salidas según cantidad de licencias
    var qEntradas, qSalidas;
    if(lics.length === 1) {
      qEntradas = '?licencia_id=eq.'+encodeURIComponent(lics[0])+'&tipo_movimiento=eq.entrada_bolson&limit=10000';
      qSalidas  = '?licencia_padre=eq.'+encodeURIComponent(lics[0])+'&tipo_movimiento=eq.salida_tr999&limit=10000';
    } else if(lics.length > 1) {
      var inClause = '('+lics.map(encodeURIComponent).join(',')+')';
      qEntradas = '?licencia_id=in.'+inClause+'&tipo_movimiento=eq.entrada_bolson&limit=10000';
      qSalidas  = '?licencia_padre=in.'+inClause+'&tipo_movimiento=eq.salida_tr999&limit=10000';
    } else {
      qEntradas = '?tipo_movimiento=eq.entrada_bolson&limit=5000';
      qSalidas  = '?tipo_movimiento=eq.salida_tr999&limit=5000';
    }

    var [movsEntrada, movsSalida] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, qEntradas),
      supabase('GET', 'bod_wms_movimientos', null, qSalidas)
    ]);
    movsEntrada = Array.isArray(movsEntrada) ? movsEntrada : [];
    movsSalida  = Array.isArray(movsSalida)  ? movsSalida  : [];

    var skuData = {};
    movsEntrada.forEach(function(m){
      if(!skuData[m.sku]) skuData[m.sku] = { sku:m.sku, descripcion:m.descripcion||'',
        entradas_952:0, salidas_tr999:0, licencia_origen:'', licencias_hijas:[], furgones:[], destinos_tr999:[] };
      var d = skuData[m.sku];
      if(!d.descripcion && m.descripcion) d.descripcion = m.descripcion;
      d.entradas_952 += Number(m.unidades);
      if(!d.licencia_origen) d.licencia_origen = m.licencia_id;
    });
    movsSalida.forEach(function(m){
      if(!skuData[m.sku]) skuData[m.sku] = { sku:m.sku, descripcion:m.descripcion||'',
        entradas_952:0, salidas_tr999:0, licencia_origen:'', licencias_hijas:[], furgones:[], destinos_tr999:[] };
      var d = skuData[m.sku];
      if(!d.descripcion && m.descripcion) d.descripcion = m.descripcion;
      d.salidas_tr999 += Number(m.unidades);
      if(m.licencia_id && !d.licencias_hijas.includes(m.licencia_id))      d.licencias_hijas.push(m.licencia_id);
      if(m.furgon_relacionado && !d.furgones.includes(m.furgon_relacionado)) d.furgones.push(m.furgon_relacionado);
      if(m.destino && !d.destinos_tr999.includes(m.destino))                d.destinos_tr999.push(m.destino);
    });

    var result = Object.values(skuData).map(function(d){
      return {
        sku:             d.sku,
        descripcion:     d.descripcion,
        entradas_952:    d.entradas_952,
        salidas_tr999:   d.salidas_tr999,
        remanente:       d.entradas_952 - d.salidas_tr999,
        licencia_origen: d.licencia_origen,
        licencia_hija:   d.licencias_hijas.join(', '),
        furgon:          d.furgones.join(', '),
        destino_tr999:   d.destinos_tr999.join(', ')
      };
    });
    res.json({ ok:true, movimientos:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/reportes/remanentes ─────────────────────────────────────
// Acepta: licencia_id | fecha | licencias (CSV)
// Causa operativa enriquecida cruzando WMS + bod_lineas + bod_tarima_furgon
app.get('/api/bod/reportes/remanentes', bodGuard, async (req, res) => {
  try {
    var f = bodBuildMovQuery(req);
    var lineasQ = f.lics.length === 1
      ? '?licencia_id=eq.'+encodeURIComponent(f.lics[0])+'&eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000'
      : f.lics.length > 1
      ? '?licencia_id=in.('+f.lics.map(encodeURIComponent).join(',')+')'+'&eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000'
      : '?eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000';

    var [movs, lineas, asigs] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, f.q),
      supabase('GET', 'bod_lineas', null, lineasQ),
      supabase('GET', 'bod_tarima_furgon', null, '?limit=5000')
    ]);
    movs   = Array.isArray(movs)   ? movs   : [];
    lineas = Array.isArray(lineas) ? lineas : [];
    asigs  = Array.isArray(asigs)  ? asigs  : [];

    // Tarimas con furgón asignado
    var tarimaConFurgon = {};
    asigs.forEach(function(a){ tarimaConFurgon[a.tarima] = a.furgon; });

    // APP por SKU
    var appMap = {}, skuTarimadas = {}, skuConFurgon = {};
    lineas.forEach(function(l){
      appMap[l.sku] = (appMap[l.sku]||0) + Number(l.cantidad);
      skuTarimadas[l.sku] = true;
      if(tarimaConFurgon[l.tarima]) skuConFurgon[l.sku] = tarimaConFurgon[l.tarima];
    });

    var hoy = new Date();
    var skuData = {};
    movs.forEach(function(m){
      if(!skuData[m.sku]) skuData[m.sku] = { sku:m.sku, descripcion:m.descripcion||'', entradas:0, salidas:0,
        fechas_entrada:[], ultima_fecha:null, licencia_origen:'', licencia_salida:'', destinos_distintos:false };
      var d = skuData[m.sku];
      if(!d.descripcion && m.descripcion) d.descripcion = m.descripcion;
      if(m.tipo_movimiento === 'entrada_bolson'){
        d.entradas += Number(m.unidades);
        if(m.fecha) d.fechas_entrada.push(m.fecha);
        if(!d.licencia_origen) d.licencia_origen = m.licencia_id;
      }
      if(m.tipo_movimiento === 'salida_tr999'){
        d.salidas += Number(m.unidades);
        if(!d.licencia_salida) d.licencia_salida = m.licencia_id;
      }
      if(m.tipo_movimiento === 'otro') d.destinos_distintos = true;
      if(m.fecha && (!d.ultima_fecha || m.fecha > d.ultima_fecha)) d.ultima_fecha = m.fecha;
    });

    // Agregar SKUs que están en APP pero no en WMS
    Object.keys(appMap).forEach(function(sku){
      if(!skuData[sku]) skuData[sku] = { sku:sku, descripcion:'', entradas:0, salidas:0, fechas_entrada:[],
        ultima_fecha:null, licencia_origen:'', licencia_salida:'', destinos_distintos:false };
    });

    var result = Object.values(skuData).filter(function(d){
      var rem = d.entradas - d.salidas;
      return rem > 0 || (appMap[d.sku] || 0) > 0;
    }).map(function(d){
      var rem = d.entradas - d.salidas;
      var app = appMap[d.sku] || 0;
      var fechaMin = d.fechas_entrada.sort()[0] || null;
      var dias = fechaMin ? Math.floor((hoy - new Date(fechaMin)) / 86400000) : null;
      var furgon = skuConFurgon[d.sku] || '';

      var causa;
      if(d.entradas === 0 && d.destinos_distintos)
        causa = 'Pendiente por ubicación distinta a 952';
      else if(d.entradas > 0 && app === 0)
        causa = 'Pendiente de entarimar';
      else if(d.entradas > 0 && app > 0 && !furgon)
        causa = 'Pendiente de carga en furgón';
      else if(d.entradas > 0 && app > 0 && furgon && d.salidas === 0)
        causa = 'Pendiente por ausencia de licencia hija';
      else if(d.entradas > 0 && d.salidas > 0 && rem > 0)
        causa = 'Remanente operativo';
      else if(d.entradas === 0 && app > 0)
        causa = 'Entarimado sin WMS';
      else if(app > 0 && !skuTarimadas[d.sku])
        causa = 'Pendiente de traslado';
      else
        causa = dias !== null && dias <= 2 ? 'Pendiente normal' : 'Atrasado';

      return { sku:d.sku, descripcion:d.descripcion, cantidad_remanente:rem,
               entrada_952:d.entradas, app_entarimada:app, salida_tr999:d.salidas,
               licencia_origen:d.licencia_origen, licencia_salida:d.licencia_salida,
               furgon:furgon, fecha_primer_ingreso:fechaMin, ultimo_movimiento:d.ultima_fecha,
               dias_en_bolson:dias, estado_remanente:causa };
    });
    res.json({ ok:true, remanentes:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/reportes/furgones ───────────────────────────────────────
app.get('/api/bod/reportes/furgones', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.query.sesion_id || '').trim();
    var q = sesId ? '?sesion_id=eq.'+encodeURIComponent(sesId) : '?limit=2000';
    var [asig, lineas] = await Promise.all([
      supabase('GET', 'bod_tarima_furgon', null, q+'&order=furgon.asc,tarima.asc'),
      supabase('GET', 'bod_lineas', null, (sesId?'?sesion_id=eq.'+encodeURIComponent(sesId):'?limit=5000')+'&eliminada=eq.false')
    ]);
    asig   = Array.isArray(asig)   ? asig   : [];
    lineas = Array.isArray(lineas) ? lineas : [];

    var tarimaFurgon = {};
    asig.forEach(function(a){ tarimaFurgon[a.tarima] = a.furgon; });
    var porFurgon = {};
    lineas.forEach(function(l){
      var furg = tarimaFurgon[l.tarima]; if(!furg) return;
      if(!porFurgon[furg]) porFurgon[furg] = {};
      if(!porFurgon[furg][l.sku]) porFurgon[furg][l.sku] = { sku:l.sku, descripcion:l.descripcion||'', cantidad:0, tarimas:new Set() };
      porFurgon[furg][l.sku].cantidad += Number(l.cantidad)||0;
      porFurgon[furg][l.sku].tarimas.add(l.tarima);
    });
    var result = Object.keys(porFurgon).sort().map(function(furg){
      return { furgon:furg, skus: Object.values(porFurgon[furg]).map(function(s){ return Object.assign({}, s, { tarimas_relacionadas: Array.from(s.tarimas) }); }) };
    });
    res.json({ ok:true, furgones:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});
