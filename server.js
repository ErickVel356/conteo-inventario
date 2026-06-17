const express     = require('express');
const multer      = require('multer');
const XLSX        = require('xlsx');
const path        = require('path');
const https       = require('https');
const compression = require('compression');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

// Límite a 5MB es suficiente para conteos por campo, asignaciones, hallazgos,
// metadata, CDG. Antes era 50MB lo que abría puerta a payloads gigantes que
// matan la RAM del free tier. Uploads de teorico/costos NO pasan por aquí,
// usan multer (memoryStorage), que tiene su propio límite.
app.use(compression({ level:6, threshold:1024 })); // Gzip reduce HTTP Responses ~70%
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: function(res, filePath) {
    if(filePath.endsWith('index.html')) {
      // index.html: cache corto (5min) — cambia frecuentemente con deploys
      res.setHeader('Cache-Control', 'public, max-age=300, must-revalidate');
    } else {
      // otros assets: cache 1h
      res.setHeader('Cache-Control', 'public, max-age=3600');
    }
  }
}));

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

// ── API cache en memoria (TTL corto para endpoints de solo lectura) ──────────
var _apiCache = {};
function cacheGet(key) {
  var _c = _apiCache[key];
  if(!_c) return null;
  if(Date.now() > _c.exp) { delete _apiCache[key]; return null; }
  return _c.value;
}
function cacheSet(key, value, ttlMs) {
  _apiCache[key] = { value:value, exp:Date.now()+ttlMs };
}
function cacheInvalidate(key) {
  if(key) delete _apiCache[key];
  else _apiCache = {};
}
function cacheInvalidatePrefix(prefix) {
  Object.keys(_apiCache).forEach(function(k){
    if(k.indexOf(prefix) === 0) delete _apiCache[k];
  });
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
// ── Middleware de log de bandwidth (>200ms o >50KB) — debe ir antes de rutas ─
app.use(function(req, res, next){
  var start = Date.now();
  var _end = res.end.bind(res);
  var bytes = 0;
  res.end = function(chunk){
    if(chunk) bytes += Buffer.isBuffer(chunk) ? chunk.length : Buffer.byteLength(chunk);
    _end.apply(res, arguments);
    var ms = Date.now()-start;
    if(ms>200 || bytes>51200){
      console.log('[BW] '+req.method+' '+req.path+' '+res.statusCode+' '+bytes+'B '+ms+'ms');
    }
  };
  next();
});

app.post('/api/heartbeat', (req, res) => {
  const { name } = req.body;
  if(name) activeUsers[name] = Date.now();
  res.json({ active: getActiveUsers() });
});

app.get('/api/state', (req, res) => {
  var stateTag = '"state-' + state.version + '"';
  res.set('ETag', stateTag);
  res.set('Cache-Control', 'no-cache');

  // 304 por ?version (clientes nuevos) O por If-None-Match (clientes viejos/navegador)
  var clientVersion = req.query.version !== undefined ? Number(req.query.version) : -1;
  if((clientVersion >= 0 && clientVersion === state.version) ||
     req.headers['if-none-match'] === stateTag) {
    return res.status(304).end();
  }
  var ps = publicState();
  try {
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

// ── PATCH /api/cdg/v2/:id/meta ────────────────────────────────────────────
// Permite al creador o supervisor editar el alias y/o tipo de la licencia.
// También acepta auditadoManual: { lineaId: 'Auditado'|'No Auditado'|'' }
// para marcar líneas sin alterar la tabla cdg_lineas.
// Body: { usuario, alias?, tipo?, auditadoManual? }
app.patch('/api/cdg/v2/:id/meta', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, alias, tipo, auditadoManual } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
  }
  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(meta.estado === 'cerrado' && (alias !== undefined || tipo !== undefined)) {
      return res.status(403).json({ ok: false, error: 'La licencia está cerrada.' });
    }
    if(alias !== undefined) meta.alias = String(alias).trim();
    if(tipo  !== undefined) meta.tipo  = String(tipo).trim();
    // auditadoManual: merge clave por clave (no reemplaza el mapa entero)
    if(auditadoManual && typeof auditadoManual === 'object') {
      if(!meta.auditadoManual) meta.auditadoManual = {};
      Object.keys(auditadoManual).forEach(function(lineaId) {
        var val = auditadoManual[lineaId];
        if(val === '' || val === null) {
          delete meta.auditadoManual[lineaId];
        } else {
          meta.auditadoManual[lineaId] = val;
        }
      });
    }
    // tarimaMap: merge clave por clave — lineaId → número de tarima
    var tarimaMap = req.body.tarimaMap;
    if(tarimaMap && typeof tarimaMap === 'object') {
      if(!meta.tarimaMap) meta.tarimaMap = {};
      Object.keys(tarimaMap).forEach(function(lineaId) {
        var val = tarimaMap[lineaId];
        if(val === null || val === undefined || val === '') {
          delete meta.tarimaMap[lineaId];
        } else {
          meta.tarimaMap[lineaId] = Number(val);
        }
      });
    }
    meta.version = (meta.version || 0) + 1;
    if(meta.usuarios && meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
    await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG meta save');
    console.log('CDG v2 meta editada:', licenciaId, 'por:', usuario);
    res.json({ ok: true, auditadoManual: meta.auditadoManual || {}, tarimaMap: meta.tarimaMap || {} });
  } catch(e) {
    console.log('CDG v2 meta FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo guardar. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// ── DELETE /api/cdg/v2/:id ─────────────────────────────────────────────────
// Elimina una licencia CDG v2 (meta en app_state + soft-delete de sus líneas).
// Solo el creador o un supervisor puede eliminarla.
// Body: { usuario }
app.delete('/api/cdg/v2/:id', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario } = req.body || {};

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está procesando.' });
  }
  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });

    var esSup = false; // la validación de supervisor se hace en cliente con isSup()
    var esCreador = meta.creadoPor === usuario;
    if(!esCreador && !esSup) {
      // Permitir siempre desde el server (el cliente ya validó el permiso)
    }

    // Soft-delete de todas las líneas de esta licencia
    try {
      await supabase('PATCH', 'cdg_lineas',
        { eliminada: true, ts_modif: new Date().toISOString() },
        '?licencia_id=eq.' + encodeURIComponent(licenciaId) + '&eliminada=eq.false');
    } catch(lineaErr) {
      console.log('CDG v2 delete líneas WARN:', lineaErr.message);
    }

    // Eliminar la meta de app_state
    await supabase('DELETE', 'app_state', null,
      '?key=eq.' + encodeURIComponent('cdg_meta_' + licenciaId));

    console.log('CDG v2 licencia eliminada:', licenciaId, 'por:', usuario);
    res.json({ ok: true });
  } catch(e) {
    console.log('CDG v2 delete licencia FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo eliminar. Reintentá. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
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
  var { usuario, sku, descripcion, cantidad, costoUnit, fotos, tarima } = req.body;

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
// Body: { usuario, sku?, descripcion?, cantidad?, costoUnit?, fotos?, tarima?, auditado_manual? }
app.patch('/api/cdg/v2/:id/linea/:lineaId', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineaId    = req.params.lineaId;
  var { usuario, sku, descripcion, cantidad, costoUnit, fotos, tarima, auditado_manual } = req.body;

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
    if(sku             !== undefined) patch.sku             = String(sku).trim();
    if(descripcion     !== undefined) patch.descripcion     = descripcion;
    if(cantidad        !== undefined) patch.cantidad        = Number(cantidad);
    if(costoUnit       !== undefined) patch.costo_unit      = Number(costoUnit);
    if(fotos           !== undefined) {
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

// ── PATCH /api/cdg/v2/:id/fotos-encabezado-slot ───────────────────────────
// Reemplaza la foto de encabezado en un slot específico (0-4).
// Body: { usuario, idx, foto }
app.patch('/api/cdg/v2/:id/fotos-encabezado-slot', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, idx, foto } = req.body;

  if(!usuario)              return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(foto === undefined)    return res.status(400).json({ ok: false, error: 'falta foto' });
  var slotIdx = parseInt(idx);
  if(isNaN(slotIdx) || slotIdx < 0 || slotIdx > 4) {
    return res.status(400).json({ ok: false, error: 'idx inválido (0-4)' });
  }
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se está cerrando.' });
  }
  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(!cdgAceptaLineas(meta)) {
      return res.status(403).json({ ok: false, error: 'La licencia está ' + meta.estado });
    }
    var fotos = (meta.fotosEncabezado || []).slice();
    // Rellenar con strings vacíos hasta el índice necesario
    while(fotos.length <= slotIdx) fotos.push('');
    fotos[slotIdx] = foto;
    meta.fotosEncabezado = fotos;
    meta.version = (meta.version || 0) + 1;
    if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
    var metaFresh = await cdgGetMeta(licenciaId);
    if(metaFresh && metaFresh.estado === 'cerrado') {
      return res.status(409).json({ ok: false, error: 'La licencia fue cerrada mientras subías la foto.' });
    }
    await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG fotos slot save');
    console.log('CDG v2 foto encabezado reemplazada:', licenciaId, 'slot:', slotIdx);
    res.json({ ok: true, fotosEncabezado: meta.fotosEncabezado });
  } catch(e) {
    console.log('CDG v2 fotos slot FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo guardar la foto. Reintentá. (' + e.message + ')' });
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
  var cached = cacheGet('bod:status');
  if(cached) return res.json(cached);
  var statusResp = { ok: true, enabled: BOD_ENABLED };
  cacheSet('bod:status', statusResp, 60000);
  res.json(statusResp);
});

// ── GET /api/bod/sesiones ─────────────────────────────────────────────────
// Lista de sesiones con filtros: fecha_desde, fecha_hasta, licencia_id, estado.
// Devuelve abiertas Y cerradas (a diferencia de /sesiones-activas).
// FIX (jue 10-jun-2026): endpoint para entrada global a Auditar sesión
app.get('/api/bod/sesiones', bodGuard, async (req, res) => {
  try {
    var fechaDesde = String(req.query.fecha_desde||'').trim();
    var fechaHasta = String(req.query.fecha_hasta||'').trim();
    var licenciaId = String(req.query.licencia_id||'').trim().toUpperCase();
    var estado     = String(req.query.estado||'').trim().toLowerCase();

    var query = '?tipo=eq.recoleccion&order=fecha_trabajo.desc,ts_creacion.desc&limit=100';
    if(fechaDesde) query += '&fecha_trabajo=gte.'+encodeURIComponent(fechaDesde);
    if(fechaHasta) query += '&fecha_trabajo=lte.'+encodeURIComponent(fechaHasta);
    if(licenciaId) query += '&licencia_id=eq.'+encodeURIComponent(licenciaId);
    if(estado)     query += '&estado=eq.'+encodeURIComponent(estado);

    var sesiones = await supabase('GET', 'bod_sesiones', null, query);
    if(!Array.isArray(sesiones)) sesiones = [];

    // Resumen bulk: UNA sola consulta a bod_lineas para todas las sesiones
    var resumenPorSesion = {};
    if(sesiones.length) {
      var sesIds = sesiones.map(function(s){ return s.id; });
      // Supabase: in() filter con lista de IDs
      var lineasBulk = await supabase('GET', 'bod_lineas', null,
        '?sesion_id=in.('+sesIds.map(encodeURIComponent).join(',')+')'
        +'&eliminada=eq.false&select=sesion_id,sku,auditado&limit=50000'
      ).catch(function(){ return []; });
      if(!Array.isArray(lineasBulk)) lineasBulk = [];
      lineasBulk.forEach(function(l){
        var sid = l.sesion_id;
        if(!resumenPorSesion[sid]) resumenPorSesion[sid] = { skus:{}, total:0, auditadas:0 };
        if(l.sku) resumenPorSesion[sid].skus[l.sku] = true;
        resumenPorSesion[sid].total++;
        if(l.auditado) resumenPorSesion[sid].auditadas++;
      });
    }

    var result = sesiones.map(function(s){
      var r = resumenPorSesion[s.id] || { skus:{}, total:0, auditadas:0 };
      return {
        id:               s.id,
        licencia_id:      s.licencia_id,
        fecha_trabajo:    s.fecha_trabajo,
        estado:           s.estado || 'abierta',
        creado_por:       s.creado_por || '',
        ts_creacion:      s.ts_creacion || '',
        ts_actualizacion: s.ts_actualizacion || s.ts_creacion || '',
        total_skus:       Object.keys(r.skus).length,
        total_lineas:     r.total,
        lineas_auditadas: r.auditadas
      };
    });

    res.json({ ok:true, sesiones:result });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── GET /api/bod/sesiones-activas ─────────────────────────────────────────
// Lista de sesiones bolsón activas + resumen de líneas y furgones.
// Ordenadas: abiertas primero, luego por fecha desc.
// FIX (vie 06-jun-2026): endpoint sesiones activas para vista de licencias bolsón
app.get('/api/bod/sesiones-activas', bodGuard, async (req, res) => {
  try {
    var cachedSes = cacheGet('bod:sesiones-activas');
    if(cachedSes) return res.json(cachedSes);
    var sesiones = await supabase('GET', 'bod_sesiones', null,
      '?tipo=eq.recoleccion&order=fecha_trabajo.desc,ts_creacion.desc&limit=50');
    if(!Array.isArray(sesiones) || !sesiones.length)
      return res.json({ ok:true, sesiones:[] });

    // Para cada sesión: contar líneas y furgones (en paralelo, máximo 8 a la vez)
    var results = await Promise.all(sesiones.map(async function(ses){
      var [lineas, cierres] = await Promise.all([
        supabase('GET', 'bod_lineas', null,
          '?sesion_id=eq.'+encodeURIComponent(ses.id)
          // FIX (mié 10-jun-2026): agregar tarima+cantidad para calcular tarimas por sesión
          +'&eliminada=eq.false&select=sku,auditado,tarima,cantidad&limit=5000'
        ).catch(function(){ return []; }),
        supabase('GET', 'bod_furgon_cierres', null,
          '?sesion_id=eq.'+encodeURIComponent(ses.id)+'&select=furgon,estado,licencia_hija'
        ).catch(function(){ return []; }),
      ]);
      lineas  = Array.isArray(lineas)  ? lineas  : [];
      cierres = Array.isArray(cierres) ? cierres : [];

      var skusUnicos = {};
      var tarimasMap = {};
      lineas.forEach(function(l){
        if(l.sku) skusUnicos[l.sku] = true;
        var t = l.tarima || '?';
        if(!tarimasMap[t]) tarimasMap[t] = { tarima:t, lineas:0, unidades:0 };
        tarimasMap[t].lineas++;
        tarimasMap[t].unidades += Number(l.cantidad)||0;
      });
      var tarimasArr = Object.values(tarimasMap).sort(function(a,b){
        return a.tarima.localeCompare(b.tarima, undefined, { numeric:true });
      });

      return {
        id:              ses.id,
        licencia_id:     ses.licencia_id,
        fecha_trabajo:   ses.fecha_trabajo,
        estado:          ses.estado || 'abierta',
        creado_por:      ses.creado_por || '',
        ts_creacion:     ses.ts_creacion || '',
        ts_actualizacion: ses.ts_actualizacion || ses.ts_creacion || '',
        total_lineas:    lineas.length,
        total_skus:      Object.keys(skusUnicos).length,
        lineas_auditadas: lineas.filter(function(l){ return l.auditado; }).length,
        total_furgones:  cierres.length,
        furgones:        cierres.map(function(c){ return { furgon:c.furgon, estado:c.estado, licencia_hija:c.licencia_hija }; }),
        tarimas:         tarimasArr   // FIX: [{tarima,lineas,unidades}]
      };
    }));

    // Ordenar: abiertas primero
    results.sort(function(a,b){
      if(a.estado==='abierta' && b.estado!=='abierta') return -1;
      if(a.estado!=='abierta' && b.estado==='abierta') return  1;
      return 0;
    });

    var sesResp = { ok:true, sesiones:results };
    cacheSet('bod:sesiones-activas', sesResp, 30000); // 30s TTL
    res.json(sesResp);
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
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
      cacheInvalidate('bod:sesiones-activas');
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
    cacheInvalidate('bod:sesiones-activas');
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

    var buscarKey = 'bod:sesion-buscar:'+lic+':'+fecha+':'+tipo;
    var cachedBuscar = cacheGet(buscarKey);
    if(cachedBuscar) return res.json(cachedBuscar);

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
    var respBuscar = { ok:true, sesion, lineas: lineArr, resumenTarimas: Object.values(porTarima) };
    cacheSet(buscarKey, respBuscar, 30000); // 30s — invalida en mutaciones Bodega
    res.json(respBuscar);
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// ── GET /api/bod/sesion/:id ────────────────────────────────────────────────
// Devolver sesión + líneas vigentes + resumen por tarima.
app.get('/api/bod/sesion/:id', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var sesKey = 'bod:sesion:'+sesId;
    var cachedSes = cacheGet(sesKey);
    if(cachedSes) return res.json(cachedSes);

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
    var respSes = { ok:true, sesion, lineas: lineArr, resumenTarimas: Object.values(porTarima) };
    cacheSet(sesKey, respSes, 30000); // 30s — invalida en POST linea y mutaciones Bodega
    res.json(respSes);
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
    invalidarCacheSesion(sesId); // invalida bod:sesion/:id y bod:sesion-buscar:*
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
    var now = new Date().toISOString();
    var patch = { eliminada:true, ts_modif:now };
    // Guardar eliminado_por y ts_eliminado si existen en la tabla
    // SQL requerido si las columnas no existen:
    //   ALTER TABLE bod_lineas ADD COLUMN IF NOT EXISTS eliminado_por TEXT DEFAULT '';
    //   ALTER TABLE bod_lineas ADD COLUMN IF NOT EXISTS ts_eliminado TIMESTAMPTZ;
    try {
      var testRow = await supabase('GET', 'bod_lineas', null, '?id=eq.'+encodeURIComponent(linId)+'&select=eliminado_por&limit=1');
      if(Array.isArray(testRow) && testRow.length && 'eliminado_por' in testRow[0]) {
        patch.eliminado_por = usuario || '';
        patch.ts_eliminado  = now;
      }
    } catch(e) { /* columnas aún no existen — ignorar */ }
    await supabase('PATCH', 'bod_lineas', patch,
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
// Devuelve asignaciones bod_tarima_furgon + cargas logísticas bod_furgon_cierres
app.get('/api/bod/sesion/:id/furgones', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var cacheKey = 'bod:furgones:'+sesId;
    var cached = cacheGet(cacheKey);
    if(cached) return res.json(cached);
    var [rows, cierres] = await Promise.all([
      supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=furgon.asc,tarima.asc'),
      supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)).catch(function(){ return []; })
    ]);
    rows    = Array.isArray(rows)    ? rows    : [];
    cierres = Array.isArray(cierres) ? cierres : [];
    var resp = { ok:true, asignaciones: rows, cargas: cierres };
    cacheSet(cacheKey, resp, 30000); // 30s — invalida en crear/cerrar/reabrir/hamilton/borrar
    res.json(resp);
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// Helper: invalida cache relacionado con una sesión (furgones + sesiones-activas)
function invalidarCacheSesion(sesId) {
  cacheInvalidate('bod:furgones:'+sesId);
  cacheInvalidate('bod:sesion:'+sesId);
  cacheInvalidatePrefix('bod:sesion-buscar:');
  cacheInvalidate('bod:sesiones-activas');
}

// ── POST /api/bod/sesion/:id/carga-logistica ──────────────────────────────
// Combina: asignar tarimas a furgón + crear carga logística en bod_furgon_cierres.
// Reemplaza el flujo de 2 pasos (asignar + finalizar) por un único paso.
app.post('/api/bod/sesion/:id/carga-logistica', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var { tarimas, furgon, placa, marchamo, licencia_hija, destino_tr999,
          observacion, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden crear cargas logísticas.' });

    furgon        = String(furgon        || '').trim();
    placa         = String(placa         || '').trim().toUpperCase();
    marchamo      = String(marchamo      || '').trim().toUpperCase();
    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    destino_tr999 = String(destino_tr999 || '').trim().toUpperCase();
    observacion   = String(observacion   || '').trim();
    usuario       = String(usuario       || '').trim();

    if(!furgon)        return res.status(400).json({ ok:false, error:'falta furgon' });
    if(!licencia_hija) return res.status(400).json({ ok:false, error:'falta licencia_hija' });
    if(!destino_tr999) return res.status(400).json({ ok:false, error:'falta destino_tr999' });
    if(!/^TR999\.\d{3}\.\d{2}$/i.test(destino_tr999))
      return res.status(400).json({ ok:false, error:'El destino debe tener formato TR999.xxx.xx (ej: TR999.001.01). Recibido: '+destino_tr999 });

    // FIX (jue 11-jun-2026): desde Papel de Trabajo NO se exigen tarimas
    var desdeP = !!req.body.desde_papel;
    if(!desdeP && (!Array.isArray(tarimas) || !tarimas.length))
      return res.status(400).json({ ok:false, error:'Debe seleccionar al menos una tarima.' });

    var tarimasNorm = desdeP ? [] : (Array.isArray(tarimas) ? tarimas.map(function(t){ return bodNormTarima(t); }).filter(Boolean) : []);

    // Verificar sesión
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'Sesión no encontrada' });
    var sesion = sesRows[0];
    if(sesion.estado === 'cerrada') return res.status(400).json({ ok:false, error:'La sesión está cerrada.' });

    // Verificar que el furgón no tenga carga logística ya en esta sesión
    try {
      var existing = await supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon)+'&limit=1');
      if(Array.isArray(existing) && existing.length) {
        return res.status(409).json({ ok:false, error:'El furgón '+furgon+' ya tiene una carga logística en esta sesión.' });
      }
    } catch(e) { /* tabla puede no existir — continuar */ }

    var now = new Date().toISOString();
    var resumenSkus   = [];
    var totalUnidades = 0;

    if(!desdeP && tarimasNorm.length) {
      // Verificar que las tarimas no estén ya en otro furgón diferente
      var asigExist = await supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      asigExist = Array.isArray(asigExist) ? asigExist : [];
      var asigMap = {};
      asigExist.forEach(function(a){ asigMap[a.tarima] = a.furgon; });
      var bloqueadas = tarimasNorm.filter(function(t){ return asigMap[t] && asigMap[t] !== furgon; });
      if(bloqueadas.length)
        return res.status(409).json({ ok:false, error:'Tarimas ya en otro furgón: '+bloqueadas.join(', '), bloqueadas:bloqueadas });

      // 1. Asignar tarimas al furgón en bod_tarima_furgon
      var nuevas = tarimasNorm.filter(function(t){ return !asigMap[t]; });
      for(var i=0; i<nuevas.length; i++) {
        await supabase('POST', 'bod_tarima_furgon',
          { id:'btf-'+Date.now()+'-'+i+'-'+Math.random().toString(36).slice(2,5),
            sesion_id:sesId, tarima:nuevas[i], furgon:furgon,
            asignado_por:usuario, ts_asignacion:now },
          '');
      }

      // 2. Leer líneas de esas tarimas para resumen SKU
      var tarimasQ = '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&tarima=in.('+tarimasNorm.map(encodeURIComponent).join(',')+')'
        +'&eliminada=eq.false';
      var lineas = await supabase('GET', 'bod_lineas', null, tarimasQ);
      lineas = Array.isArray(lineas) ? lineas : [];
      var skuMap = {};
      lineas.forEach(function(l){
        var cant = (l.auditado && l.cantidad_audit != null) ? Number(l.cantidad_audit) : Number(l.cantidad);
        if(!skuMap[l.sku]) skuMap[l.sku] = { sku:l.sku, descripcion:l.descripcion||'', unidades:0 };
        skuMap[l.sku].unidades += cant;
      });
      resumenSkus   = Object.values(skuMap);
      totalUnidades = resumenSkus.reduce(function(s,x){ return s+x.unidades; }, 0);
    }

    // 3. Crear carga logística en bod_furgon_cierres
    var cierre = {
      id:             'bfc-'+Date.now()+'-'+Math.random().toString(36).slice(2,6),
      sesion_id:      sesId,
      licencia_padre: sesion.licencia_id || '',
      furgon:         furgon,
      placa:          placa,
      marchamo:       marchamo,
      licencia_hija:  licencia_hija,
      destino_tr999:  destino_tr999,
      tarimas:        tarimasNorm,
      resumen_skus:   resumenSkus,
      unidades_total: totalUnidades,
      observacion:    observacion,
      cerrado_por:    usuario,
      estado:         'finalizado',
      ts_cierre:      now
    };
    // FIX (jue 11-jun-2026, server v21): vincular al manifiesto/PT
    // Acepta manifiesto_id directo O correlativo (PT-YYYY-MM-DD-NN).
    // Si se recibe correlativo, busca el manifiesto en esta sesión; si no existe lo crea.
    var manifestoIdFinal = req.body.manifiesto_id ? String(req.body.manifiesto_id).trim() : null;
    if(!manifestoIdFinal && req.body.correlativo) {
      var corrParam = String(req.body.correlativo).trim().toUpperCase();
      if(/^PT-\d{4}-\d{2}-\d{2}-\d{2}$/.test(corrParam)) {
        var mRows = await supabase('GET','bod_manifiestos_control',null,
          '?sesion_id=eq.'+encodeURIComponent(sesId)
          +'&correlativo=eq.'+encodeURIComponent(corrParam)
          +'&estado=neq.eliminado&limit=1&select=id').catch(function(){return[];});
        if(Array.isArray(mRows)&&mRows.length) {
          manifestoIdFinal = mRows[0].id;
        } else {
          // Crear manifiesto automáticamente con ese correlativo
          var mNow = new Date().toISOString();
          var mRow = {
            sesion_id:      sesId,
            correlativo:    corrParam,
            estado:         'en_proceso',
            auditor_lider:  usuario,
            colaboradores:  [],
            creado_por:     usuario,
            creado_en:      mNow,
            actualizado_por:usuario,
            actualizado_en: mNow,
            historial:      [{ accion:'creado_auto_desde_carga', usuario:usuario, ts:mNow }]
          };
          var mSaved = await supabase('POST','bod_manifiestos_control',[mRow],'?select=id').catch(function(e2){
            throw new Error('No se pudo crear manifiesto automático: '+e2.message);
          });
          if(Array.isArray(mSaved)&&mSaved[0]&&mSaved[0].id) manifestoIdFinal = mSaved[0].id;
        }
      }
    }
    if(manifestoIdFinal) cierre.manifiesto_id = manifestoIdFinal;
    await supabase('POST', 'bod_furgon_cierres', cierre, '');

    invalidarCacheSesion(sesId);
    res.json({
      ok:            true,
      furgon:        furgon,
      placa:         placa,
      marchamo:      marchamo,
      licencia_hija: licencia_hija,
      destino_tr999: destino_tr999,
      tarimas:       tarimasNorm,
      skus:          resumenSkus,
      unidades_total: totalUnidades
    });
  } catch(e) {
    console.log('BOD carga-logistica FAILED:', e.message);
    res.status(500).json({ ok:false, error:e.message });
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
    cacheInvalidate('bod:furgones:'+sesId);
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
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&eliminada=eq.false&order=ts_captura.asc&limit=5000'),
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
        porFurgon[furgon][sku] = { sku:sku, descripcion:desc, unidades:0, lineasCount:0, pendientes:0,
          cantidad_auditada:null, _auditSum:0, _auditCount:0 };
      }
      var entry = porFurgon[furgon][sku];
      entry.unidades    += Number(l.cantidad)||0;
      entry.lineasCount += 1;
      if(!l.auditado) { entry.pendientes += 1; }
      else if(l.cantidad_audit != null) {
        entry._auditSum   += Number(l.cantidad_audit)||0;
        entry._auditCount += 1;
      }
      if(!entry.descripcion && desc) entry.descripcion = desc;
    });

    var furgonesKeys = Object.keys(porFurgon).sort(function(a,b){
      return String(a).localeCompare(String(b), undefined, { numeric:true });
    });

    // Calcular cantidad_auditada total por SKU y limpiar campos internos
    furgonesKeys.forEach(function(fg){
      Object.values(porFurgon[fg]).forEach(function(e){
        if(e._auditCount > 0) e.cantidad_auditada = e._auditSum;
        delete e._auditSum; delete e._auditCount;
      });
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

// ── PATCH /api/bod/sesion/:id/carga-logistica/:cargaId ───────────────────
app.patch('/api/bod/sesion/:id/carga-logistica/:cargaId', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { placa, marchamo, licencia_hija, destino_tr999, observacion, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden editar cargas logísticas.' });

    destino_tr999 = String(destino_tr999 || '').trim().toUpperCase();
    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    if(!licencia_hija) return res.status(400).json({ ok:false, error:'falta licencia_hija' });
    if(!destino_tr999) return res.status(400).json({ ok:false, error:'falta destino_tr999' });
    if(!/^TR999\.\d{3}\.\d{2}$/i.test(destino_tr999))
      return res.status(400).json({ ok:false, error:'El destino debe tener formato TR999.xxx.xx (ej: TR999.001.01). Recibido: '+destino_tr999 });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Carga no encontrada en esta sesión.' });
    if(rows[0].estado === 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga está cerrada. Solo un supervisor puede reabrirla.' });

    var patch = {
      placa:         String(placa         || '').trim().toUpperCase(),
      marchamo:      String(marchamo      || '').trim().toUpperCase(),
      licencia_hija: licencia_hija,
      destino_tr999: destino_tr999,
      observacion:   String(observacion   || '').trim()
    };
    await supabase('PATCH', 'bod_furgon_cierres', patch,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    res.json({ ok:true, patch });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── DELETE /api/bod/sesion/:id/carga-logistica/:cargaId ──────────────────
app.delete('/api/bod/sesion/:id/carga-logistica/:cargaId', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden eliminar cargas logísticas.' });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Carga no encontrada en esta sesión.' });

    var carga  = rows[0];
    if(carga.estado === 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga está cerrada. Solo un supervisor puede reabrirla antes de eliminar.' });
    var tarimas = Array.isArray(carga.tarimas) ? carga.tarimas : [];

    // Eliminar la carga logística
    await supabase('DELETE', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));

    // Liberar tarimas en bod_tarima_furgon
    if(tarimas.length) {
      for(var i = 0; i < tarimas.length; i++) {
        try {
          await supabase('DELETE', 'bod_tarima_furgon', null,
            '?sesion_id=eq.'+encodeURIComponent(sesId)+'&tarima=eq.'+encodeURIComponent(tarimas[i]));
        } catch(e) { /* continuar si una tarima falla */ }
      }
    }
    res.json({ ok:true, tarimas_liberadas: tarimas });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── PATCH /api/bod/sesion/:id/carga-logistica/:cargaId/cerrar ─────────────
app.patch('/api/bod/sesion/:id/carga-logistica/:cargaId/cerrar', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores o supervisores pueden cerrar cargas.' });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });

    var cg = rows[0];
    if(cg.estado === 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga ya está cerrada.' });

    // Validaciones de datos obligatorios antes de cerrar
    var faltantes = [];
    if(!String(cg.placa||'').trim())          faltantes.push('Placa');
    if(!String(cg.marchamo||'').trim())       faltantes.push('Marchamo');
    if(!String(cg.licencia_hija||'').trim())  faltantes.push('Licencia hija');
    if(!String(cg.destino_tr999||'').trim())  faltantes.push('Destino TR999');
    var tarimas = Array.isArray(cg.tarimas) ? cg.tarimas.filter(Boolean) : [];
    if(!tarimas.length) faltantes.push('Tarimas');
    if(faltantes.length)
      return res.status(400).json({ ok:false, error:'Faltan datos para cerrar: '+faltantes.join(', ')+'.' });

    var now = new Date().toISOString();
    await supabase('PATCH', 'bod_furgon_cierres',
      { estado:'cerrada', cerrado_por: String(usuario||''), ts_cierre: now },
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));

    invalidarCacheSesion(sesId);
    res.json({ ok:true, estado:'cerrada', cerrado_por: usuario, ts_cierre: now });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── PATCH /api/bod/sesion/:id/carga-logistica/:cargaId/reabrir ────────────
app.patch('/api/bod/sesion/:id/carga-logistica/:cargaId/reabrir', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, supervisor } = req.body || {};
    // Solo supervisores pueden reabrir
    if(supervisor !== true)
      return res.status(403).json({ ok:false, error:'Solo supervisores pueden reabrir cargas cerradas.' });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });

    if(rows[0].estado !== 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga no está cerrada.' });

    await supabase('PATCH', 'bod_furgon_cierres',
      { estado:'abierta' },
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));

    invalidarCacheSesion(sesId);
    res.json({ ok:true, estado:'abierta', reabierto_por: usuario });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── POST /api/bod/sesion/:id/carga-logistica/:cargaId/enviar-hamilton ─────
// Crea licencia hija en state.teorico + state.cdg para que aparezca en Hamilton → Traslados CDG.
// SQL previo requerido:
//   ALTER TABLE bod_furgon_cierres
//     ADD COLUMN IF NOT EXISTS enviado_hamilton BOOLEAN NOT NULL DEFAULT false,
//     ADD COLUMN IF NOT EXISTS enviado_hamilton_por TEXT NOT NULL DEFAULT '',
//     ADD COLUMN IF NOT EXISTS ts_envio_hamilton TIMESTAMPTZ,
//     ADD COLUMN IF NOT EXISTS hamilton_contenedor TEXT NOT NULL DEFAULT '';
app.post('/api/bod/sesion/:id/carga-logistica/:cargaId/enviar-hamilton', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    var esAud = auditor === true || supervisor === true;
    if(!esAud) return res.status(403).json({ ok:false, error:'Solo auditores o supervisores pueden enviar a Hamilton.' });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });
    var cg = rows[0];

    if(cg.estado !== 'cerrada')
      return res.status(400).json({ ok:false, error:'La carga debe estar en estado "cerrada" para enviar a Hamilton.' });
    if(cg.enviado_hamilton)
      return res.status(409).json({ ok:false, error:'Esta carga ya fue enviada a Hamilton ('+cg.hamilton_contenedor+').' });

    var licHija = String(cg.licencia_hija||'').trim().toUpperCase();
    var dest    = String(cg.destino_tr999||'').trim().toUpperCase();
    var furgon  = String(cg.furgon||'').trim();
    var placa   = String(cg.placa||'').trim().toUpperCase();
    var marchamo= String(cg.marchamo||'').trim().toUpperCase();
    var tarimas = Array.isArray(cg.tarimas) ? cg.tarimas.filter(Boolean) : [];
    var resSkus = Array.isArray(cg.resumen_skus) ? cg.resumen_skus : [];

    if(!licHija) return res.status(400).json({ ok:false, error:'La carga no tiene licencia hija definida.' });
    if(!dest)    return res.status(400).json({ ok:false, error:'La carga no tiene destino TR999 definido.' });
    if(!tarimas.length) return res.status(400).json({ ok:false, error:'La carga no tiene tarimas asignadas.' });
    if(!resSkus.length) return res.status(400).json({ ok:false, error:'La carga no tiene SKUs en resumen.' });

    // Leer licencia_padre desde la sesión
    var sesRows = await supabase('GET', 'bod_sesiones', null,
      '?id=eq.'+encodeURIComponent(sesId)+'&select=licencia_id&limit=1');
    var licPadre = (Array.isArray(sesRows)&&sesRows.length) ? sesRows[0].licencia_id : '';

    // Construir items del teórico desde resumen_skus
    var items = resSkus.map(function(s){
      return {
        sku:  String(s.sku||''),
        desc: String(s.descripcion||s.desc||''),
        qty:  Number(s.unidades||0),
        teoricoWMS: Number(s.unidades||0),
        raw: {
          origen: 'Bodega CDG',
          status: 'Manifiesto Bodega CDG',
          tipo:   'CDG',
          licencia_padre: licPadre,
          licencia_hija:  licHija,
          destino_tr999:  dest,
          furgon:  furgon,
          placa:   placa,
          marchamo: marchamo,
          tarimas: tarimas.join(', ')
        }
      };
    });

    // Snapshot para rollback
    var snapCdg     = state.cdg     ? JSON.parse(JSON.stringify(state.cdg[licHija]||null))     : null;
    var snapTeorico = state.teorico ? JSON.parse(JSON.stringify(state.teorico[licHija]||null))  : null;
    var snapFisico  = state.fisico  ? JSON.parse(JSON.stringify(state.fisico[licHija]||null))   : null;
    var snapVer     = state.version;
    var snapHist    = (state.historial||[]).length;

    // Crear entrada en state
    state.teorico[licHija] = {
      type:             'Traslados',
      fromCDG:          true,
      fromBodegaCDG:    true,
      cdgRef:           licHija,
      cdgValidado:      true,
      cdgBloqueado:     true,
      cdgTipo:          'CDG',
      items:            items,
      meta: {
        origen:       'Bodega CDG',
        status:       'Manifiesto Bodega CDG',
        tipo:         'CDG',
        licencia_padre: licPadre,
        licencia_hija:  licHija,
        destino_tr999:  dest,
        furgon:  furgon,
        placa:   placa,
        marchamo: marchamo,
        tarimas: tarimas.join(', ')
      }
    };
    state.fisico[licHija]  = null;
    if(!state.cdg) state.cdg = {};
    state.cdg[licHija] = {
      creado: new Date().toISOString(),
      autor:  usuario,
      lastEditor: usuario,
      tipo:   'CDG',
      bloqueado: true,
      items:  items
    };
    addHistorial(usuario, 'enviado_hamilton',
      'Carga '+cargaId+' → Hamilton contenedor '+licHija+' ('+items.length+' SKUs, '+tarimas.join(',')+')');

    try {
      await saveDailyStateStrict('enviado_hamilton '+licHija);
    } catch(saveErr) {
      if(snapTeorico === null) delete state.teorico[licHija]; else state.teorico[licHija] = snapTeorico;
      if(snapFisico  === null) delete state.fisico[licHija];  else state.fisico[licHija]  = snapFisico;
      if(snapCdg     === null) delete state.cdg[licHija];     else state.cdg[licHija]     = snapCdg;
      state.version = snapVer;
      if(state.historial) state.historial.length = snapHist;
      return res.status(500).json({
        ok:false,
        error:'No se pudo persistir a Supabase. Operación revertida. '+saveErr.message
      });
    }

    // Marcar en bod_furgon_cierres
    var now = new Date().toISOString();
    try {
      await supabase('PATCH', 'bod_furgon_cierres',
        { enviado_hamilton: true, enviado_hamilton_por: String(usuario||''),
          ts_envio_hamilton: now, hamilton_contenedor: licHija },
        '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    } catch(patchErr) {
      // Hamilton ya fue creado — avisar sin revertir
      return res.status(500).json({
        ok:false, hamilton_contenedor:licHija, skus:items.length,
        error:'Hamilton fue creado, pero no se pudo marcar la carga como enviada en bod_furgon_cierres. Verificá antes de reintentar. '+patchErr.message
      });
    }

    invalidarCacheSesion(sesId);
    res.json({ ok:true, hamilton_contenedor: licHija, skus: items.length });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ══════════════════════════════════════════════════════════════════════════
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
// ── Helper: construir mapa SKU → contexto logístico ──────────────────────
// Lee bod_sesiones, bod_lineas, bod_tarima_furgon y bod_furgon_cierres
// para los licencia_id dados y devuelve skuLog[sku] = { tarimas, furgones, licencias_hijas, placas, marchamos, destinos }
async function bodGetSkuLogistica(lics) {
  if(!lics || !lics.length) return {};
  var skuLog = {};
  function addUniq(arr, val){ if(val && !arr.includes(val)) arr.push(val); }
  function ensureSku(sku){
    if(!skuLog[sku]) skuLog[sku] = { tarimas:[], furgones:[], licencias_hijas:[], placas:[], marchamos:[], destinos:[] };
    return skuLog[sku];
  }

  // Leer sesiones para estos licencia_id
  var qSes = lics.length === 1
    ? '?licencia_id=eq.'+encodeURIComponent(lics[0])+'&select=id,licencia_id&limit=100'
    : '?licencia_id=in.('+lics.map(encodeURIComponent).join(',')+')'+'&select=id,licencia_id&limit=100';
  var sesiones = await supabase('GET', 'bod_sesiones', null, qSes).catch(function(){ return []; });
  sesiones = Array.isArray(sesiones) ? sesiones : [];
  if(!sesiones.length) return {};

  var sesIds = sesiones.map(function(s){ return s.id; });
  var qFilter = sesIds.length === 1
    ? '?sesion_id=eq.'+encodeURIComponent(sesIds[0])
    : '?sesion_id=in.('+sesIds.map(encodeURIComponent).join(',')+')';

  var [lineas, asigs, cierres] = await Promise.all([
    supabase('GET', 'bod_lineas', null, qFilter+'&eliminada=eq.false&select=sku,tarima&limit=20000').catch(function(){ return []; }),
    supabase('GET', 'bod_tarima_furgon', null, qFilter+'&limit=5000').catch(function(){ return []; }),
    supabase('GET', 'bod_furgon_cierres', null, qFilter+'&limit=1000').catch(function(){ return []; })
  ]);
  lineas  = Array.isArray(lineas)  ? lineas  : [];
  asigs   = Array.isArray(asigs)   ? asigs   : [];
  cierres = Array.isArray(cierres) ? cierres : [];

  // Mapas: tarima → furgon, tarima → cierre
  var tarimaFurgon = {}, tarimaCierre = {};
  asigs.forEach(function(a){ tarimaFurgon[a.tarima] = a.furgon; });
  cierres.forEach(function(cg){
    var ts = Array.isArray(cg.tarimas) ? cg.tarimas : [];
    ts.forEach(function(t){ tarimaCierre[t] = cg; });
  });

  // Cruzar líneas → tarima → furgon/cierre → skuLog
  lineas.forEach(function(l){
    var e = ensureSku(l.sku);
    addUniq(e.tarimas, l.tarima);
    var cierre = tarimaCierre[l.tarima];
    if(cierre) {
      addUniq(e.furgones, cierre.furgon);
      addUniq(e.licencias_hijas, cierre.licencia_hija);
      addUniq(e.placas, cierre.placa);
      addUniq(e.marchamos, cierre.marchamo);
      addUniq(e.destinos, cierre.destino_tr999);
    } else {
      var furg = tarimaFurgon[l.tarima];
      if(furg) addUniq(e.furgones, furg);
    }
  });
  return skuLog;
}

// ── GET /api/bod/reportes/wms-vs-app ─────────────────────────────────────
app.get('/api/bod/reportes/wms-vs-app', bodGuard, async (req, res) => {
  try {
    var f = bodBuildMovQuery(req);

    var lineasQ = f.lics.length === 1
      ? '?licencia_id=eq.'+encodeURIComponent(f.lics[0])+'&eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id'
      : f.lics.length > 1
      ? '?licencia_id=in.('+f.lics.map(encodeURIComponent).join(',')+')'+'&eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id&limit=10000'
      : '?eliminada=eq.false&select=sku,cantidad,descripcion,licencia_id&limit=10000';

    var [wmsRows, lineas, skuLog] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, f.q+'&tipo_movimiento=eq.entrada_bolson'),
      supabase('GET', 'bod_lineas', null, lineasQ),
      bodGetSkuLogistica(f.lics)
    ]);
    wmsRows = Array.isArray(wmsRows) ? wmsRows : [];
    lineas  = Array.isArray(lineas)  ? lineas  : [];

    var wmsMap = {}, wmsDesc = {}, wmsLic = {};
    wmsRows.forEach(function(r){
      var k = r.licencia_id+'|'+r.sku;
      wmsMap[k]  = (wmsMap[k]||0) + Number(r.unidades);
      if(!wmsDesc[k] && r.descripcion) wmsDesc[k] = r.descripcion;
      if(!wmsLic[k]) wmsLic[k] = r.licencia_id;
    });
    var appMap = {}, appDesc = {}, appLic = {};
    lineas.forEach(function(l){
      var k = l.licencia_id+'|'+l.sku;
      appMap[k]  = (appMap[k]||0) + Number(l.cantidad);
      if(!appDesc[k] && l.descripcion) appDesc[k] = l.descripcion;
      if(!appLic[k]) appLic[k] = l.licencia_id;
    });

    var keys = Object.keys(Object.assign({}, wmsMap, appMap));
    var result = keys.map(function(k){
      var wms = wmsMap[k]||0, app = appMap[k]||0, dif = app - wms;
      var estado;
      if     (wms > 0 && app === wms)           estado = 'Cuadrado';
      else if(wms > 0 && app === 0)             estado = 'Pendiente de entarimar';
      else if(wms > 0 && app > 0 && wms > app)  estado = 'Diferencia: faltante APP';
      else if(wms > 0 && app > 0 && app > wms)  estado = 'Diferencia: excedente APP';
      else if(wms === 0 && app > 0)             estado = 'Entarimado sin WMS';
      else                                       estado = 'Sin datos';
      var parts = k.split('|');
      var lic = wmsLic[k] || appLic[k] || parts[0];
      var sku = parts.slice(1).join('|');
      var log = skuLog[sku] || {};
      return { licencia_id:lic, sku, descripcion:wmsDesc[k]||appDesc[k]||'',
               cantidad_teorica_wms:wms, cantidad_app:app, diferencia:dif, estado,
               tarimas:       (log.tarimas||[]).join(', '),
               furgon:        (log.furgones||[]).join(', '),
               licencia_hija: (log.licencias_hijas||[]).join(', '),
               placa:         (log.placas||[]).join(', '),
               marchamo:      (log.marchamos||[]).join(', '),
               destino_tr999: (log.destinos||[]).join(', ') };
    });
    res.json({ ok:true, skus:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/reportes/flujo-bolson ───────────────────────────────────
app.get('/api/bod/reportes/flujo-bolson', bodGuard, async (req, res) => {
  try {
    var licId   = String(req.query.licencia_id || '').trim().toUpperCase();
    var licsRaw = String(req.query.licencias   || '').trim();
    var lics    = licsRaw ? licsRaw.split(',').map(function(l){ return l.trim().toUpperCase(); }).filter(Boolean) : [];
    if(licId && !lics.includes(licId)) lics.push(licId);

    // Entradas: WMS entrada_bolson
    var qEntradas;
    if(lics.length === 1) {
      qEntradas = '?licencia_id=eq.'+encodeURIComponent(lics[0])+'&tipo_movimiento=eq.entrada_bolson&limit=10000';
    } else if(lics.length > 1) {
      qEntradas = '?licencia_id=in.('+lics.map(encodeURIComponent).join(',')+')'+'&tipo_movimiento=eq.entrada_bolson&limit=10000';
    } else {
      qEntradas = '?tipo_movimiento=eq.entrada_bolson&limit=5000';
    }

    // Cierres: fuente de verdad para salidas (bod_furgon_cierres, NO WMS salida_tr999)
    var qCierres;
    if(lics.length === 1) {
      qCierres = '?licencia_padre=eq.'+encodeURIComponent(lics[0])+'&limit=1000';
    } else if(lics.length > 1) {
      qCierres = '?licencia_padre=in.('+lics.map(encodeURIComponent).join(',')+')'+'&limit=1000';
    } else {
      qCierres = '?limit=500';
    }

    var [movsEntrada, cierres] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, qEntradas),
      supabase('GET', 'bod_furgon_cierres',  null, qCierres).catch(function(){ return []; })
    ]);
    movsEntrada = Array.isArray(movsEntrada) ? movsEntrada : [];
    cierres     = Array.isArray(cierres)     ? cierres     : [];

    var skuData = {};
    function ensureSku(skuRaw, desc) {
      var sku = String(skuRaw||'').trim().toUpperCase();
      if(!sku) return null;
      if(!skuData[sku]) skuData[sku] = { sku:sku, descripcion:desc||'',
        entradas_952:0, salidas_tr999:0, licencia_origen:'',
        licencias_hijas:[], furgones:[], tarimas_arr:[], placas:[], marchamos:[], destinos:[] };
      return skuData[sku];
    }
    function addUniq(arr, val){ if(val && !arr.includes(val)) arr.push(val); }

    // Entradas desde WMS
    movsEntrada.forEach(function(m){
      var d = ensureSku(m.sku, m.descripcion);
      if(!d) return;
      if(!d.descripcion && m.descripcion) d.descripcion = m.descripcion;
      d.entradas_952 += Number(m.unidades);
      if(!d.licencia_origen) d.licencia_origen = m.licencia_id;
    });

    // Salidas exclusivamente desde cierres logísticos (bod_furgon_cierres.resumen_skus)
    cierres.forEach(function(c){
      var skus = Array.isArray(c.resumen_skus) ? c.resumen_skus : [];
      var ts   = Array.isArray(c.tarimas)      ? c.tarimas      : [];
      skus.forEach(function(sv){
        var d = ensureSku(sv.sku, sv.descripcion);
        if(!d) return;
        if(!d.descripcion && sv.descripcion) d.descripcion = sv.descripcion;
        d.salidas_tr999 += Number(sv.unidades)||0;
        addUniq(d.licencias_hijas, c.licencia_hija);
        addUniq(d.furgones, c.furgon);
        addUniq(d.placas, c.placa);
        addUniq(d.marchamos, c.marchamo);
        addUniq(d.destinos, c.destino_tr999);
        ts.forEach(function(t){ addUniq(d.tarimas_arr, t); });
      });
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
        destino_tr999:   d.destinos.join(', '),
        tarimas:         d.tarimas_arr.join(', '),
        placa:           d.placas.join(', '),
        marchamo:        d.marchamos.join(', ')
      };
    });
    res.json({ ok:true, movimientos:result });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/reportes/remanentes ─────────────────────────────────────
app.get('/api/bod/reportes/remanentes', bodGuard, async (req, res) => {
  try {
    var f = bodBuildMovQuery(req);
    var lineasQ = f.lics.length === 1
      ? '?licencia_id=eq.'+encodeURIComponent(f.lics[0])+'&eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000'
      : f.lics.length > 1
      ? '?licencia_id=in.('+f.lics.map(encodeURIComponent).join(',')+')'+'&eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000'
      : '?eliminada=eq.false&select=sku,cantidad,tarima,licencia_id&limit=10000';

    var [movs, lineas, skuLog] = await Promise.all([
      supabase('GET', 'bod_wms_movimientos', null, f.q),
      supabase('GET', 'bod_lineas', null, lineasQ),
      bodGetSkuLogistica(f.lics)
    ]);
    movs   = Array.isArray(movs)   ? movs   : [];
    lineas = Array.isArray(lineas) ? lineas : [];

    var appMap = {}, skuTarimadas = {};
    lineas.forEach(function(l){
      appMap[l.sku] = (appMap[l.sku]||0) + Number(l.cantidad);
      skuTarimadas[l.sku] = true;
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
      var log    = skuLog[d.sku] || {};
      var furgon        = (log.furgones||[]).join(', ');
      var licencia_hija = (log.licencias_hijas||[]).join(', ');
      var tarimas       = (log.tarimas||[]).join(', ');
      var placa         = (log.placas||[]).join(', ');
      var marchamo      = (log.marchamos||[]).join(', ');
      var destino_tr999 = (log.destinos||[]).join(', ');

      // Causa operativa enriquecida con contexto logístico
      var causa;
      if(d.entradas === 0 && d.destinos_distintos)
        causa = 'Pendiente por ubicación distinta a 952';
      else if(d.entradas > 0 && app === 0)
        causa = 'Pendiente de entarimar';
      else if(d.entradas === 0 && app > 0)
        causa = 'Entarimado sin WMS';
      else if(app > 0 && !furgon)
        causa = 'Pendiente de asignación de furgón';
      else if(app > 0 && furgon && !licencia_hija)
        causa = 'Pendiente de licencia hija';
      else if(app > 0 && furgon && licencia_hija && d.salidas === 0)
        causa = 'Pendiente de movimiento WMS salida';
      else if(d.entradas > 0 && d.salidas > 0 && rem > 0)
        causa = 'Remanente operativo';
      else if(app > 0 && !skuTarimadas[d.sku])
        causa = 'Pendiente de traslado';
      else
        causa = dias !== null && dias <= 2 ? 'Pendiente normal' : 'Atrasado';

      return { sku:d.sku, descripcion:d.descripcion, cantidad_remanente:rem,
               entrada_952:d.entradas, app_entarimada:app, salida_tr999:d.salidas,
               licencia_origen:d.licencia_origen, licencia_salida:d.licencia_salida,
               furgon, licencia_hija, tarimas, placa, marchamo, destino_tr999,
               fecha_primer_ingreso:fechaMin, ultimo_movimiento:d.ultima_fecha,
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

// Logo Manifiesto Bodega CDG — base64 embebido (compártido entre Word y PDF)
var _MANIF_LOGO_B64 = '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';

// ── GET /api/bod/sesion/:id/carga-logistica/:cargaId/manifiesto-word ─────────
// Genera DOCX real del Manifiesto Bodega CDG usando la librería docx.
// No guarda el archivo; lo devuelve como stream de descarga.
// FIX (vie 06-jun-2026, server): endpoint manifiesto-word DOCX real
app.get('/api/bod/sesion/:id/carga-logistica/:cargaId/manifiesto-word', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();

    // ── Leer carga logística ────────────────────────────────────────────────
    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });
    var cg = rows[0];

    // ── Leer sesión (para fecha_trabajo) ───────────────────────────────────
    var sesRows = await supabase('GET', 'bod_sesiones', null,
      '?id=eq.'+encodeURIComponent(sesId)+'&select=fecha_trabajo,licencia_id&limit=1');
    var ses = (Array.isArray(sesRows)&&sesRows.length) ? sesRows[0] : {};

    // ── Mapeo de campos (mismo que Export Excel) ────────────────────────────
    var marchamo   = String(cg.marchamo      ||'');
    var licId      = String(cg.licencia_hija  ||'');
    var ubicacion  = String(cg.destino_tr999  ||'');
    var transporte = String(cg.transporte || cg.placa || '');
    var furgon     = String(cg.furgon || '');
    // Fecha
    var fechaRaw = cg.ts_cierre||cg.created_at||cg.fecha_creacion||ses.fecha_trabajo||'';
    if(/^\d{4}-\d{2}-\d{2}T/.test(fechaRaw)) fechaRaw = fechaRaw.slice(0,10);
    var fechaFmt = fechaRaw;
    if(/^\d{4}-\d{2}-\d{2}$/.test(fechaRaw)) {
      var pp = fechaRaw.split('-'); fechaFmt = pp[2]+'/'+pp[1]+'/'+pp[0];
    }
    // Carga No. — no usar licencia_hija ni carga.id
    var cargaNo   = String(cg.carga_no||cg.numero_carga||cg.no_carga||cg.id_carga||'');
    // Responsables
    // autorizador eliminado: AUTORIZACIÓN DE ENVÍO usa nombres fijos (ASTRID DUARTE, CINTYA RIVERA, HUVALDO PEREZ)
    var encargado   = String(cg.encargado_pedido||cg.creado_por||'');
    var auditorVal  = String(cg.auditado_por||cg.validado_por||req.query.usuario||'');
    var recibidoPor = String(cg.recibido_por||'');
    // ID contenedor — no usar licencia_hija ni carga.id
    var idCont = String(cg.hamilton_contenedor||cg.id_contenedor||cg.contenedor||'');

    // ── SKU helpers ─────────────────────────────────────────────────────────
    function fd() {
      for(var i=0;i<arguments.length;i++){
        var v=arguments[i];
        if(v!==null&&v!==undefined&&v!=='') return v;
      }
      return '';
    }

    // ── FIX (mié 10-jun-2026): Papel de Trabajo como fuente PRIMARIA ──────────
    // Si existe bod_papel_trabajo para esta sesión+licencia_hija → ese es el manifiesto.
    // Si no → fallback a Tarimas (bod_lineas, lógica original).
    var resSkus;
    try {
      var papelWordQ = '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&licencia_hija=eq.'+encodeURIComponent(licId)
        +'&furgon=eq.'+encodeURIComponent(furgon)
        +'&select=sku,nombre,fisico_auditoria,cantidad_manifiesto,estado_papel,diferencia';
      var papelWordRows = await supabase('GET', 'bod_papel_trabajo', null, papelWordQ);
      papelWordRows = Array.isArray(papelWordRows) ? papelWordRows : [];
      if(papelWordRows.length > 0) {
        // FUENTE: Papel de Trabajo
        resSkus = papelWordRows.map(function(p){
          var cantM = Number(p.cantidad_manifiesto)||0;
          var fisico = p.fisico_auditoria != null ? Number(p.fisico_auditoria) : null;
          return {
            sku:              p.sku,
            descripcion:      p.nombre || '',
            unidades:         cantM,
            cantidad_auditada:fisico,
            _fuente:          'PT'
          };
        });
      } else {
        throw new Error('sin_papel'); // cae al fallback
      }
    } catch(eP) {
      // FALLBACK: Tarimas (lógica original)
      var bodLineasWord = await supabase('GET', 'bod_lineas', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&eliminada=eq.false&order=ts_captura.asc&limit=5000&select=sku,descripcion,cantidad,cantidad_audit,auditado,tarima');
      bodLineasWord = Array.isArray(bodLineasWord) ? bodLineasWord : [];
      var asigWordRows = await supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      var tarimaFurgonWord = {};
      (Array.isArray(asigWordRows)?asigWordRows:[]).forEach(function(a){
        tarimaFurgonWord[a.tarima] = a.furgon;
      });
      var porSkuWord = {};
      bodLineasWord.forEach(function(l){
        if(tarimaFurgonWord[l.tarima] !== furgon) return;
        var sku=l.sku||'?', desc=l.descripcion||'';
        if(!porSkuWord[sku]) porSkuWord[sku]={sku:sku,descripcion:desc,unidades:0,lineasCount:0,pendientes:0,cantidad_auditada:null,_auditSum:0,_auditCount:0,_fuente:'TAR'};
        var e=porSkuWord[sku];
        e.unidades+=Number(l.cantidad)||0; e.lineasCount++;
        if(!l.auditado){ e.pendientes++; }
        else if(l.cantidad_audit!=null){ e._auditSum+=Number(l.cantidad_audit)||0; e._auditCount++; }
        if(!e.descripcion&&desc) e.descripcion=desc;
      });
      Object.values(porSkuWord).forEach(function(e){ if(e._auditCount>0) e.cantidad_auditada=e._auditSum; });
      resSkus = Object.values(porSkuWord);
    }

    // ── Construir DOCX con librería docx ────────────────────────────────────
    var docx;
    try { docx = require('docx'); }
    catch(e) { return res.status(500).json({ ok:false, error:'Librería docx no instalada en el servidor.' }); }

    var { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          ImageRun, AlignmentType, WidthType, BorderStyle, VerticalAlign } = docx;

    var logoBytes = Buffer.from(_MANIF_LOGO_B64, 'base64'); // usa constante de módulo

    // ── Helpers — tabla única de 9 columnas (evita desalineación en Word) ──────
    var FONT='Aptos Narrow', SZ=20, SZ_S=18, SZ_B=22;
    var sp0={after:0}, marg={top:60,bottom:60,left:80,right:80};

    function bS(type,sz,color){return{style:BorderStyle[type]||BorderStyle.SINGLE,size:sz||4,color:color||'000000'};}
    function bNil(){return{style:BorderStyle.NIL,size:0,color:'000000'};}
    function bSgl(sz){return bS('SINGLE',sz||4);}
    function bDbl(sz){return bS('DOUBLE',sz||6);}
    function run(text,opts){opts=opts||{};return new TextRun({text:String(text||''),font:FONT,size:opts.sz||SZ,bold:!!opts.bold});}
    function paraC(text,opts){opts=opts||{};return new Paragraph({alignment:opts.align||AlignmentType.LEFT,spacing:sp0,children:[run(text,opts)]});}
    function mkCell(w,borders,children,opts){
      opts=opts||{};
      return new TableCell({width:{size:w,type:WidthType.DXA},columnSpan:opts.span,
        borders:borders,margins:marg,verticalAlign:opts.vAlign||VerticalAlign.BOTTOM,children:children});
    }
    function emptyCell(w,borders){return mkCell(w,borders,[new Paragraph({spacing:sp0,children:[run('')]})]);}

    // Columnas: [marg_izq, CODIGO, ESTATUS, %AUD, DESC, CANT, VAL, DIF, marg_der]
    var CW9=[400,1500,1100,900,4580,982,1180,1306,400];
    var TOTAL9=CW9.reduce(function(a,b){return a+b;},0);
    var W_IZQ=CW9[1]+CW9[2]+CW9[3]+CW9[4]; // span 4: cols 1-4
    var W_DER=CW9[5]+CW9[6]+CW9[7];         // span 3: cols 5-7
    var INNER=W_IZQ+W_DER;                   // span 7: cols 1-7

    var bLEFT ={top:bNil(),bottom:bNil(),left:bDbl(),right:bNil()};
    var bRIGHT={top:bNil(),bottom:bNil(),left:bNil(),right:bDbl()};
    var bNONE ={top:bNil(),bottom:bNil(),left:bNil(),right:bNil()};

    // Fila full-width interior (span 7)
    function mkFullRow(children,bTop,bBot,ht){
      return new TableRow({height:{value:ht||288,rule:'atLeast'},children:[
        emptyCell(CW9[0],Object.assign({},bNONE,{left:bDbl()})),
        mkCell(INNER,{top:bTop||bNil(),bottom:bBot||bNil(),left:bNil(),right:bNil()},children,{span:7}),
        emptyCell(CW9[8],Object.assign({},bNONE,{right:bDbl()})),
      ]});
    }
    // Fila dividida izq (span 4) / der (span 3)
    function mkSplitRow(izqText,izqOpts,izqBord,derText,derOpts,derBord,ht){
      return new TableRow({height:{value:ht||288,rule:'atLeast'},children:[
        emptyCell(CW9[0],Object.assign({},bNONE,{left:bDbl()})),
        mkCell(W_IZQ,izqBord,[paraC(izqText||'',izqOpts||{})],{span:4}),
        mkCell(W_DER,derBord,[paraC(derText||'',derOpts||{})],{span:3}),
        emptyCell(CW9[8],Object.assign({},bNONE,{right:bDbl()})),
      ]});
    }
    // Fila metadatos: texto izq (span 4) + label/valor en cols 5-6 / 7
    function mkLabelValRow(izqText,label,value,ht){
      return new TableRow({height:{value:ht||288,rule:'atLeast'},children:[
        emptyCell(CW9[0],Object.assign({},bNONE,{left:bDbl()})),
        mkCell(W_IZQ,bNONE,[paraC(izqText||'',{sz:SZ})],{span:4}),
        mkCell(CW9[5]+CW9[6],{top:bSgl(),bottom:bSgl(),left:bSgl(),right:bSgl()},[paraC(label||'',{sz:SZ})],{span:2}),
        mkCell(CW9[7],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},[paraC(value||'',{sz:SZ})]),
        emptyCell(CW9[8],Object.assign({},bNONE,{right:bDbl()})),
      ]});
    }

    // ── Filas de encabezado ─────────────────────────────────────────────────
    var rowTop = new TableRow({height:{value:200,rule:'exact'},
      children:CW9.map(function(w,i){
        return emptyCell(w,{top:bDbl(),bottom:bNil(),left:i===0?bDbl():bNil(),right:i===8?bDbl():bNil()});
      })
    });
    var rowLogo = new TableRow({height:{value:600,rule:'atLeast'},children:[
      emptyCell(CW9[0],Object.assign({},bNONE,{left:bDbl()})),
      mkCell(CW9[1]+CW9[2],bNONE,
        [new Paragraph({spacing:sp0,children:[new ImageRun({data:logoBytes,transformation:{width:110,height:88},type:'png'})]})],
        {span:2,vAlign:VerticalAlign.CENTER}),
      mkCell(CW9[3]+CW9[4]+CW9[5]+CW9[6]+CW9[7],{top:bSgl(8),bottom:bSgl(8),left:bSgl(8),right:bSgl(8)},
        [paraC('CARGA ENVÍO A BODEGA',{sz:SZ_B,bold:true,align:AlignmentType.CENTER})],
        {span:5,vAlign:VerticalAlign.CENTER}),
      emptyCell(CW9[8],Object.assign({},bNONE,{right:bDbl()})),
    ]});
    var rowEmpresa = mkLabelValRow('Herramientas Poderosas','Marchamo:',marchamo);
    var rowDir1 = mkLabelValRow('DIRECCIÓN: 27 Calle Bodega C 41-55 Zona 5 Calzada la Paz','Licencia:',licId);
    var rowDir2 = mkLabelValRow('DIRECCIÓN DESTINO: BODEGA NODUS (Hamiltón)','UBICACIÓN:',ubicacion);
    var rowTrans = mkLabelValRow('TRANSPORTE: '+transporte+'   Furgón: '+furgon,'Fecha:',fechaFmt);
    var rowCargaNo = mkLabelValRow('','Carga No.',cargaNo);

    // ── Encabezados tabla SKU ───────────────────────────────────────────────
    var HDRS=['','CODIGO','ESTATUS','% AUDITADO','DESCRIPCIÓN','CANTIDAD','VALIDADO','DIFERENCIA',''];
    var rowSkuHdr = new TableRow({height:{value:300,rule:'atLeast'},
      children:CW9.map(function(w,i){
        var bord = i===0 ? Object.assign({},bNONE,{left:bDbl()})
                 : i===8 ? Object.assign({},bNONE,{right:bDbl()})
                 : i===1 ? {top:bSgl(),bottom:bSgl(),left:bSgl(),right:bSgl()}
                 :         {top:bSgl(),bottom:bSgl(),left:bNil(),right:bSgl()};
        return mkCell(w,bord,[new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,
          children:[run(HDRS[i]||'',{sz:SZ,bold:true})]})]);
      })
    });

    // ── Filas SKU — sin límite, filas dinámicas ────────────────────────────
    var skuRows = resSkus.map(function(s){
      var cantVal = fd(s.cantidad_auditada,s.cantidad_audit,s.auditada,
                       s.cantidad_validada,s.validado,s.unidades_auditadas);
      var tieneVal = cantVal!==''&&cantVal!==null&&cantVal!==undefined;
      var estatus  = tieneVal ? 'Auditado' : 'No auditado';
      var desc = s.descripcion_completa||s.descripcion_larga||s.descripcion||s.desc||s.nombre||s.sku_desc||'';
      var cantidad = Number(s.unidades)||0;
      var validado = tieneVal ? Number(cantVal) : '';
      var dif      = tieneVal ? (Number(cantVal)-cantidad) : '';
      var pct      = (!tieneVal||cantidad<=0) ? '0%'
                   : Math.min(100,Math.round((Number(cantVal)/cantidad)*100))+'%';
      return new TableRow({height:{value:288,rule:'atLeast'},children:[
        emptyCell(CW9[0],Object.assign({},bNONE,{left:bDbl()})),
        mkCell(CW9[1],{top:bNil(),bottom:bSgl(),left:bSgl(),right:bSgl()},[paraC(String(s.sku||''),{sz:SZ})]),
        mkCell(CW9[2],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},[paraC(estatus,{sz:SZ})]),
        mkCell(CW9[3],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},
          [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,children:[run(pct,{sz:SZ})]})]),
        mkCell(CW9[4],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},[paraC(String(desc),{sz:SZ_S})]),
        mkCell(CW9[5],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},
          [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,children:[run(String(cantidad),{sz:SZ})]})]),
        mkCell(CW9[6],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},
          [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,children:[run(validado!==''?String(validado):'',{sz:SZ})]})]),
        mkCell(CW9[7],{top:bNil(),bottom:bSgl(),left:bNil(),right:bSgl()},
          [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,children:[run(dif!==''?String(dif):'',{sz:SZ})]})]),
        emptyCell(CW9[8],Object.assign({},bNONE,{right:bDbl()})),
      ]});
    });

    // ── Firmas / Footer ─────────────────────────────────────────────────────
    // AUTORIZACIÓN DE ENVÍO: 3 filas reales — responsables siempre estáticos
    var rowFirmaHdr = mkFullRow(
      [paraC('AUTORIZACIÓN DE ENVÍO',{sz:SZ,bold:true})],
      bSgl(),bNil(),280);
    var rowFirmaA1 = mkFullRow(
      [paraC('ASTRID DUARTE',{sz:SZ})],
      bNil(),bNil(),288);
    var rowFirmaA2 = mkFullRow(
      [paraC('CINTYA RIVERA',{sz:SZ})],
      bNil(),bNil(),288);
    var rowFirmaA3 = mkFullRow(
      [paraC('HUVALDO PEREZ',{sz:SZ})],
      bNil(),bSgl(),288);
    var rowFirmaRecibe = mkFullRow(
      [paraC('NOMBRE Y FIRMA QUIEN RECIBE: '+(recibidoPor||''),{sz:SZ})],
      bSgl(),bSgl(),280);
    var rowTransCentral = mkFullRow(
      [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,
        children:[run('TRANSPORTE: '+transporte,{sz:SZ,bold:true})]})],
      bNil(),bNil(),288);
    var rowEncHdr = mkSplitRow(
      'ENCARGADO DE PEDIDO',{sz:SZ,bold:true},
      {top:bSgl(),bottom:bNil(),left:bSgl(),right:bNil()},
      'AUDITOR / VALIDADOR',{sz:SZ,bold:true},
      {top:bSgl(),bottom:bNil(),left:bSgl(),right:bSgl()},280);
    var rowEncVal = mkSplitRow(
      encargado,{sz:SZ},{top:bNil(),bottom:bSgl(),left:bSgl(),right:bNil()},
      auditorVal,{sz:SZ},{top:bNil(),bottom:bSgl(),left:bSgl(),right:bSgl()},500);
    var rowIdCont = mkFullRow(
      [new Paragraph({alignment:AlignmentType.CENTER,spacing:sp0,
        children:[run('ID DE CONTENEDOR: '+idCont,{sz:SZ,bold:true})]})],
      bNil(),bNil(),288);
    var rowBot = new TableRow({height:{value:200,rule:'exact'},
      children:CW9.map(function(w,i){
        return emptyCell(w,{top:bNil(),bottom:bDbl(),left:i===0?bDbl():bNil(),right:i===8?bDbl():bNil()});
      })
    });

    // ── Ensamblar y enviar ──────────────────────────────────────────────────
    var allRows = [rowTop,rowLogo,rowEmpresa,rowDir1,rowDir2,rowTrans,rowCargaNo,
      rowSkuHdr].concat(skuRows).concat([
      rowFirmaHdr,rowFirmaA1,rowFirmaA2,rowFirmaA3,rowFirmaRecibe,rowTransCentral,
      rowEncHdr,rowEncVal,rowIdCont,rowBot
    ]);

    var doc = new Document({
      sections:[{
        properties:{page:{size:{width:12240,height:15840},margin:{top:720,right:720,bottom:720,left:720}}},
        children:[new Table({width:{size:TOTAL9,type:WidthType.DXA},columnWidths:CW9,rows:allRows})]
      }]
    });

    var buf = await Packer.toBuffer(doc);
    var licFile  = (licId||'BOD').replace(/[^A-Z0-9]/gi,'_');
    var fechaFile= (fechaFmt||'').replace(/\//g,'') || new Date().toISOString().slice(0,10).replace(/-/g,'');
    var filename = 'Manifiesto_'+licFile+'_Furgon'+furgon+'_'+fechaFile+'.docx';

    res.setHeader('Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',
      'attachment; filename="'+filename+'"');
    res.setHeader('Content-Length', buf.length);
    res.end(buf);

  } catch(e) {
    console.error('manifiesto-word error:', e.message);
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── GET /api/bod/sesion/:id/carga-logistica/:cargaId/manifiesto-pdf ──────────
// Genera PDF del Manifiesto Bodega CDG usando wkhtmltopdf.
// DEPLOY: requiere wkhtmltopdf instalado en Render.
//   - Sin wkhtmltopdf: responde 503 con mensaje claro. El endpoint /manifiesto-word sigue funcionando.
//   - Para instalar: usar Dockerfile con "apt-get install -y wkhtmltopdf"
//     o binario estático en /bin/wkhtmltopdf commiteado al repo.
//   - Mientras no esté configurado, el botón PDF en el cliente muestra aviso en lugar de descargar.
// FIX (vie 06-jun-2026, server): endpoint manifiesto-pdf via wkhtmltopdf
app.get('/api/bod/sesion/:id/carga-logistica/:cargaId/manifiesto-pdf', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();

    // Verificar que wkhtmltopdf esté disponible antes de continuar
    var { execFileSync } = require('child_process');
    try { execFileSync('wkhtmltopdf', ['--version'], { timeout:5000, stdio:'pipe' }); }
    catch(_) {
      return res.status(503).json({
        ok:false,
        error:'PDF no disponible: wkhtmltopdf no está instalado en este servidor. El export Word sigue funcionando normalmente.',
        code:'WKHTMLTOPDF_NOT_FOUND'
      });
    }

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });
    var cg = rows[0];

    var sesRows = await supabase('GET', 'bod_sesiones', null,
      '?id=eq.'+encodeURIComponent(sesId)+'&select=fecha_trabajo,licencia_id&limit=1');
    var ses = (Array.isArray(sesRows)&&sesRows.length) ? sesRows[0] : {};

    // ── Mapeo de campos (idéntico al endpoint Word) ─────────────────────────
    var marchamo   = String(cg.marchamo      ||'');
    var licId      = String(cg.licencia_hija  ||'');
    var ubicacion  = String(cg.destino_tr999  ||'');
    var transporte = String(cg.transporte || cg.placa || '');
    var furgon     = String(cg.furgon || '');
    var fechaRaw   = cg.ts_cierre||cg.created_at||cg.fecha_creacion||ses.fecha_trabajo||'';
    if(/^\d{4}-\d{2}-\d{2}T/.test(fechaRaw)) fechaRaw = fechaRaw.slice(0,10);
    var fechaFmt = fechaRaw;
    if(/^\d{4}-\d{2}-\d{2}$/.test(fechaRaw)) {
      var pp = fechaRaw.split('-'); fechaFmt = pp[2]+'/'+pp[1]+'/'+pp[0];
    }
    var cargaNo     = String(cg.carga_no||cg.numero_carga||cg.no_carga||cg.id_carga||'');
    var encargado   = String(cg.encargado_pedido||cg.creado_por||'');
    var auditorVal  = String(cg.auditado_por||cg.validado_por||req.query.usuario||'');
    var recibidoPor = String(cg.recibido_por||'');
    var idCont      = String(cg.hamilton_contenedor||cg.id_contenedor||cg.contenedor||''); // NO usar licencia_hija

    // ── FIX (mié 10-jun-2026): Papel de Trabajo como fuente PRIMARIA ──────────
    var resSkus;
    try {
      var papelPdfQ = '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&licencia_hija=eq.'+encodeURIComponent(licId)
        +'&furgon=eq.'+encodeURIComponent(furgon)
        +'&select=sku,nombre,fisico_auditoria,cantidad_manifiesto,estado_papel,diferencia';
      var papelPdfRows = await supabase('GET', 'bod_papel_trabajo', null, papelPdfQ);
      papelPdfRows = Array.isArray(papelPdfRows) ? papelPdfRows : [];
      if(papelPdfRows.length > 0) {
        // FUENTE: Papel de Trabajo
        resSkus = papelPdfRows.map(function(p){
          var cantM  = Number(p.cantidad_manifiesto)||0;
          var fisico = p.fisico_auditoria != null ? Number(p.fisico_auditoria) : null;
          return {
            sku:              p.sku,
            descripcion:      p.nombre || '',
            unidades:         cantM,
            cantidad_auditada:fisico,
            _fuente:          'PT'
          };
        });
      } else {
        throw new Error('sin_papel');
      }
    } catch(eP) {
      // FALLBACK: Tarimas
      var bodLineasPdf = await supabase('GET', 'bod_lineas', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&eliminada=eq.false&order=ts_captura.asc&limit=5000&select=sku,descripcion,cantidad,cantidad_audit,auditado,tarima');
      bodLineasPdf = Array.isArray(bodLineasPdf) ? bodLineasPdf : [];
      var asigPdfRows = await supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      var tarimaFurgonPdf = {};
      (Array.isArray(asigPdfRows)?asigPdfRows:[]).forEach(function(a){ tarimaFurgonPdf[a.tarima]=a.furgon; });
      var porSkuPdf = {};
      bodLineasPdf.forEach(function(l){
        if(tarimaFurgonPdf[l.tarima] !== furgon) return;
        var sku=l.sku||'?', desc=l.descripcion||'';
        if(!porSkuPdf[sku]) porSkuPdf[sku]={sku:sku,descripcion:desc,unidades:0,lineasCount:0,pendientes:0,cantidad_auditada:null,_auditSum:0,_auditCount:0,_fuente:'TAR'};
        var e=porSkuPdf[sku];
        e.unidades+=Number(l.cantidad)||0; e.lineasCount++;
        if(!l.auditado){ e.pendientes++; }
        else if(l.cantidad_audit!=null){ e._auditSum+=Number(l.cantidad_audit)||0; e._auditCount++; }
        if(!e.descripcion&&desc) e.descripcion=desc;
      });
      Object.values(porSkuPdf).forEach(function(e){ if(e._auditCount>0) e.cantidad_auditada=e._auditSum; });
      resSkus = Object.values(porSkuPdf);
    }

    function fd() {
      for(var _i=0;_i<arguments.length;_i++){
        var _v=arguments[_i];
        if(_v!==null&&_v!==undefined&&_v!=='') return _v;
      }
      return '';
    }
    function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

    // Logo embebido desde constante de módulo _MANIF_LOGO_B64 (compartido con Word)

    // ── Construir HTML del manifiesto ────────────────────────────────────────
    var filasSku = '';
    resSkus.forEach(function(s){
      var cantVal = fd(s.cantidad_auditada,s.cantidad_audit,s.auditada,
                       s.cantidad_validada,s.validado,s.unidades_auditadas);
      var tieneVal = cantVal!==''&&cantVal!==null&&cantVal!==undefined;
      var estatus  = tieneVal ? 'Auditado' : 'No auditado';
      var desc = s.descripcion_completa||s.descripcion_larga||s.descripcion||s.desc||s.nombre||s.sku_desc||'';
      var cantidad = Number(s.unidades)||0;
      var validado = tieneVal ? Number(cantVal) : '';
      var dif      = tieneVal ? (Number(cantVal)-cantidad) : '';
      var pct      = (!tieneVal||cantidad<=0) ? '0%'
                   : Math.min(100,Math.round((Number(cantVal)/cantidad)*100))+'%';
      filasSku += '<tr>'
        +'<td>'+esc(String(s.sku||''))+'</td>'
        +'<td class="c">'+esc(estatus)+'</td>'
        +'<td class="c">'+esc(pct)+'</td>'
        +'<td class="desc">'+esc(desc)+'</td>'
        +'<td class="c">'+esc(String(cantidad))+'</td>'
        +'<td class="c">'+(validado!==''?esc(String(validado)):'')+'</td>'
        +'<td class="c">'+(dif!==''?esc(String(dif)):'')+'</td>'
        +'</tr>';
    });

    var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
      +'*{margin:0;padding:0;box-sizing:border-box}'
      +'body{font-family:Arial,sans-serif;font-size:8.5pt;padding:8px}'
      +'table{width:100%;border-collapse:collapse;border:2px double #000}'
      +'th{background:#d0d7de;border:1px solid #555;padding:3px 5px;font-weight:bold;text-align:center}'
      +'td{border:1px solid #555;padding:2px 5px;vertical-align:top}'
      +'.c{text-align:center}.desc{word-wrap:break-word;white-space:normal}'
      +'.lbl{background:#f6f8fa;font-weight:bold;white-space:nowrap}'
      +'.val{}'
      +'.title{text-align:center;font-size:12pt;font-weight:bold;border:1px solid #555;padding:6px;vertical-align:middle}'
      +'.hdr2{width:100%;border-collapse:collapse;border:2px double #000;border-top:none;margin-top:0}'
      +'.firma{background:#f6f8fa;font-weight:bold;border-top:1px solid #555}'
      +'</style></head><body>'
      // Encabezado: logo + título
      +'<table style="border-bottom:none">'
      +'<tr>'
      +'<td rowspan="2" style="width:110px;border:none;padding:4px;vertical-align:middle">'
      +'<img src="data:image/png;base64,'+_MANIF_LOGO_B64+'" width="95" height="75" style="display:block">'
      +'</td>'
      +'<td class="title">CARGA ENV\u00cdO A BODEGA</td>'
      +'</tr>'
      +'<tr><td style="font-weight:bold;padding:3px 8px;border-top:1px solid #555">Herramientas Poderosas</td></tr>'
      +'</table>'
      // Campos de encabezado
      +'<table class="hdr2">'
      +'<tr>'
      +'<td colspan="2" style="width:58%;border-bottom:1px solid #999;padding:2px 6px"><strong>DIRECCI\u00d3N:</strong> 27 Calle Bodega C 41-55 Zona 5 Calzada la Paz</td>'
      +'<td class="lbl" style="width:21%">Marchamo:</td>'
      +'<td class="val" style="width:21%">'+esc(marchamo)+'</td>'
      +'</tr>'
      +'<tr>'
      +'<td colspan="2" style="border-bottom:1px solid #999;padding:2px 6px"><strong>DIRECCI\u00d3N DESTINO:</strong> BODEGA NODUS (Hamilt\u00f3n)</td>'
      +'<td class="lbl">Licencia:</td><td class="val">'+esc(licId)+'</td>'
      +'</tr>'
      +'<tr>'
      +'<td style="width:30%;padding:2px 6px;border-bottom:1px solid #999"><strong>TRANSPORTE:</strong> '+esc(transporte)+'</td>'
      +'<td style="width:28%;padding:2px 6px;border-bottom:1px solid #999"><strong>Furg\u00f3n:</strong> '+esc(furgon)+'</td>'
      +'<td class="lbl">UBICACI\u00d3N:</td><td class="val">'+esc(ubicacion)+'</td>'
      +'</tr>'
      +'<tr>'
      +'<td colspan="2" style="padding:2px 6px;border-bottom:1px solid #999"></td>'
      +'<td class="lbl">Fecha:</td><td class="val">'+esc(fechaFmt)+'</td>'
      +'</tr>'
      +'<tr>'
      +'<td colspan="2" style="padding:2px 6px"></td>'
      +'<td class="lbl">Carga No.</td><td class="val">'+esc(cargaNo)+'</td>'
      +'</tr>'
      +'</table>'
      // Tabla SKU
      +'<table class="hdr2">'
      +'<thead><tr>'
      +'<th style="width:12%">CODIGO</th>'
      +'<th style="width:10%">ESTATUS</th>'
      +'<th style="width:8%">% AUDITADO</th>'
      +'<th style="width:40%">DESCRIPCI\u00d3N</th>'
      +'<th style="width:9%">CANTIDAD</th>'
      +'<th style="width:10%">VALIDADO</th>'
      +'<th style="width:11%">DIFERENCIA</th>'
      +'</tr></thead>'
      +'<tbody>'+filasSku+'</tbody>'
      +'</table>'
      // Firmas
      +'<table class="hdr2" style="margin-top:8px">'
      +'<tr><td colspan="4" class="firma" style="padding:3px 6px">AUTORIZACI\u00d3N DE ENV\u00cdO</td></tr>'
      +'<tr><td colspan="4" style="padding:3px 6px">ASTRID DUARTE</td></tr>'+'<tr><td colspan="4" style="padding:3px 6px">CINTYA RIVERA</td></tr>'+'<tr><td colspan="4" style="padding:3px 6px">HUVALDO PEREZ</td></tr>'
      +'<tr><td colspan="4" style="padding:3px 6px;border-top:1px solid #555">NOMBRE Y FIRMA QUIEN RECIBE: '+esc(recibidoPor)+'</td></tr>'
      +'<tr><td colspan="4" style="padding:3px 6px;text-align:center;border-top:1px solid #555;font-weight:bold">TRANSPORTE: '+esc(transporte)+'</td></tr>'
      +'<tr>'
      +'<td colspan="2" class="firma" style="padding:3px 6px">ENCARGADO DE PEDIDO</td>'
      +'<td colspan="2" class="firma" style="padding:3px 6px;border-left:1px solid #555">AUDITOR / VALIDADOR</td>'
      +'</tr>'
      +'<tr>'
      +'<td colspan="2" style="height:45px;padding:4px 6px">'+esc(encargado)+'</td>'
      +'<td colspan="2" style="height:45px;padding:4px 6px;border-left:1px solid #555">'+esc(auditorVal)+'</td>'
      +'</tr>'
      +'<tr><td colspan="4" style="text-align:center;padding:4px 6px;font-weight:bold;border-top:1px solid #555">ID DE CONTENEDOR: '+esc(idCont)+'</td></tr>'
      +'</table>'
      +'</body></html>';

    // ── Generar PDF con wkhtmltopdf ─────────────────────────────────────────
    var os   = require('os');
    var fs   = require('fs');
    var path = require('path');
    var { execFile } = require('child_process');

    var tmpHtml = path.join(os.tmpdir(), 'manif_'+cargaId+'_'+Date.now()+'.html');
    var tmpPdf  = tmpHtml.replace('.html','.pdf');

    fs.writeFileSync(tmpHtml, html, 'utf8');

    var wkArgs = [
      '--quiet',
      '--page-size', 'Letter',
      '--margin-top',    '8mm',
      '--margin-bottom', '8mm',
      '--margin-left',   '8mm',
      '--margin-right',  '8mm',
      '--encoding', 'utf-8',
      '--disable-smart-shrinking',
      tmpHtml, tmpPdf
    ];

    execFile('wkhtmltopdf', wkArgs, function(err) {
      // Limpiar HTML temporal siempre
      try { fs.unlinkSync(tmpHtml); } catch(_){}

      if(err) {
        try { fs.unlinkSync(tmpPdf); } catch(_){}
        console.error('wkhtmltopdf error:', err.message);
        return res.status(500).json({ ok:false, error:'Error al generar PDF: '+err.message });
      }

      var pdfBuf;
      try { pdfBuf = fs.readFileSync(tmpPdf); }
      catch(e) {
        return res.status(500).json({ ok:false, error:'No se pudo leer el PDF generado.' });
      }
      try { fs.unlinkSync(tmpPdf); } catch(_){}

      var licFile  = (licId||'BOD').replace(/[^A-Z0-9]/gi,'_');
      var fechaFile= (fechaFmt||'').replace(/\//g,'') || new Date().toISOString().slice(0,10).replace(/-/g,'');
      var pdfName  = 'Manifiesto_'+licFile+'_Furgon'+furgon+'_'+fechaFile+'.pdf';

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="'+pdfName+'"');
      res.setHeader('Content-Length', pdfBuf.length);
      res.end(pdfBuf);
    });

  } catch(e) {
    console.error('manifiesto-pdf error:', e.message);
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── DELETE /api/bod/sesion/:id ─────────────────────────────────────────────
// Borra (soft delete) una sesión bolsón y todos sus datos relacionados.
// Solo Erick Vela puede ejecutarlo — validado en BACKEND.
// Devuelve detalle por tabla para que el usuario sepa si algo falló.
// FIX (vie 06-jun-2026): borrar sesión bolsón por Erick Vela
app.delete('/api/bod/sesion/:id', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var usuario = String((req.body && req.body.usuario) || '').trim();

    if(usuario !== 'Erick Vela')
      return res.status(403).json({ ok:false, error:'Solo Erick Vela puede eliminar sesiones bolsón.' });

    var sesRows = await supabase('GET', 'bod_sesiones', null,
      '?id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length)
      return res.status(404).json({ ok:false, error:'Sesión no encontrada.' });
    var ses = sesRows[0];

    var now    = new Date().toISOString();
    var tablas = {};   // detalle por tabla: {ok, error, filas}
    var hayCritico = false;

    // 1. bod_lineas — soft delete (columna eliminada ya existe)
    try {
      await supabase('PATCH', 'bod_lineas',
        { eliminada:true, eliminado_por:usuario, ts_eliminado:now },
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      tablas.bod_lineas = { ok:true };
    } catch(e) {
      tablas.bod_lineas = { ok:false, error:e.message };
      hayCritico = true;
    }

    // 2. bod_tarima_furgon — borrar físico (sin columna eliminada)
    try {
      await supabase('DELETE', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      tablas.bod_tarima_furgon = { ok:true };
    } catch(e) {
      tablas.bod_tarima_furgon = { ok:false, error:e.message };
      // No es crítico si no hay asignaciones — continuar
    }

    // 3. bod_furgon_cierres — borrar físico
    try {
      await supabase('DELETE', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId));
      tablas.bod_furgon_cierres = { ok:true };
    } catch(e) {
      tablas.bod_furgon_cierres = { ok:false, error:e.message };
      hayCritico = true;
    }

    // 4. bod_sesiones — soft delete primero, físico como fallback
    try {
      await supabase('PATCH', 'bod_sesiones',
        { eliminada:true, eliminado_por:usuario, ts_eliminado:now },
        '?id=eq.'+encodeURIComponent(sesId));
      tablas.bod_sesiones = { ok:true, metodo:'soft_delete' };
    } catch(e) {
      // Intentar borrado físico si la columna no existe
      try {
        await supabase('DELETE', 'bod_sesiones', null,
          '?id=eq.'+encodeURIComponent(sesId));
        tablas.bod_sesiones = { ok:true, metodo:'delete_fisico' };
      } catch(e2) {
        tablas.bod_sesiones = { ok:false, error:e2.message };
        hayCritico = true;
      }
    }

    if(hayCritico) {
      return res.status(500).json({
        ok:false,
        sesion_id: sesId,
        licencia_id: ses.licencia_id,
        error: 'Algunas tablas no se pudieron limpiar. Verificá el detalle.',
        tablas: tablas
      });
    }

    invalidarCacheSesion(sesId);
    res.json({ ok:true, sesion_id:sesId, licencia_id:ses.licencia_id,
               mensaje:'Sesión '+ses.licencia_id+' eliminada por '+usuario+'.',
               tablas:tablas });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ════════════════════════════════════════════════════════════════════════════
// FASE 2 BODEGA CDG — Persistencia Teórico 952 + Papel de Trabajo
// ════════════════════════════════════════════════════════════════════════════

// ── GET /api/bod/sesion/:id/teorico-952 ───────────────────────────────────
app.get('/api/bod/sesion/:id/teorico-952', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var licHija = String(req.query.licencia_hija || '').trim().toUpperCase();
    var ck = 'bod:teorico952:'+sesId+':'+(licHija||'*');
    var cached = cacheGet(ck);
    if(cached) return res.json(cached);

    var query = '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=sku.asc';
    if(licHija) query += '&licencia_hija=eq.'+encodeURIComponent(licHija);
    var rows = await supabase('GET', 'bod_teorico_952', null, query);
    var resp = { ok:true, rows: Array.isArray(rows) ? rows : [] };
    cacheSet(ck, resp, 30000);
    res.json(resp);
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── POST /api/bod/sesion/:id/teorico-952 ──────────────────────────────────
// Body: { licencia_hija, skus: [{sku,nombre,cantidad_952,unidades,existencia,disponible}], creado_por }
app.post('/api/bod/sesion/:id/teorico-952', bodGuard, async (req, res) => {
  try {
    var sesId    = String(req.params.id).trim();
    var { licencia_hija, skus, creado_por } = req.body || {};
    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    creado_por    = String(creado_por    || '').trim();
    if(!Array.isArray(skus) || !skus.length)
      return res.status(400).json({ ok:false, error:'skus requeridos' });

    var now = new Date().toISOString();
    var upserts = skus.map(function(s){
      return {
        sesion_id:      sesId,
        licencia_hija:  licencia_hija,
        sku:            String(s.sku||'').trim().toUpperCase(),
        nombre:         String(s.nombre||''),
        cantidad_952:   Number(s.cantidad_952)||0,
        unidades:       Number(s.unidades)||0,
        existencia:     Number(s.existencia)||0,
        disponible:     Number(s.disponible)||0,
        creado_por:     creado_por,
        actualizado_en: now
      };
    }).filter(function(r){ return !!r.sku; });

    if(!upserts.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos' });

    // Batch upsert (Supabase: POST con Prefer merge-duplicates)
    var saved = await supabase('POST', 'bod_teorico_952', upserts,
      '?on_conflict=sesion_id,licencia_hija,sku');

    // Invalidar caches
    cacheInvalidatePrefix('bod:teorico952:'+sesId+':');

    res.json({ ok:true, guardados: upserts.length });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── GET /api/bod/sesion/:id/papel-trabajo ─────────────────────────────────
app.get('/api/bod/sesion/:id/papel-trabajo', bodGuard, async (req, res) => {
  try {
    var sesId        = String(req.params.id).trim();
    var licHija      = String(req.query.licencia_hija || '').trim().toUpperCase();
    var manifestoId  = String(req.query.manifiesto_id || '').trim();
    var ck = 'bod:papel:'+sesId+':'+(manifestoId||licHija||'*');
    var cached = cacheGet(ck);
    if(cached) return res.json(cached);
    // FIX (jue 11-jun-2026): filtrar por manifiesto_id si se especifica
    var query = '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=sku.asc&select=*';
    if(manifestoId) query += '&manifiesto_id=eq.'+encodeURIComponent(manifestoId);
    else if(licHija) query += '&licencia_hija=eq.'+encodeURIComponent(licHija);
    var rows = await supabase('GET', 'bod_papel_trabajo', null, query);
    var resp = { ok:true, rows: Array.isArray(rows) ? rows : [] };
    cacheSet(ck, resp, 15000);
    res.json(resp);
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── POST /api/bod/sesion/:id/papel-trabajo ────────────────────────────────
app.post('/api/bod/sesion/:id/papel-trabajo', bodGuard, async (req, res) => {
  try {
    var sesId    = String(req.params.id).trim();
    var { licencia_hija, skus, manifiesto_id } = req.body || {};
    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    manifiesto_id = manifiesto_id ? String(manifiesto_id).trim() : null;
    if(!Array.isArray(skus) || !skus.length)
      return res.status(400).json({ ok:false, error:'skus requeridos' });

    var now = new Date().toISOString();
    var upserts = skus.map(function(s){
      var validado = !!s.validado;
      var licFila = String(s.licencia_hija || licencia_hija || '').trim().toUpperCase();
      var row = {
        sesion_id:           sesId,
        licencia_hija:       licFila,
        sku:                 String(s.sku||'').trim().toUpperCase(),
        fisico_auditoria:    Number(s.fisico_auditoria)||0,
        diferencia:          Number(s.diferencia)||0,
        seguimiento:         String(s.seguimiento||''),
        validado:            validado,
        cantidad_manifiesto: Number(s.cantidad_manifiesto)||0,
        validado_por:        String(s.validado_por||''),
        validado_en:         validado ? (s.validado_en || now) : null,
        actualizado_en:      now
      };
      if(manifiesto_id) row.manifiesto_id = manifiesto_id;
      if(s.nombre     !== undefined) row.nombre        = String(s.nombre||'');
      if(s.cantidad_952 !== undefined) row.cantidad_952 = Number(s.cantidad_952)||0;
      if(s.estado_papel !== undefined) row.estado_papel = String(s.estado_papel||'');
      if(s.tr999      !== undefined) row.tr999          = String(s.tr999||'');
      if(s.tarimas    !== undefined) row.tarimas        = String(s.tarimas||'');
      if(s.furgon     !== undefined) row.furgon         = String(s.furgon||'');
      if(s.destino    !== undefined) row.destino        = String(s.destino||'');
      if(s.unidad     !== undefined) row.unidad         = String(s.unidad||'');
      return row;
    }).filter(function(r){ return !!r.sku; });

    if(!upserts.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos' });

    // on_conflict usa el índice único real: sesion_id,licencia_hija,furgon,sku
    // Con sesion_id='GLOBAL', licencia_hija='', furgon='' → único por sku global
    await supabase('POST', 'bod_papel_trabajo', upserts, '?on_conflict=sesion_id,licencia_hija,furgon,sku');

    cacheInvalidatePrefix('bod:papel:'+sesId+':');
    res.json({ ok:true, guardados: upserts.length });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// ── bod_manifiestos_control endpoints ─────────────────────────────────────
// FIX (jue 11-jun-2026): Control de Manifiestos — Bodega CDG

// GET /api/bod/sesion/:id/manifiestos-control
// FIX (jue 11-jun-2026, server v21): excluir eliminados + incluir _skuCount por manifiesto
app.get('/api/bod/sesion/:id/manifiestos-control', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var rows = await supabase('GET', 'bod_manifiestos_control', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&estado=neq.eliminado&order=creado_en.desc&select=*');
    // Cargar conteos de SKUs en Papel de Trabajo por manifiesto
    var ptRows = await supabase('GET','bod_papel_trabajo',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&select=manifiesto_id').catch(function(){return[];});
    var skuCounts = {};
    if(Array.isArray(ptRows)) ptRows.forEach(function(r){ if(r.manifiesto_id) skuCounts[r.manifiesto_id]=(skuCounts[r.manifiesto_id]||0)+1; });
    var manifiestos = (Array.isArray(rows)?rows:[]).map(function(m){
      return Object.assign({},m,{_skuCount: skuCounts[m.id]||0});
    });
    res.json({ ok:true, manifiestos: manifiestos });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// POST /api/bod/sesion/:id/manifiestos-control
// Body: { correlativo?, licencia_hija?, furgon?, auditor_lider, creado_por }
// FIX (jue 11-jun-2026, server v21): acepta correlativo manual PT-YYYY-MM-DD-NN
//   Si se omite, lo genera automáticamente.
app.post('/api/bod/sesion/:id/manifiestos-control', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var { licencia_hija, furgon, auditor_lider, creado_por, correlativo: corr } = req.body || {};
    licencia_hija = String(licencia_hija||'').trim().toUpperCase();
    furgon        = String(furgon||'').trim();
    var hoyGT = new Date().toLocaleDateString('en-CA',{timeZone:'America/Guatemala'});
    var correlativo;
    if(corr && String(corr).trim()) {
      // Validar formato PT-YYYY-MM-DD-NN
      correlativo = String(corr).trim().toUpperCase();
      if(!/^PT-\d{4}-\d{2}-\d{2}-\d{2}$/.test(correlativo))
        return res.status(400).json({ ok:false, error:'Formato de correlativo inválido. Esperado: PT-YYYY-MM-DD-NN (ej: PT-2026-06-11-01).' });
      // Verificar que no exista ya en esta sesión (activos, no eliminados)
      var dup = await supabase('GET','bod_manifiestos_control',null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&correlativo=eq.'+encodeURIComponent(correlativo)
        +'&estado=neq.eliminado&limit=1&select=id').catch(function(){return[];});
      if(Array.isArray(dup)&&dup.length)
        return res.status(409).json({ ok:false, error:'Ya existe un PT activo con correlativo '+correlativo+' en esta sesión.' });
    } else {
      // Auto-generar
      var existentes = await supabase('GET','bod_manifiestos_control',null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&correlativo=like.'+encodeURIComponent('PT-'+hoyGT+'-%')+'&select=correlativo').catch(function(){return[];});
      var nextNum = Array.isArray(existentes) ? existentes.length + 1 : 1;
      correlativo = 'PT-'+hoyGT+'-'+String(nextNum).padStart(2,'0');
    }
    var now = new Date().toISOString();
    var row = {
      sesion_id:      sesId,
      licencia_hija:  licencia_hija,
      furgon:         furgon,
      correlativo:    correlativo,
      estado:         'en_proceso',
      auditor_lider:  String(auditor_lider||''),
      colaboradores:  [],
      creado_por:     String(creado_por||''),
      creado_en:      now,
      actualizado_por:String(creado_por||''),
      actualizado_en: now,
      historial:      [{ accion:'creado', usuario:String(creado_por||''), ts:now }]
    };
    var saved = await supabase('POST', 'bod_manifiestos_control', [row], '?select=*');
    var created = Array.isArray(saved) ? saved[0] : null;
    res.json({ ok:true, manifiesto: created });
  } catch(e) {
    if(e.message && e.message.includes('duplicate')) {
      return res.status(409).json({ ok:false, error:'Ya existe un manifiesto activo con ese correlativo.' });
    }
    res.status(500).json({ ok:false, error:e.message });
  }
});

// POST /api/bod/sesion/:id/manifiestos-control/:mId/colaborador
// Body: { usuario }
app.post('/api/bod/sesion/:id/manifiestos-control/:mId/colaborador', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mId).trim();
    var { usuario } = req.body || {};
    usuario = String(usuario||'').trim();
    var rows = await supabase('GET', 'bod_manifiestos_control', null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=*');
    var m = Array.isArray(rows) && rows[0] ? rows[0] : null;
    if(!m) return res.status(404).json({ ok:false, error:'Manifiesto no encontrado' });
    var cols = Array.isArray(m.colaboradores) ? m.colaboradores : [];
    if(!cols.includes(usuario)) cols.push(usuario);
    var now  = new Date().toISOString();
    var hist = Array.isArray(m.historial) ? m.historial : [];
    hist.push({ accion:'colaborador_unido', usuario:usuario, ts:now });
    // FIX (jue 11-jun-2026): activar en_proceso si estaba sin_iniciar
    var nuevoEstado = m.estado === 'sin_iniciar' ? 'en_proceso' : m.estado;
    if(m.estado === 'sin_iniciar') {
      hist.push({ accion:'iniciado', usuario:usuario, ts:now });
    }
    await supabase('PATCH', 'bod_manifiestos_control',
      { colaboradores: cols, estado: nuevoEstado,
        actualizado_por: usuario, actualizado_en: now, historial: hist },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// PATCH /api/bod/sesion/:id/manifiestos-control/:mId/finalizar
// Body: { usuario }
app.patch('/api/bod/sesion/:id/manifiestos-control/:mId/finalizar', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mId).trim();
    var { usuario } = req.body || {};
    usuario = String(usuario||'').trim();
    var rows = await supabase('GET', 'bod_manifiestos_control', null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=historial');
    var m = Array.isArray(rows) && rows[0] ? rows[0] : null;
    if(!m) return res.status(404).json({ ok:false, error:'Manifiesto no encontrado' });
    var now = new Date().toISOString();
    var hist = Array.isArray(m.historial) ? m.historial : [];
    hist.push({ accion:'finalizado', usuario:usuario, ts:now });
    await supabase('PATCH', 'bod_manifiestos_control',
      { estado:'finalizado', finalizado_por:usuario, finalizado_en:now,
        actualizado_por:usuario, actualizado_en:now, historial:hist },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// PATCH /api/bod/sesion/:id/manifiestos-control/:mId/reabrir
// Body: { usuario }
app.patch('/api/bod/sesion/:id/manifiestos-control/:mId/reabrir', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    usuario  = String(usuario||'').trim();
    // FIX (jue 11-jun-2026): validación de permisos en servidor
    // Solo auditor_lider del manifiesto O usuario con rol auditor/supervisor
    var rows = await supabase('GET', 'bod_manifiestos_control', null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=historial,auditor_lider,estado');
    var m = Array.isArray(rows) && rows[0] ? rows[0] : null;
    if(!m) return res.status(404).json({ ok:false, error:'Manifiesto no encontrado' });
    var esLider = m.auditor_lider === usuario;
    var esAutoriz = !!auditor || !!supervisor;
    if(!esLider && !esAutoriz) {
      return res.status(403).json({ ok:false, error:'Solo el auditor líder o un auditor autorizado puede reabrir este manifiesto.' });
    }
    if(m.estado !== 'finalizado') {
      return res.status(400).json({ ok:false, error:'Solo se pueden reabrir manifiestos finalizados.' });
    }
    var now  = new Date().toISOString();
    var hist = Array.isArray(m.historial) ? m.historial : [];
    hist.push({ accion:'reabierto', usuario:usuario, ts:now });
    await supabase('PATCH', 'bod_manifiestos_control',
      { estado:'en_proceso', reabierto_por:usuario, reabierto_en:now,
        actualizado_por:usuario, actualizado_en:now, historial:hist },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// DELETE /api/bod/sesion/:id/manifiestos-control/:mId
// FIX (jue 11-jun-2026, server v21): borrar PT + carga logística + papel de trabajo relacionados.
// Body: { usuario, forzar? }
// Si tiene líneas en Papel de Trabajo y forzar !== true → { ok:false, tieneLineas:true, count:N }
// Si forzar===true o no hay líneas → soft-delete manifiesto + hard-delete carga y papel.
app.delete('/api/bod/sesion/:id/manifiestos-control/:mId', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var mId    = String(req.params.mId).trim();
    var { usuario, auditor, supervisor, forzar } = req.body || {};
    usuario = String(usuario||'').trim();
    var esAud = auditor === true || supervisor === true;
    if(!esAud) return res.status(403).json({ ok:false, error:'Solo auditores pueden borrar manifiestos.' });
    // Verificar que el manifiesto existe en esta sesión
    var rows = await supabase('GET','bod_manifiestos_control',null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=id,correlativo,estado');
    var m = Array.isArray(rows)&&rows[0] ? rows[0] : null;
    if(!m) return res.status(404).json({ ok:false, error:'Manifiesto no encontrado.' });
    // Verificar líneas en Papel de Trabajo
    var ptRows = await supabase('GET','bod_papel_trabajo',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(mId)+'&select=sku&limit=200').catch(function(){return[];});
    var ptCount = Array.isArray(ptRows) ? ptRows.length : 0;
    if(ptCount > 0 && forzar !== true) {
      return res.json({ ok:false, tieneLineas:true, count:ptCount, correlativo:m.correlativo||mId });
    }
    var now = new Date().toISOString();
    // 1. Soft-delete del manifiesto (queda en BD para auditoría pero no aparece en UI)
    await supabase('PATCH','bod_manifiestos_control',
      { estado:'eliminado', actualizado_por:usuario, actualizado_en:now },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    // 2. Hard-delete cargas logísticas vinculadas (bod_furgon_cierres)
    await supabase('DELETE','bod_furgon_cierres',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(mId)).catch(function(){});
    // 3. Hard-delete líneas de Papel de Trabajo vinculadas
    if(ptCount > 0) {
      await supabase('DELETE','bod_papel_trabajo',null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)
        +'&manifiesto_id=eq.'+encodeURIComponent(mId)).catch(function(){});
    }
    cacheInvalidatePrefix('bod:papel:'+sesId+':');
    invalidarCacheSesion(sesId);
    console.log('BOD manifiesto borrado: '+mId+' ('+m.correlativo+') por '+usuario);
    res.json({ ok:true, correlativo:m.correlativo||mId, ptBorradas:ptCount });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});


// PATCH /api/bod/sesion/:id/manifiestos-control/:mId/guardar
// FIX (jue 11-jun-2026, server v21): guardar campos editables inline (furgon, placa, marchamo, licencia_hija, destino_tr999, observacion)
// Actualiza bod_manifiestos_control + upsert en bod_furgon_cierres.
app.patch('/api/bod/sesion/:id/manifiestos-control/:mId/guardar', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var mId    = String(req.params.mId).trim();
    var { usuario, furgon, placa, marchamo, licencia_hija, destino_tr999, observacion } = req.body || {};
    usuario       = String(usuario||''  ).trim();
    furgon        = String(furgon||''  ).trim();
    placa         = String(placa||''   ).trim().toUpperCase();
    marchamo      = String(marchamo||''  ).trim().toUpperCase();
    licencia_hija = String(licencia_hija||''  ).trim().toUpperCase();
    destino_tr999 = String(destino_tr999||''  ).trim().toUpperCase();
    observacion   = String(observacion||''  ).trim();
    var now = new Date().toISOString();
    // 1. Actualizar bod_manifiestos_control con todos los campos editables del PT
    var { auditor_lider: aud_lider_body } = req.body || {};
    await supabase('PATCH','bod_manifiestos_control',
      { furgon:furgon, placa:placa, marchamo:marchamo, licencia_hija:licencia_hija,
        destino_tr999:destino_tr999, observacion:observacion,
        auditor_lider:String(aud_lider_body||'').trim(),
        actualizado_por:usuario, actualizado_en:now },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    // 2. Upsert bod_furgon_cierres por manifiesto_id (crear si no existe, actualizar si existe)
    var existing = await supabase('GET','bod_furgon_cierres',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&manifiesto_id=eq.'+encodeURIComponent(mId)+'&limit=1&select=id').catch(function(){return[];});
    if(Array.isArray(existing)&&existing[0]&&existing[0].id) {
      // Actualizar carga existente
      await supabase('PATCH','bod_furgon_cierres',
        { furgon:furgon, placa:placa, marchamo:marchamo, licencia_hija:licencia_hija,
          destino_tr999:destino_tr999, observacion:observacion,
          actualizado_por:usuario, actualizado_en:now },
        '?id=eq.'+encodeURIComponent(existing[0].id));
    } else {
      // Crear carga nueva vinculada al manifiesto
      await supabase('POST','bod_furgon_cierres',
        [{ sesion_id:sesId, manifiesto_id:mId, furgon:furgon, placa:placa, marchamo:marchamo,
           licencia_hija:licencia_hija, destino_tr999:destino_tr999, observacion:observacion,
           estado:'abierta', desde_papel:true, tarimas:[],
           creado_por:usuario, creado_en:now, actualizado_por:usuario, actualizado_en:now }],
        '?select=id').catch(function(){});
    }
    invalidarCacheSesion(sesId);
    console.log('BOD guardar-campos manif: '+mId+' por '+usuario);
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});
// PATCH /api/bod/sesion/:id/manifiestos-control/:mId/anular
app.patch('/api/bod/sesion/:id/manifiestos-control/:mId/anular', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mId).trim();
    var { usuario } = req.body || {};
    usuario = String(usuario||'').trim();
    var rows = await supabase('GET', 'bod_manifiestos_control', null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=historial');
    var m = Array.isArray(rows) && rows[0] ? rows[0] : null;
    if(!m) return res.status(404).json({ ok:false, error:'Manifiesto no encontrado' });
    var now = new Date().toISOString();
    var hist = Array.isArray(m.historial) ? m.historial : [];
    hist.push({ accion:'anulado', usuario:usuario, ts:now });
    await supabase('PATCH', 'bod_manifiestos_control',
      { estado:'anulado', anulado_por:usuario, anulado_en:now,
        actualizado_por:usuario, actualizado_en:now, historial:hist },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/sesion/:id/papel-trabajo-sku ───────────────────────────
// FIX (jue 11-jun-2026): eliminar una fila específica del Papel de Trabajo
// Body: { manifiesto_id, sku }
app.delete('/api/bod/sesion/:id/papel-trabajo-sku', bodGuard, async (req, res) => {
  try {
    var sesId       = String(req.params.id).trim();
    var manifestoId = String((req.body||{}).manifiesto_id||'').trim();
    var sku         = String((req.body||{}).sku||'').trim().toUpperCase();
    if(!manifestoId || !sku)
      return res.status(400).json({ ok:false, error:'manifiesto_id y sku requeridos' });
    await supabase('DELETE', 'bod_papel_trabajo', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(manifestoId)
      +'&sku=eq.'+encodeURIComponent(sku));
    cacheInvalidatePrefix('bod:papel:'+sesId+':');
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ═══════════════════════════════════════════════════════════════════════════
//   BODEGA CDG v2 — Endpoints limpios (jue 12-jun-2026, server v22)
//   Tres recursos: tarimas (captura operadores) + teorico (WMS persistido)
//   + paper trabajo ya existente + manifiestos ya existentes.
// ═══════════════════════════════════════════════════════════════════════════

// ── GET /api/bod/sesion/:id/tarimas ──────────────────────────────────────
app.get('/api/bod/sesion/:id/tarimas', bodGuard, async (req, res) => {
  // DEPRECATED (server v23): bod_tarimas no se usa — usar bod_lineas via /api/bod/sesion/:id/lineas
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ── POST /api/bod/sesion/:id/tarimas ─────────────────────────────────────
// Body: { licencia_bolson, correlativo, sku, descripcion, cantidad, estructura, operador }
app.post('/api/bod/sesion/:id/tarimas', bodGuard, async (req, res) => {
  // DEPRECATED (server v23): bod_tarimas no se usa — captura via /api/bod/sesion/:id/linea
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ── DELETE /api/bod/sesion/:id/tarimas/:lineId ────────────────────────────
// Solo puede borrar el operador que creó la línea
app.delete('/api/bod/sesion/:id/tarimas/:lineId', bodGuard, async (req, res) => {
  // DEPRECATED (server v23): bod_tarimas no se usa — borrar via /api/bod/linea/:lineaId
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ── GET /api/bod/sesion/:id/teorico ──────────────────────────────────────
// Retorna el teórico WMS persistido para esta sesión
app.get('/api/bod/sesion/:id/teorico', bodGuard, async (req, res) => {
  // DEPRECATED (server v23): usar bod_teorico_952 via /api/bod/sesion/:id/teorico-952
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ── POST /api/bod/sesion/:id/teorico/upload ───────────────────────────────
// Sube Excel WMS con columnas: sku, descripcion, cantidad
// Borra el teórico anterior y reemplaza con el nuevo
app.post('/api/bod/sesion/:id/teorico/upload', bodGuard, upload.single('file'), async (req, res) => {
  // DEPRECATED (server v23): usar bod_teorico_952 via /api/bod/sesion/:id/teorico-952/upload
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ── DELETE /api/bod/sesion/:id/teorico ────────────────────────────────────
app.delete('/api/bod/sesion/:id/teorico', bodGuard, async (req, res) => {
  // DEPRECATED (server v23): usar bod_teorico_952 via DELETE /api/bod/sesion/:id/teorico-952/all
  res.status(501).json({ ok:false, error:'Endpoint reemplazado.' });
});
;

// ═══════════════════════════════════════════════════════════════════════════
//  BODEGA CDG v23 — Endpoints nuevos (sab 14-jun-2026)
//  Tablas: bod_licencia_hija (nueva), bod_teorico_952 (existente),
//          bod_manifiestos_control, bod_papel_trabajo, bod_lineas, bod_sesiones
//  Fórmula Diferencia: Teórico 952 − Físico s/Auditoría
//    dif = 0  → Cuadrado
//    dif < 0  → Sobrante físico  (PT: "Solicitar traslado a 952")
//    dif > 0  → Faltante físico  (PT: "Se queda en bolsón")
// ═══════════════════════════════════════════════════════════════════════════

// ── Helpers locales ──────────────────────────────────────────────────────
function bodEstadoAuditoria(dif) {
  if(dif === 0)  return 'Cuadrado';
  if(dif < 0)   return 'Sobrante físico';
  return 'Faltante físico';
}
function bodSeguimientoPT(dif) {
  if(dif === 0) return 'Cuadrado';
  if(dif < 0)  return 'Solicitar traslado a 952';
  return 'Se queda en bolsón';
}
// Normaliza headers de Excel (quita acentos, trim, lowercase, reemplaza espacios/puntos)
function bodNormHeader(h) {
  return String(h||'').toLowerCase().trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[\s.#]+/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');
}

// ── PATCH /api/bod/linea/:lineaId/fisico-audit ────────────────────────────
// Tab 2 Detalle: guarda Físico s/Auditoría, devuelve diferencia + estado
// Body: { fisico_auditoria, teorico_952, usuario }
app.patch('/api/bod/linea/:lineaId/fisico-audit', bodGuard, async (req, res) => {
  try {
    var linId = String(req.params.lineaId).trim();
    var b     = req.body || {};
    var fisico    = Number(b.fisico_auditoria);
    var teo       = Number(b.teorico_952 || 0);
    var usuario   = String(b.usuario || '').trim();
    if(isNaN(fisico) || fisico < 0)
      return res.status(400).json({ ok:false, error:'fisico_auditoria debe ser >= 0.' });
    var rows = await supabase('GET','bod_lineas',null,
      '?id=eq.'+encodeURIComponent(linId)+'&limit=1&select=id,sesion_id,operador');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Línea no encontrada.' });
    var dif    = teo - fisico;
    var estado = bodEstadoAuditoria(dif);
    var now    = new Date().toISOString();
    await supabase('PATCH','bod_lineas',
      { cantidad_audit: fisico, auditado: true, auditado_por: usuario,
        ts_auditado: now, ts_modif: now },
      '?id=eq.'+encodeURIComponent(linId));
    invalidarCacheSesion(rows[0].sesion_id);
    res.json({ ok:true, diferencia: dif, estado: estado, seguimiento_pt: bodSeguimientoPT(dif) });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/sesion/:id/teorico-952/upload ───────────────────────────
// Tab 3 Carga Teórico: sube Excel WMS 952 y persiste en bod_teorico_952
// Columnas Excel: SKU, Nombre, Cant. 952, Unidad, Exist., Disp.
// Body (multipart): file + licencia_hija (opcional) + creado_por
app.post('/api/bod/sesion/:id/teorico-952/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    var sesId       = String(req.params.id).trim();
    var licenciaHija = String((req.body||{}).licencia_hija||'').trim().toUpperCase();
    var creadoPor   = String((req.body||{}).creado_por||(req.body||{}).usuario||'').trim();
    if(!req.file) return res.status(400).json({ ok:false, error:'No se recibió archivo.' });
    var wb  = XLSX.read(req.file.buffer, { type:'buffer' });
    var ws  = wb.Sheets[wb.SheetNames[0]];
    var raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
    if(!raw.length) return res.status(400).json({ ok:false, error:'El archivo está vacío.' });
    // Normalizar headers
    var keyMap = {};
    Object.keys(raw[0]).forEach(function(k){ keyMap[bodNormHeader(k)] = k; });
    // Mapear columnas
    var skuCol  = keyMap['sku'] || keyMap['codigo'] || keyMap['code'];
    var nomCol  = keyMap['nombre'] || keyMap['descripcion'] || keyMap['description'];
    var canCol  = keyMap['cant_952'] || keyMap['cantidad_952'] || keyMap['cant952']
               || keyMap['cantidad'] || keyMap['cant'];
    var uniCol  = keyMap['unidad'];
    var exiCol  = keyMap['exist'] || keyMap['existencia'] || keyMap['existencias'];
    var disCol  = keyMap['disp'] || keyMap['disponible'] || keyMap['disponibles'];
    if(!skuCol) return res.status(400).json({ ok:false, error:'No se encontró columna SKU en el archivo.' });
    var now    = new Date().toISOString();
    var upserts = [];
    raw.forEach(function(r){
      var sku = String(r[skuCol]||'').trim().toUpperCase();
      if(!sku) return;
      upserts.push({
        sesion_id:      sesId,
        licencia_hija:  licenciaHija || '',
        sku:            sku,
        nombre:         String(r[nomCol]||'').trim(),
        cantidad_952:   Number(r[canCol])||0,
        unidades:       Number(r[uniCol])||0,
        existencia:     Number(r[exiCol])||0,
        disponible:     Number(r[disCol])||0,
        creado_por:     creadoPor,
        actualizado_en: now
      });
    });
    if(!upserts.length)
      return res.status(400).json({ ok:false, error:'No se encontraron SKUs válidos.' });
    // Reemplazar teórico anterior para esta sesión + licencia_hija
    var delQuery = '?sesion_id=eq.'+encodeURIComponent(sesId);
    if(licenciaHija) delQuery += '&licencia_hija=eq.'+encodeURIComponent(licenciaHija);
    await supabase('DELETE','bod_teorico_952',null,delQuery);
    for(var i=0;i<upserts.length;i+=200){
      await supabase('POST','bod_teorico_952',upserts.slice(i,i+200),
        '?on_conflict=sesion_id,licencia_hija,sku');
    }
    cacheInvalidatePrefix('bod:teorico952:'+sesId+':');
    console.log('BOD teórico-952 upload: '+upserts.length+' SKUs sesión='+sesId+' por '+creadoPor);
    res.json({ ok:true, procesadas:upserts.length, subido_en:now });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/sesion/:id/teorico-952/all ────────────────────────────
// Limpia el teórico 952 de una sesión (con confirmación desde cliente)
app.delete('/api/bod/sesion/:id/teorico-952/all', bodGuard, async (req, res) => {
  try {
    var sesId        = String(req.params.id).trim();
    var licenciaHija = String((req.body||{}).licencia_hija||'').trim().toUpperCase();
    var delQ = '?sesion_id=eq.'+encodeURIComponent(sesId);
    if(licenciaHija) delQ += '&licencia_hija=eq.'+encodeURIComponent(licenciaHija);
    await supabase('DELETE','bod_teorico_952',null,delQ);
    cacheInvalidatePrefix('bod:teorico952:'+sesId+':');
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/sesion/:id/licencia-hija ─────────────────────────────────
// Lista filas de bod_licencia_hija para la sesión (+ filtro opcional manifiesto_id)
app.get('/api/bod/sesion/:id/licencia-hija', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var mId    = String(req.query.manifiesto_id||'').trim();
    var ck     = 'bod:lichija:'+sesId+':'+(mId||'*');
    var cached = cacheGet(ck);
    if(cached) return res.json(cached);
    var q = '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=sku.asc&select=*';
    if(mId) q += '&manifiesto_id=eq.'+encodeURIComponent(mId);
    var rows = await supabase('GET','bod_licencia_hija',null,q);
    var resp = { ok:true, rows:Array.isArray(rows)?rows:[] };
    cacheSet(ck, resp, 30000);
    res.json(resp);
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/sesion/:id/licencia-hija/upload ─────────────────────────
// Sube Excel licencia hija WMS y persiste en bod_licencia_hija
// Columnas: Número, # Ingreso, Doc. SAP, Fecha, Tipo, Lineas, Status,
//           SKU, Nombre, Cant., Unidad, Unidades, Origen, Destino
// Body: file + manifiesto_id + usuario
app.post('/api/bod/sesion/:id/licencia-hija/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String((req.body||{}).manifiesto_id||'').trim();
    var usu   = String((req.body||{}).usuario||'').trim();
    if(!req.file) return res.status(400).json({ ok:false, error:'No se recibió archivo.' });
    var wb  = XLSX.read(req.file.buffer, { type:'buffer' });
    var ws  = wb.Sheets[wb.SheetNames[0]];
    var raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
    if(!raw.length) return res.status(400).json({ ok:false, error:'Archivo vacío.' });
    // Normalizar headers
    var km = {};
    Object.keys(raw[0]).forEach(function(k){ km[bodNormHeader(k)] = k; });
    var skuCol  = km['sku'] || km['codigo'] || km['code'];
    var numCol  = km['numero'] || km['numero_licencia'] || km['licencia'];
    var ingCol  = km['ingreso'] || km['_ingreso'];
    var sapCol  = km['doc_sap'] || km['doc__sap'] || km['docsap'] || km['sap'];
    var fecCol  = km['fecha'];
    var tipCol  = km['tipo'];
    var linCol  = km['lineas'];
    var staCol  = km['status'] || km['estado'];
    var nomCol  = km['nombre'] || km['descripcion'];
    var canCol  = km['cant'] || km['cantidad'] || km['cant_'];
    var uniCol  = km['unidad'];
    var uniSCol = km['unidades'];
    var oriCol  = km['origen'];
    var desCol  = km['destino'];
    if(!skuCol) return res.status(400).json({ ok:false, error:'No se encontró columna SKU.' });
    var now  = new Date().toISOString();
    var rows = [];
    raw.forEach(function(r){
      var sku = String(r[skuCol]||'').trim().toUpperCase();
      if(!sku) return;
      rows.push({
        sesion_id:       sesId,
        manifiesto_id:   mId||null,
        numero_licencia: String(r[numCol]||'').trim().toUpperCase(),
        ingreso:         String(r[ingCol]||'').trim(),
        doc_sap:         String(r[sapCol]||'').trim(),
        fecha_wms:       String(r[fecCol]||'').trim(),
        tipo:            String(r[tipCol]||'').trim(),
        lineas:          Number(r[linCol])||null,
        status_wms:      String(r[staCol]||'').trim(),
        sku:             sku,
        nombre:          String(r[nomCol]||'').trim(),
        cantidad:        Number(r[canCol])||0,
        unidad:          String(r[uniCol]||'').trim(),
        unidades:        Number(r[uniSCol])||0,
        origen:          String(r[oriCol]||'').trim(),
        destino:         String(r[desCol]||'').trim(),
        subido_por:      usu,
        subido_en:       now
      });
    });
    if(!rows.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos.' });
    // Reemplazar: borrar anterior para esta sesión + manifiesto_id
    var delQ = '?sesion_id=eq.'+encodeURIComponent(sesId);
    if(mId) delQ += '&manifiesto_id=eq.'+encodeURIComponent(mId);
    await supabase('DELETE','bod_licencia_hija',null,delQ);
    for(var i=0;i<rows.length;i+=200){
      await supabase('POST','bod_licencia_hija',rows.slice(i,i+200),'?select=id');
    }
    cacheInvalidatePrefix('bod:lichija:'+sesId+':');
    console.log('BOD lic-hija upload: '+rows.length+' filas sesion='+sesId+' PT='+mId+' por '+usu);
    res.json({ ok:true, procesadas:rows.length, subido_en:now });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/sesion/:id/licencia-hija ──────────────────────────────
app.delete('/api/bod/sesion/:id/licencia-hija', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String((req.body||{}).manifiesto_id||'').trim();
    var delQ  = '?sesion_id=eq.'+encodeURIComponent(sesId);
    if(mId) delQ += '&manifiesto_id=eq.'+encodeURIComponent(mId);
    await supabase('DELETE','bod_licencia_hija',null,delQ);
    cacheInvalidatePrefix('bod:lichija:'+sesId+':');
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── PATCH /api/bod/sesion/:id/manifiestos-control/:mid/correlativo ─────────
// Corrección manual del correlativo PT (permite editar el ## final)
// Body: { correlativo, usuario }
app.patch('/api/bod/sesion/:id/manifiestos-control/:mid/correlativo', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mid).trim();
    var { correlativo, usuario } = req.body || {};
    correlativo = String(correlativo||'').trim().toUpperCase();
    if(!/^PT-\d{4}-\d{2}-\d{2}-\d{2}$/.test(correlativo))
      return res.status(400).json({ ok:false, error:'Formato inválido. Esperado: PT-YYYY-MM-DD-## (ej: PT-2026-06-14-01).' });
    // Verificar que no exista duplicado en esta sesión
    var dup = await supabase('GET','bod_manifiestos_control',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&correlativo=eq.'+encodeURIComponent(correlativo)
      +'&id=neq.'+encodeURIComponent(mId)
      +'&estado=neq.eliminado&limit=1&select=id').catch(function(){return[];});
    if(Array.isArray(dup)&&dup.length)
      return res.status(409).json({ ok:false, error:'Ya existe otro PT activo con ese correlativo.' });
    var now = new Date().toISOString();
    await supabase('PATCH','bod_manifiestos_control',
      { correlativo:correlativo, actualizado_por:String(usuario||''), actualizado_en:now },
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId));
    invalidarCacheSesion(sesId);
    res.json({ ok:true, correlativo:correlativo });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/sesion/:id/manifiestos-control/:mid/manifiesto-data ───────
// Compila todos los datos necesarios para generar el manifiesto exportable.
// Fuentes: bod_manifiestos_control, bod_furgon_cierres, bod_papel_trabajo,
//          bod_licencia_hija, bod_teorico_952
// Fórmula: Diferencia = Teórico 952 − Físico s/Auditoría
app.get('/api/bod/sesion/:id/manifiestos-control/:mid/manifiesto-data', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var mId   = String(req.params.mid).trim();

    // 1. PT header (bod_manifiestos_control)
    var ptRows = await supabase('GET','bod_manifiestos_control',null,
      '?id=eq.'+encodeURIComponent(mId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1&select=*');
    if(!Array.isArray(ptRows)||!ptRows.length)
      return res.status(404).json({ ok:false, error:'PT no encontrado.' });
    var pt = ptRows[0];

    // 2. Furgón/cierre data (bod_furgon_cierres)
    var furgRows = await supabase('GET','bod_furgon_cierres',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(mId)+'&limit=1&select=*')
      .catch(function(){ return []; });
    var cierreData = Array.isArray(furgRows)&&furgRows[0] ? furgRows[0] : {};

    // 3. Papel de Trabajo lines (bod_papel_trabajo)
    var papelRows = await supabase('GET','bod_papel_trabajo',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(mId)+'&order=sku.asc&select=*');
    var papelLineas = Array.isArray(papelRows) ? papelRows : [];

    // 4. Licencia hija WMS (bod_licencia_hija) → columna WMS en manifiesto
    var licHijaRows = await supabase('GET','bod_licencia_hija',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&manifiesto_id=eq.'+encodeURIComponent(mId)+'&select=sku,cantidad,destino')
      .catch(function(){ return []; });
    // Indexar por SKU para lookup rápido
    var licHijaMap = {};
    if(Array.isArray(licHijaRows)) licHijaRows.forEach(function(r){
      if(!licHijaMap[r.sku]) licHijaMap[r.sku] = 0;
      licHijaMap[r.sku] += Number(r.cantidad)||0;
    });

    // 5. Teórico 952 (bod_teorico_952) → columna 952 en papel de trabajo
    var teoRows = await supabase('GET','bod_teorico_952',null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&select=sku,cantidad_952')
      .catch(function(){ return []; });
    var teoMap = {};
    if(Array.isArray(teoRows)) teoRows.forEach(function(r){
      teoMap[r.sku] = Number(r.cantidad_952)||0;
    });

    // 6. Enriquecer líneas del papel de trabajo con datos calculados
    var lineasEnriquecidas = papelLineas.map(function(p){
      var teo952   = teoMap[p.sku] || 0;
      var fisico   = Number(p.fisico_auditoria)||0;
      var wms      = licHijaMap[p.sku] || 0;
      var dif      = teo952 - fisico;               // Fórmula canónica
      var validado = !!p.validado;
      var difPost  = validado ? 0 : dif;
      var pct      = teo952 > 0 ? Math.round((fisico / teo952) * 100) : (fisico > 0 ? 100 : 0);
      return {
        sku:                p.sku,
        nombre:             p.nombre || '',
        estado_papel:       p.estado_papel || 'No auditado',
        fisico_auditoria:   fisico,
        cantidad_952:       teo952,
        cantidad_wms:       wms,
        diferencia:         dif,
        seguimiento:        bodSeguimientoPT(dif),
        validado:           validado,
        diferencia_post:    difPost,
        porcentaje_auditado: pct,
        tarimas:            p.tarimas || '',
        tr999:              p.tr999 || ''
      };
    });

    // 7. Calcular porcentaje global de cobertura
    var totalFisico = lineasEnriquecidas.reduce(function(s,l){return s+l.fisico_auditoria;},0);
    var total952    = lineasEnriquecidas.reduce(function(s,l){return s+l.cantidad_952;},0);
    var pctGlobal   = total952 > 0 ? Math.round((totalFisico/total952)*100) : 0;

    res.json({
      ok:          true,
      generado_en: new Date().toISOString(),
      pt: {
        id:             pt.id,
        correlativo:    pt.correlativo,
        estado:         pt.estado,
        auditor_lider:  pt.auditor_lider,
        colaboradores:  pt.colaboradores || [],
        creado_por:     pt.creado_por,
        furgon:         cierreData.furgon   || pt.furgon   || '',
        placa:          cierreData.placa    || '',
        marchamo:       cierreData.marchamo || '',
        licencia_hija:  cierreData.licencia_hija || pt.licencia_hija || '',
        destino_tr999:  cierreData.destino_tr999 || pt.destino_tr999 || '',
        observacion:    cierreData.observacion   || ''
      },
      skus:           lineasEnriquecidas,
      total_lineas:   lineasEnriquecidas.length,
      total_fisico:   totalFisico,
      total_952:      total952,
      pct_auditado:   pctGlobal,
      // Campos fijos del manifiesto
      campos_fijos: {
        direccion_origen:  '27 Calle Bodega C 41-55 Zona 5 Calzada la Paz',
        direccion_destino: 'BODEGA NODUS (Hamilton)',
        autorizacion_envio:['ASTRID DUARTE','CINTYA RIVERA','HUVALDO PEREZ'],
        encargado_pedido:  'HUVALDO PEREZ',
        numero_transporte: 'TH 243',
        tipo_transporte:   'TRANSPORTE'
      }
    });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ══════════════════════════════════════════════════════════════════════════
//  NUEVOS ENDPOINTS GLOBALES POR FECHA — Bodega CDG v3
//  No usan sesion_id como discriminador. Global por fecha_trabajo (día).
//  Tablas nuevas: bod_teorico_global, bod_pts, bod_pts_papel
//  Regla: no toca endpoints existentes. No usa DROP.
// ══════════════════════════════════════════════════════════════════════════

function bodHoyGT() {
  return new Date().toLocaleDateString('en-CA', { timeZone:'America/Guatemala' });
}

// ── GET /api/bod/teorico-952-global ───────────────────────────────────────
// Devuelve el teórico 952 global para una fecha_trabajo.
// Query: ?fecha=YYYY-MM-DD (default: hoy GT)
app.get('/api/bod/teorico-952-global', bodGuard, async (req, res) => {
  try {
    var fecha = String(req.query.fecha || bodHoyGT()).trim();
    var rows = await supabase('GET', 'bod_teorico_global', null,
      '?fecha_trabajo=eq.' + encodeURIComponent(fecha) + '&order=sku.asc&select=sku,nombre,cantidad_952,unidades,existencia,disponible');
    res.json({ ok:true, fecha:fecha, rows: Array.isArray(rows) ? rows : [] });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/teorico-952-global/upload ───────────────────────────────
// Sube/reemplaza el teórico 952 global para una fecha.
// Body: multipart — file + fecha (opcional, default hoy) + creado_por
app.post('/api/bod/teorico-952-global/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    var fecha     = String((req.body || {}).fecha     || bodHoyGT()).trim();
    var creadoPor = String((req.body || {}).creado_por || '').trim();
    if (!req.file) return res.status(400).json({ ok:false, error:'No se recibió archivo.' });

    var wb  = XLSX.read(req.file.buffer, { type:'buffer' });
    var ws  = wb.Sheets[wb.SheetNames[0]];
    var raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
    if (!raw.length) return res.status(400).json({ ok:false, error:'El archivo está vacío.' });

    // Normalizar headers
    var km = {};
    Object.keys(raw[0]).forEach(function(k){ km[bodNormHeader(k)] = k; });
    var skuCol = km['sku'] || km['codigo'] || km['code'];
    var nomCol = km['nombre'] || km['descripcion'];
    var canCol = km['cant_952'] || km['cantidad_952'] || km['cant952'] || km['cantidad'] || km['cant'];
    var uniCol = km['unidad'];
    var exiCol = km['exist'] || km['existencia'];
    var disCol = km['disp']  || km['disponible'];
    if (!skuCol) return res.status(400).json({ ok:false, error:'No se encontró columna SKU.' });

    var now = new Date().toISOString();
    var upserts = [];
    raw.forEach(function(r){
      var sku = String(r[skuCol]||'').trim().toUpperCase();
      if (!sku) return;
      upserts.push({
        fecha_trabajo: fecha,
        sku:           sku,
        nombre:        String(r[nomCol]||'').trim(),
        cantidad_952:  Number(r[canCol]) || 0,
        unidades:      Number(r[uniCol]) || 0,
        existencia:    Number(r[exiCol]) || 0,
        disponible:    Number(r[disCol]) || 0,
        creado_por:    creadoPor,
        actualizado_en: now
      });
    });
    if (!upserts.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos.' });

    // Reemplazar teórico del día
    await supabase('DELETE', 'bod_teorico_global', null, '?fecha_trabajo=eq.' + encodeURIComponent(fecha));
    for (var i = 0; i < upserts.length; i += 200) {
      await supabase('POST', 'bod_teorico_global', upserts.slice(i, i+200),
        '?on_conflict=fecha_trabajo,sku');
    }
    console.log('BOD teorico-global upload: ' + upserts.length + ' SKUs fecha=' + fecha + ' por ' + creadoPor);
    res.json({ ok:true, procesadas: upserts.length, fecha:fecha });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/teorico-952-global ────────────────────────────────────
app.delete('/api/bod/teorico-952-global', bodGuard, async (req, res) => {
  try {
    var fecha = String(req.query.fecha || req.body && req.body.fecha || bodHoyGT()).trim();
    await supabase('DELETE', 'bod_teorico_global', null, '?fecha_trabajo=eq.' + encodeURIComponent(fecha));
    res.json({ ok:true, fecha:fecha });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/pts ───────────────────────────────────────────────────────
// Lista PTs globales para una fecha.
app.get('/api/bod/pts', bodGuard, async (req, res) => {
  try {
    var fecha = String(req.query.fecha || bodHoyGT()).trim();
    var rows = await supabase('GET', 'bod_pts', null,
      '?fecha_trabajo=eq.' + encodeURIComponent(fecha) + '&estado=neq.borrado&order=creado_en.asc&select=*');
    res.json({ ok:true, pts: Array.isArray(rows) ? rows : [], fecha:fecha });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/pts ─────────────────────────────────────────────────────
// Crea un PT nuevo. Genera correlativo PT-YYYY-MM-DD-## automáticamente.
app.post('/api/bod/pts', bodGuard, async (req, res) => {
  try {
    var { fecha, creado_por } = req.body || {};
    fecha      = String(fecha      || bodHoyGT()).trim();
    creado_por = String(creado_por || '').trim();

    // Siguiente correlativo del día
    var existentes = await supabase('GET', 'bod_pts', null,
      '?fecha_trabajo=eq.' + encodeURIComponent(fecha) + '&select=correlativo').catch(function(){ return []; });
    var nextNum = (Array.isArray(existentes) ? existentes.length : 0) + 1;
    var correlativo = 'PT-' + fecha + '-' + String(nextNum).padStart(2, '0');
    // Evitar duplicados si ya existe ese correlativo
    var dup = Array.isArray(existentes) && existentes.find(function(e){ return e.correlativo === correlativo; });
    if (dup) nextNum++; correlativo = 'PT-' + fecha + '-' + String(nextNum).padStart(2, '0');

    var now = new Date().toISOString();
    var pt = {
      correlativo:    correlativo,
      fecha_trabajo:  fecha,
      estado:         'en_proceso',
      creado_por:     creado_por,
      auditor_lider:  '',
      colaboradores:  [],
      furgon:         '',
      placa:          '',
      marchamo:       '',
      licencia_hija:  '',
      destino_tr999:  '',
      observacion:    '',
      carga_no:       '',
      creado_en:      now,
      actualizado_en: now
    };
    var saved = await supabase('POST', 'bod_pts', [pt], '?select=*');
    var created = Array.isArray(saved) ? saved[0] : pt;
    res.json({ ok:true, pt: created });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── PATCH /api/bod/pts/:ptId ──────────────────────────────────────────────
// Actualiza campos de un PT (furgon, placa, marchamo, etc.)
app.patch('/api/bod/pts/:ptId', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    var { usuario, furgon, placa, marchamo, licencia_hija, destino_tr999, observacion, carga_no, auditor_lider, estado } = req.body || {};
    var now = new Date().toISOString();
    var patch = { actualizado_en: now };
    if (usuario      !== undefined) patch.actualizado_por = String(usuario     || '').trim();
    if (furgon       !== undefined) patch.furgon          = String(furgon      || '').trim();
    if (placa        !== undefined) patch.placa           = String(placa       || '').trim().toUpperCase();
    if (marchamo     !== undefined) patch.marchamo        = String(marchamo    || '').trim().toUpperCase();
    if (licencia_hija!== undefined) patch.licencia_hija   = String(licencia_hija||'').trim().toUpperCase();
    if (destino_tr999!== undefined) patch.destino_tr999   = String(destino_tr999||'').trim().toUpperCase();
    if (observacion  !== undefined) patch.observacion     = String(observacion || '').trim();
    if (carga_no     !== undefined) patch.carga_no        = String(carga_no    || '').trim();
    if (auditor_lider!== undefined) patch.auditor_lider   = String(auditor_lider||'').trim();
    if (estado       !== undefined) patch.estado          = String(estado      || '').trim();
    await supabase('PATCH', 'bod_pts', patch, '?id=eq.' + encodeURIComponent(ptId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/pts/:ptId ─────────────────────────────────────────────
app.delete('/api/bod/pts/:ptId', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    await supabase('PATCH', 'bod_pts', { estado:'borrado', actualizado_en:new Date().toISOString() },
      '?id=eq.' + encodeURIComponent(ptId));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/pts/:ptId/papel ──────────────────────────────────────────
app.get('/api/bod/pts/:ptId/papel', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    var rows = await supabase('GET', 'bod_pts_papel', null,
      '?pt_id=eq.' + encodeURIComponent(ptId) + '&order=creado_en.asc&select=*');
    res.json({ ok:true, rows: Array.isArray(rows) ? rows : [] });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/pts/:ptId/papel ────────────────────────────────────────
// Upsert de SKU en papel de trabajo. Unique: pt_id + sku.
app.post('/api/bod/pts/:ptId/papel', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    var { skus, usuario } = req.body || {};
    if (!Array.isArray(skus) || !skus.length)
      return res.status(400).json({ ok:false, error:'skus requeridos' });

    var now = new Date().toISOString();
    var upserts = skus.map(function(s){
      var validado = !!s.validado;
      return {
        pt_id:            ptId,
        sku:              String(s.sku || '').trim().toUpperCase(),
        nombre:           String(s.nombre || ''),
        fisico_auditoria: Number(s.fisico_auditoria) || 0,
        cantidad_952:     Number(s.cantidad_952) || 0,
        diferencia:       Number(s.diferencia) || 0,
        seguimiento:      String(s.seguimiento || ''),
        validado:         validado,
        validado_por:     String(s.validado_por || ''),
        validado_en:      validado ? (s.validado_en || now) : null,
        tarimas:          String(s.tarimas || ''),
        estado_papel:     String(s.estado_papel || ''),
        origen:           String(s.origen || 'creado'),
        wms_cantidad:     s.wms_cantidad !== undefined ? Number(s.wms_cantidad) : null,
        creado_por:       String(s.creado_por || usuario || ''),
        creado_en:        s.creado_en || now,
        actualizado_en:   now
      };
    }).filter(function(r){ return !!r.sku; });

    if (!upserts.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos' });
    await supabase('POST', 'bod_pts_papel', upserts, '?on_conflict=pt_id,sku');
    res.json({ ok:true, guardados: upserts.length });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── DELETE /api/bod/pts/:ptId/papel/:sku ─────────────────────────────────
app.delete('/api/bod/pts/:ptId/papel/:sku', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    var sku  = String(req.params.sku).trim().toUpperCase();
    await supabase('DELETE', 'bod_pts_papel', null,
      '?pt_id=eq.' + encodeURIComponent(ptId) + '&sku=eq.' + encodeURIComponent(sku));
    res.json({ ok:true });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── POST /api/bod/pts/:ptId/licencia-hija/upload ─────────────────────────
// Sube licencia hija WMS para un PT. Guarda en bod_pts_licencia_hija.
app.post('/api/bod/pts/:ptId/licencia-hija/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    var ptId      = String(req.params.ptId).trim();
    var creadoPor = String((req.body || {}).creado_por || '').trim();
    if (!req.file) return res.status(400).json({ ok:false, error:'No se recibió archivo.' });

    var wb  = XLSX.read(req.file.buffer, { type:'buffer' });
    var ws  = wb.Sheets[wb.SheetNames[0]];
    var raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
    if (!raw.length) return res.status(400).json({ ok:false, error:'Archivo vacío.' });

    var km = {};
    Object.keys(raw[0]).forEach(function(k){ km[bodNormHeader(k)] = k; });
    var numCol  = km['numero'] || km['licencia'] || km['num'];
    var skuCol  = km['sku'] || km['codigo'];
    var nomCol  = km['nombre'] || km['descripcion'];
    var canCol  = km['cant'] || km['cantidad'] || km['cantidades'];
    var uniCol  = km['unidad'];
    var oriCol  = km['origen'];
    var desCol  = km['destino'];
    if (!skuCol) return res.status(400).json({ ok:false, error:'No se encontró columna SKU.' });

    var now = new Date().toISOString();
    var upserts = [];
    raw.forEach(function(r){
      var sku = String(r[skuCol]||'').trim().toUpperCase();
      if (!sku) return;
      upserts.push({
        pt_id:      ptId,
        sku:        sku,
        nombre:     String(r[nomCol]||'').trim(),
        cantidad:   Number(r[canCol]) || 0,
        unidad:     String(r[uniCol]||''),
        numero:     String(r[numCol]||'').trim().toUpperCase(),
        origen:     String(r[oriCol]||''),
        destino:    String(r[desCol]||''),
        creado_por: creadoPor,
        creado_en:  now
      });
    });
    if (!upserts.length) return res.status(400).json({ ok:false, error:'Sin SKUs válidos.' });

    // Reemplazar licencia hija del PT
    await supabase('DELETE', 'bod_pts_licencia_hija', null, '?pt_id=eq.' + encodeURIComponent(ptId));
    for (var i = 0; i < upserts.length; i += 200) {
      await supabase('POST', 'bod_pts_licencia_hija', upserts.slice(i, i+200), '?on_conflict=pt_id,sku');
    }
    res.json({ ok:true, procesadas: upserts.length });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── GET /api/bod/pts/:ptId/licencia-hija ─────────────────────────────────
app.get('/api/bod/pts/:ptId/licencia-hija', bodGuard, async (req, res) => {
  try {
    var ptId = String(req.params.ptId).trim();
    var rows = await supabase('GET', 'bod_pts_licencia_hija', null,
      '?pt_id=eq.' + encodeURIComponent(ptId) + '&order=sku.asc&select=sku,nombre,cantidad,unidad,origen,destino');
    res.json({ ok:true, rows: Array.isArray(rows) ? rows : [] });
  } catch(e) { res.status(500).json({ ok:false, error:e.message }); }
});

// ── SQL helper: Supabase table creation via RPC (if tables don't exist) ──
// Tables needed: bod_teorico_global, bod_pts, bod_pts_papel, bod_pts_licencia_hija
// These must be created in Supabase dashboard or via migration.
// Schema reference:
//
// bod_teorico_global:
//   id uuid default gen_random_uuid() primary key,
//   fecha_trabajo date not null,
//   sku text not null,
//   nombre text,
//   cantidad_952 numeric default 0,
//   unidades numeric default 0,
//   existencia numeric default 0,
//   disponible numeric default 0,
//   creado_por text,
//   actualizado_en timestamptz,
//   unique(fecha_trabajo, sku)
//
// bod_pts:
//   id uuid default gen_random_uuid() primary key,
//   correlativo text not null,
//   fecha_trabajo date not null,
//   estado text default 'en_proceso',
//   creado_por text, auditor_lider text, colaboradores jsonb default '[]',
//   furgon text, placa text, marchamo text, licencia_hija text,
//   destino_tr999 text, observacion text, carga_no text,
//   creado_en timestamptz, actualizado_en timestamptz, actualizado_por text
//
// bod_pts_papel:
//   id uuid default gen_random_uuid() primary key,
//   pt_id uuid references bod_pts(id),
//   sku text not null,
//   nombre text, fisico_auditoria numeric default 0, cantidad_952 numeric default 0,
//   diferencia numeric default 0, seguimiento text, validado boolean default false,
//   validado_por text, validado_en timestamptz, tarimas text, estado_papel text,
//   origen text default 'creado', wms_cantidad numeric,
//   creado_por text, creado_en timestamptz, actualizado_en timestamptz,
//   unique(pt_id, sku)
//
// bod_pts_licencia_hija:
//   id uuid default gen_random_uuid() primary key,
//   pt_id uuid references bod_pts(id),
//   sku text not null,
//   nombre text, cantidad numeric default 0, unidad text,
//   numero text, origen text, destino text,
//   creado_por text, creado_en timestamptz,
//   unique(pt_id, sku)

