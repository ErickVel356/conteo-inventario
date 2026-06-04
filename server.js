const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const https   = require('https');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

// LÃ­mite a 5MB es suficiente para conteos por campo, asignaciones, hallazgos,
// metadata, CDG. Antes era 50MB lo que abrÃ­a puerta a payloads gigantes que
// matan la RAM del free tier. Uploads de teorico/costos NO pasan por aquÃ­,
// usan multer (memoryStorage), que tiene su propio lÃ­mite.
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// â”€â”€ Supabase config (set via environment variables in Render) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const SUPABASE_URL = process.env.SUPABASE_URL;  // https://xxx.supabase.co
const SUPABASE_KEY = process.env.SUPABASE_KEY;  // anon/publishable key (RLS allow_all)

// â”€â”€ Simple Supabase REST client â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function supabase(method, table, body, query) {
  return new Promise((resolve, reject) => {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return resolve(null); // No Supabase configured â€” use memory only
    }
    const url  = new URL(`${SUPABASE_URL}/rest/v1/${table}${query||''}`);
    const data = body ? JSON.stringify(body) : null;
    // FIX BUG LATENTE (miÃ© 20-may-2026): agregar 'resolution=merge-duplicates'
    // al header Prefer. Sin esto, los POST con ?on_conflict=key fallan con
    // HTTP 409 'duplicate key' porque PostgREST no sabe que debe hacer UPSERT.
    // Antes el error se tragaba silenciosamente (resolve null) y el save NUNCA
    // persistÃ­a vÃ­a server.js â€” los datos llegaban a Supabase solo por
    // saveDirectToSupabase del cliente (que usa PATCH directo).
    // Con el fix v8 que ahora detecta errores HTTP, este 409 se hizo visible.
    // Esta lÃ­nea lo arregla de raÃ­z.
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
        // FIX CRÃTICO (miÃ© 20-may-2026, rev ChatGPT): rechazar la Promise
        // si Supabase responde con error HTTP (4xx/5xx). Antes resolvÃ­amos
        // SIEMPRE con null aunque el server hubiera rechazado el guardado,
        // lo que dejaba a saveDailyState y dbSet creyendo que todo OK.
        // Esto era la causa raÃ­z REAL del bug que vimos en producciÃ³n.
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

// FIX CRÃTICO (miÃ© 20-may-2026, rev Claude2): timeout para evitar que un
// Supabase lento cuelgue al cliente hasta el lÃ­mite de Render/Cloudflare
// (~100s). Si dbSet tarda mÃ¡s de 15s, rechazamos para que el endpoint
// responda con error custom al cliente. El dbSet sigue en vuelo en el
// background â€” si eventualmente completa, mejor (UPSERT idempotente).
function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(
      () => reject(new Error(label + ' timeout despuÃ©s de ' + ms + 'ms')),
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

// â”€â”€ In-memory state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
// FIX (lun 1-jun-2026, v19): versiÃ³n estricta que SÃ rechaza si dbSet falla.
// saveDailyState() tiene .catch interno â€” withTimeout() recibe una promesa
// resuelta aunque Supabase devuelva 500, y el endpoint responderÃ­a ok:true
// sin haber persistido. saveDailyStateStrict() propaga el error para que el
// caller pueda responder 500 al cliente en operaciones crÃ­ticas (ej. cerrar v2).
function saveDailyStateStrict(label) {
  return dbSet('daily_state', buildDailyStatePayload());
}

// â”€â”€ Load state from Supabase on startup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async function loadState() {
  try {
    const saved = await dbGet('daily_state');
    if(saved && saved.teorico) {
      state = { ...state, ...saved };
      console.log('State restored from Supabase âœ“ date:', saved.date);
    } else {
      console.log('No saved state found â€” starting fresh');
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

// â”€â”€ Save state to Supabase (debounced) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Debounce a 1500ms para agrupar escrituras frecuentes (varios users tecleando
// simultÃ¡neamente). Antes era 200ms â€” provocaba serializar el state completo
// (~7MB) por cada tecleo, lo que dispara el uso de RAM en Render free (512MB)
// y causa "heap out of memory". 1.5s aÃºn es lo suficientemente rÃ¡pido para no
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
    alertasWMS:     state.alertasWMS     || {},  // FIX (rev ChatGPT BLOQUEANTE v5.2.22): sin esto el cliente nunca recibÃ­a las alertas y el badge no aparecÃ­a
    date:           state.date,
    version:        state.version,
    activeUsers:    getActiveUsers(),
    locks:          getLocks()
  };
}

// â”€â”€ API â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/heartbeat', (req, res) => {
  const { name } = req.body;
  if(name) activeUsers[name] = Date.now();
  res.json({ active: getActiveUsers() });
});

app.get('/api/state', (req, res) => {
  var ps = publicState();
  try {
    // DiagnÃ³stico de tamaÃ±o: permite monitorear crecimiento del state sin crashear.
    // FIX (mar 2-jun-2026, server v20): campo stateSizeBytes agregado para auditorÃ­a.
    ps.stateSizeBytes = Buffer.byteLength(JSON.stringify(buildDailyStatePayload()), 'utf8');
  } catch(e) { /* no bloquear el estado si falla el cÃ¡lculo */ }
  res.json(ps);
});

// Full state save (from client). Merges, never deletes existing data.
app.post('/api/conteo', (req, res) => {
  const { cont, data, usuario } = req.body;
  if(!cont || !data) return res.status(400).json({ ok:false });
  state.fisico[cont] = data;
  addHistorial(usuario||'â€”', 'Conteo guardado', cont);
  // Debounced â€” agrupa escrituras concurrentes. addHistorial ya llama scheduleSave,
  // asÃ­ que aquÃ­ no hace falta llamar de nuevo.
  res.json({ ok:true, version:state.version });
});

app.post('/api/asign', (req, res) => {
  const { cont, name, action, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.asignaciones[cont]) state.asignaciones[cont] = [];
  if(action === 'add') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(usuario||'â€”', 'AsignaciÃ³n', cont+' â†’ '+name);
  } else if(action === 'remove') {
    state.asignaciones[cont] = state.asignaciones[cont].filter(n => n !== name);
    addHistorial(usuario||'â€”', 'AsignaciÃ³n removida', cont+' â† '+name);
  } else if(action === 'self') {
    if(!state.asignaciones[cont].includes(name)) state.asignaciones[cont].push(name);
    addHistorial(name, 'Auto-asignaciÃ³n', cont);
  }
  state.version++;
  // Debounced â€” addHistorial dentro de cada rama ya llamÃ³ scheduleSave.
  res.json({ ok:true, version:state.version });
});

// â”€â”€ CDG â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.get('/api/cdg', (req, res) => res.json(state.cdg||{}));

app.post('/api/cdg/save', (req, res) => {
  const { contId, items, usuario, tipo, fotoGral, fotos } = req.body;
  if(!contId) return res.status(400).json({ ok:false });
  // FIX (sÃ¡b 30-may-2026, v19): validaciÃ³n defensiva de items.
  // Si items llega null/undefined (payload mal formado, tablet vieja, etc.),
  // rechazar antes de tocar el state. Sin esto, items:null pasa la guardia
  // ((null||[]).length = 0) y luego state.cdg[contId].items = null rompe el conteo.
  if(!Array.isArray(items)) {
    return res.status(400).json({
      ok: false,
      error: 'No se guardÃ³: el conteo llegÃ³ sin lista de lÃ­neas vÃ¡lida. RecargÃ¡ la pantalla e intentÃ¡ de nuevo.'
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
  // FIX (sÃ¡b 30-may-2026, v19): protecciÃ³n contra sobreescritura accidental.
  // Escenario: Julio guarda 50 items. Otro usuario abre CDG con lista vacÃ­a y
  // guarda 1 item â†’ borra los 50 de Julio (Bug B1 CDG v1).
  // Guardia: si ya hay items y el nuevo array tiene menos de la mitad Y el
  // usuario no es el autor original â†’ rechazar con 409 y mensaje claro.
  // Nota: items ya validado como Array arriba, usar .length directo.
  var existentes    = (state.cdg[contId].items || []).length;
  var nuevos        = items.length;
  var autorOriginal = state.cdg[contId].autor || state.cdg[contId].lastEditor;
  if(existentes > 2 && nuevos < existentes / 2 && autorOriginal && autorOriginal !== usuario) {
    console.log('CDG save BLOQUEADO: ' + usuario + ' intentÃ³ reducir ' + contId + ' de ' + existentes + ' a ' + nuevos + ' items (autor: ' + autorOriginal + ')');
    return res.status(409).json({
      ok:    false,
      error: 'Este conteo ya tiene ' + existentes + ' lÃ­neas. RecargÃ¡ el conteo antes de guardar. Para reemplazarlo usÃ¡ desbloqueo de supervisor o CDG v2 multiusuario.'
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

// CDG Unlock (supervisor only â€” enforced client-side)
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
  addHistorial(usuario||'â€”', 'DesbloqueÃ³ CDG', contId);
  state.version++;
  // Debounced â€” addHistorial ya disparÃ³ scheduleSave.
  res.json({ ok:true });
});


// â”€â”€ cdgEnriquecerItemsWMS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// FIX (lun 1-jun-2026, server v20 rev): devuelve la UNIÃ“N de items CDG + SKUs WMS.
//
// CORRECCIÃ“N vs versiÃ³n anterior (solo hacÃ­a map sobre items CDG):
//   La versiÃ³n anterior agregaba teoricoWMS a SKUs que CDG sÃ­ contÃ³, pero dejaba
//   invisible cualquier SKU que existiera en WMS y CDG no hubiera contado.
//   Un SKU WMS con 50 unidades y qty CDG = 0 quedaba fuera de Hamilton â†’ faltante
//   invisible. Esto era bloqueante.
//
// Comportamiento correcto:
//   1. SKU en CDG y WMS â†’ qty = Validado CDG, teoricoWMS = cantidad WMS.
//   2. SKU en CDG, NO en WMS â†’ item sin cambios (sin teoricoWMS). Fallback legacy OK.
//   3. SKU en WMS, NO en CDG â†’ nueva lÃ­nea: qty=0, teoricoWMS=cantidad WMS.
//      Aparece en Hamilton: TeÃ³rico s/WMS=N, Validado CDG=0.
//
// DEUDA TÃ‰CNICA (no resuelta en este fix):
//   Los exports de Traslados CDG y updateSummaryCard usan item.qty para calcular
//   diferencias. Con los nuevos items WMS-only (qty=0), el resumen Hamilton puede
//   contar mÃ¡s faltantes de lo que muestra la tabla si no se actualiza esa lÃ³gica.
//   Fix pendiente: ajustar buildReporte / exportCDG para leer teoricoWMS como base
//   de diferencia cuando esCDG, igual que renderConteo ya hace.
//
// Si no hay WMS o falla la consulta â†’ devuelve items sin cambios (backward-compatible).
// Se usa en /api/cdg/finalizar (v1) y en el cierre CDG v2.
async function cdgEnriquecerItemsWMS(licenciaId, items) {
  if(!SUPABASE_URL || !SUPABASE_KEY || !licenciaId) return items || [];
  try {
    var wmsRows = await supabase('GET', 'cdg_wms', null,
      '?licencia_id=eq.' + encodeURIComponent(cdgNormId(licenciaId)) + '&limit=1');
    if(!Array.isArray(wmsRows) || !wmsRows.length) return items || []; // sin WMS â€” OK
    var wmsRow = wmsRows[0];
    var skusWMS = typeof wmsRow.skus === 'string' ? JSON.parse(wmsRow.skus) : (wmsRow.skus || []);

    // Mapa WMS: SKU_NORMALIZADO â†’ { cantidad, descripcion }
    var wmsMap = {};
    skusWMS.forEach(function(s){
      var sk = String(s.sku || '').trim().toUpperCase();
      if(!wmsMap[sk]) {
        wmsMap[sk] = { cantidad: 0, descripcion: s.descripcion || '' };
      }
      wmsMap[sk].cantidad += (Number(s.cantidad) || 0);
      // Preferir descripciÃ³n no vacÃ­a si la anterior era vacÃ­a
      if(!wmsMap[sk].descripcion && s.descripcion) wmsMap[sk].descripcion = s.descripcion;
    });

    // Conjunto de SKUs ya cubiertos por CDG (para detectar solo-WMS al final)
    var itemsCdg = Array.isArray(items) ? items : [];
    var skusCDG = {};
    var resultado = itemsCdg.map(function(item) {
      var sk = String(item.sku || '').trim().toUpperCase();
      skusCDG[sk] = true;
      if(wmsMap[sk] !== undefined) {
        // SKU en CDG y en WMS â†’ agregar teoricoWMS
        return Object.assign({}, item, { teoricoWMS: wmsMap[sk].cantidad });
      }
      // SKU en CDG pero no en WMS â†’ sin teoricoWMS (fallback legacy)
      return item;
    });

    // SKUs en WMS que CDG NO contÃ³ â†’ agregar como lÃ­neas con qty=0
    var tipoEfectivo = 'CDG'; // valor por defecto; el llamador lo conoce pero no se pasa aquÃ­
    Object.keys(wmsMap).forEach(function(sk){
      if(skusCDG[sk]) return; // ya cubierto por CDG
      var entrada = wmsMap[sk];
      resultado.push({
        sku:         sk,
        desc:        entrada.descripcion || '',
        qty:         0,           // Validado CDG = 0 (no fue contado)
        teoricoWMS:  entrada.cantidad,
        raw:         { origen: tipoEfectivo, status: tipoEfectivo + ' Validado' },
        _soloWMS:    true         // marca interna para diagnÃ³stico; no afecta renderConteo
      });
    });

    return resultado;
  } catch(e) {
    console.log('cdgEnriquecerItemsWMS warn (non-fatal):', e.message);
    return items || []; // fallo silencioso â€” no romper el cierre CDG
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
    // FIX (sÃ¡b 23-may-2026): raw.origen y raw.status ahora reflejan el tipo real
    // (KTM, CDG, Otros). Antes era hardcoded 'CDG' lo que hacÃ­a que el export
    // de Traslados mostrara "CDG" para TODOS los contenedores finalizados desde
    // esta secciÃ³n, incluyendo los KTM. Erick + Ever confirmaron 23-may que KTM
    // debe verse como categorÃ­a distinta en el export.
    var tipoEfectivo = tipo || 'CDG';
    // FIX (lun 1-jun-2026, server v20): enriquecer items con teoricoWMS desde cdg_wms.
    // qty = Validado CDG (el conteo fÃ­sico CDG). teoricoWMS = teÃ³rico del WMS si existe.
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
    // Asegurar que el tipo quede registrado tambiÃ©n si el contenedor ya existÃ­a
    // (defensivo, no rompe nada si ya estaba)
    if(tipo && !state.teorico[num].cdgTipo) state.teorico[num].cdgTipo = tipo;
  }
  addHistorial(usuario, 'CDG finalizado â†’ Traslado', contId+' â†’ '+num);
  state.version++;

  // FIX CRÃTICO (miÃ© 20-may-2026): mismo patrÃ³n que upload. Cancelar debounced
  // pending y awaitear el save, responder con error si Supabase falla.
  // FIX (rev Claude2): timeout 15s. FIX (rev ChatGPT): mensaje operativo.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'CDG final save'
    );
    console.log('CDG final save: persisted to Supabase âœ“');
  } catch(saveErr) {
    console.log('CDG final save FAILED:', saveErr.message);
    return res.status(500).json({
      ok: false,
      error: 'El CDG sÃ­ se procesÃ³, pero NO quedÃ³ guardado. No cierres la app. VolvÃ© a finalizarlo. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, traslado:num, version:state.version });
});

// â”€â”€ Costos â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/costos', upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const sn = wb.SheetNames.find(n =>
      n.toLowerCase().includes('existencia') || n.toLowerCase().includes('sap')
    ) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header:1, defval:'', raw:false });
    const hdr  = rows[0].map(h => norm(String(h)));
    const cSku  = findCol(hdr, ['articulo','artÃ­culo','sku','codigo']);
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
// FIX CRÃTICO (miÃ© 20-may-2026): dbSet ya no se hace sin await. Mismo bug
// que upload teorico â€” si Render reiniciaba despuÃ©s de responder, los costos
// se perdÃ­an silenciosamente.
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
    console.log('Costos save: persisted to Supabase âœ“');
    res.json({ ok:true, count: Object.keys(c).length });
  } catch(e) {
    console.log('Costos save FAILED:', e.message);
    res.status(500).json({
      ok: false,
      error: 'Los costos sÃ­ se procesaron, pero NO quedaron guardados. No cierres la app. VolvÃ© a subirlos. (' + e.message + ')'
    });
  }
});

app.get('/api/costos', (req, res) => res.json(state.costos));

// â”€â”€ Hallazgos â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/hallazgo', (req, res) => {
  const { hallazgo, action, id } = req.body;
  if(!state.hallazgos) state.hallazgos = [];
  if(action === 'add' && hallazgo) {
    state.hallazgos.push(hallazgo);
  } else if(action === 'edit' && id) {
    const idx = state.hallazgos.findIndex(h => h.id === id);
    if(idx >= 0) state.hallazgos[idx] = hallazgo;
  } else if(action === 'delete' && id) {
    // FIX (miÃ© 20-may-2026 noche): soporte para eliminar hallazgos desde el
    // botÃ³n "ðŸ—‘ Eliminar" del cliente (visible solo para Erick Vela). Sin
    // este case el server caerÃ­a en el branch implÃ­cito y reescribirÃ­a
    // Supabase reviviendo el hallazgo eliminado.
    state.hallazgos = state.hallazgos.filter(h => h.id !== id);
  }
  state.version++;
  // Debounced â€” hallazgos pueden venir en rÃ¡faga
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// â”€â”€ Metadata (puerta, fechaIngreso, fechaFurgon, placas per container) â”€â”€â”€â”€
app.post('/api/metadata', (req, res) => {
  const { cont, metadata, puerta, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false });
  if(!state.conteoMetadata) state.conteoMetadata = {};
  if(!state.puertas)        state.puertas        = {};
  if(metadata)             state.conteoMetadata[cont] = metadata;
  if(puerta !== undefined) state.puertas[cont]        = puerta;
  addHistorial(usuario||'â€”', 'Metadata actualizada', cont);
  state.version++;
  // Debounced â€” addHistorial ya disparÃ³ scheduleSave.
  res.json({ ok:true, version:state.version });
});

// FIX (mar 19-may-2026): ediciÃ³n manual de fechaCarga del teorico.
// Permite asignar/cambiar la fecha de trabajo de un contenedor desde la UI,
// por ejemplo para contenedores histÃ³ricos que no tenÃ­an fechaCarga.
// Formato esperado: 'YYYY-MM-DD' o cadena vacÃ­a ('') para limpiar.
//
// FIX (rev Claude2): valida tambiÃ©n la fecha calendÃ¡ricamente, no solo
// el formato. Rechaza overflows como 2026-02-30 o 9999-99-99.
//
// FIX (miÃ© 20-may-2026, post-deploy v8.1): mismo patrÃ³n que upload teorico.
// Antes el endpoint dependÃ­a de addHistorial â†’ scheduleSave debounced (1.5s)
// para persistir. Si Render reiniciaba en ese gap, la ediciÃ³n se perdÃ­a
// silenciosamente. Ahora cancelamos pending, awaitamos con timeout, y
// respondemos error si Supabase falla. Igual que /api/upload y CDG finalizar.
app.post('/api/teorico/fecha-carga', async (req, res) => {
  const { cont, fechaCarga, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false, error:'falta cont' });
  if(!state.teorico[cont]) return res.status(404).json({ ok:false, error:'cont no existe' });
  // Validar formato + fecha calendÃ¡rica real
  if(fechaCarga !== '' && fechaCarga !== null && fechaCarga !== undefined) {
    var s = String(fechaCarga);
    if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      return res.status(400).json({ ok:false, error:'formato debe ser YYYY-MM-DD o vacÃ­o' });
    }
    // Validar fecha calendÃ¡rica: el truco es comparar contra round-trip via Date.
    // Si entrÃ³ 2026-02-30, new Date lo normaliza a 2026-03-02, y la comparaciÃ³n falla.
    var d = new Date(s + 'T00:00:00Z');
    if(isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) {
      return res.status(400).json({ ok:false, error:'fecha calendÃ¡rica invÃ¡lida' });
    }
  }
  // FIX (miÃ© 28-may-2026, server v18): rollback ante fallo de persistencia.
  // Antes: si Supabase fallaba, el cambio quedaba aplicado en memoria pero NO
  // en Supabase. El cliente veÃ­a "error" pero la fecha cambiada en pantalla
  // hasta el prÃ³ximo polling. ConfusiÃ³n + divergencia.
  // Ahora: snapshot del valor previo + version + historial. En catch, restaura.
  //
  // FIX (rev cruzada Claude3 + ChatGPT, post-v18 borrador): addHistorial()
  // muta state.version e state.historial internamente. Snapshot debe capturar
  // AMBOS antes de cualquier mutaciÃ³n para revertir correctamente. Antes del
  // fix, el rollback hacÃ­a state.version-- (revertÃ­a solo 1 de 2 incrementos)
  // y dejaba el historial con la entrada del intento fallido.
  var prevFechaCarga    = state.teorico[cont].fechaCarga;
  var snapshotVersion   = state.version;
  var snapshotHistorial = state.historial.slice();

  state.teorico[cont].fechaCarga = fechaCarga || null;
  addHistorial(usuario||'â€”', 'CambiÃ³ fecha de trabajo a ' + (fechaCarga || '(vacÃ­o)'), cont);
  state.version++;

  // Cancelar debounced pending y await el save para garantizar persistencia.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'Fecha-carga save'
    );
    console.log('Fecha-carga save: persisted to Supabase âœ“');
  } catch(saveErr) {
    console.log('Fecha-carga save FAILED:', saveErr.message);
    // Rollback completo: restaurar valor + historial + version (revierte
    // tanto el version++ del endpoint como el version++ interno de addHistorial)
    state.teorico[cont].fechaCarga = prevFechaCarga;
    state.historial = snapshotHistorial;
    state.version   = snapshotVersion;
    return res.status(500).json({
      ok: false,
      error: 'La fecha NO quedÃ³ guardada. ReintentÃ¡. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, version:state.version, fechaCarga: state.teorico[cont].fechaCarga });
});

// FIX (sÃ¡b 23-may-2026): endpoint para clasificaciÃ³n manual de contenedores.
// Resuelve el bug B1 (server-memory desactualizada): antes el cliente llamaba
// saveDirectToSupabase con PATCH directo, lo que persistÃ­a en Supabase pero
// dejaba al server.js con state.teorico EN MEMORIA sin la clasificaciÃ³n.
// En el siguiente GET /api/state el server respondÃ­a con su memoria vieja
// y el cliente perdÃ­a visualmente la clasificaciÃ³n manual.
//
// Ahora el cliente llama este endpoint. El server:
//   1. Actualiza state.teorico[cont].clasificacion + clasificacionManual
//   2. Awaitea el save a Supabase con timeout 15s (mismo patrÃ³n que fecha-carga)
//   3. Cancela debounced pending para que no pise el cambio
//
// Permite limpiar el override (clasificacion=null, manual=false) para
// soportar la lÃ³gica "quitar manual si coincide con el automÃ¡tico" del UI.
app.post('/api/clasificacion/set', async (req, res) => {
  const { cont, clasificacion, manual, usuario } = req.body;
  if(!cont) return res.status(400).json({ ok:false, error:'falta cont' });
  if(!state.teorico[cont]) return res.status(404).json({ ok:false, error:'cont no existe' });

  // Validar clasificaciÃ³n si viene con valor
  if(clasificacion !== null && clasificacion !== undefined && clasificacion !== '') {
    var allowed = ['Auditado', 'En RevisiÃ³n', 'No auditado'];
    if(!allowed.includes(String(clasificacion))) {
      return res.status(400).json({ ok:false, error:'clasificacion invÃ¡lida (debe ser ' + allowed.join('|') + ')' });
    }
  }

  // FIX (miÃ© 28-may-2026, server v18): rollback ante fallo de persistencia.
  // Snapshot ANTES de mutar (los 2 campos que vamos a tocar).
  // FIX (rev cruzada post-v18 borrador): tambiÃ©n capturamos version + historial
  // porque addHistorial() los muta internamente. Sin esto, el rollback dejaba
  // version en N+1 y el historial con la entrada del intento fallido.
  var prevClasificacion       = state.teorico[cont].clasificacion;
  var prevClasificacionManual = state.teorico[cont].clasificacionManual;
  var prevTeniaClasif         = 'clasificacion'       in state.teorico[cont];
  var prevTeniaManual         = 'clasificacionManual' in state.teorico[cont];
  var snapshotVersion         = state.version;
  var snapshotHistorial       = state.historial.slice();

  // Si clasificacion es null/vacÃ­o Y manual no es true â†’ quitar override
  if((clasificacion === null || clasificacion === undefined || clasificacion === '') && !manual) {
    delete state.teorico[cont].clasificacion;
    delete state.teorico[cont].clasificacionManual;
    addHistorial(usuario||'â€”', 'QuitÃ³ clasificaciÃ³n manual', cont);
  } else {
    state.teorico[cont].clasificacion = clasificacion;
    state.teorico[cont].clasificacionManual = !!manual;
    addHistorial(usuario||'â€”', 'ClasificaciÃ³n ' + (manual ? 'manual' : 'auto') + ': ' + clasificacion, cont);
  }
  state.version++;

  // Cancelar debounced pending y await el save para garantizar persistencia.
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'ClasificaciÃ³n save'
    );
    console.log('ClasificaciÃ³n save: persisted to Supabase âœ“');
  } catch(saveErr) {
    console.log('ClasificaciÃ³n save FAILED:', saveErr.message);
    // Rollback completo: restaurar campos previos (respetando si existÃ­an),
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
      error: 'La clasificaciÃ³n NO quedÃ³ guardada. ReintentÃ¡. (' + saveErr.message + ')'
    });
  }

  res.json({
    ok: true,
    version: state.version,
    clasificacion: state.teorico[cont].clasificacion || null,
    manual: !!state.teorico[cont].clasificacionManual
  });
});

// FIX (sÃ¡b 23-may-2026): endpoint para borrar CDG.
// Misma motivaciÃ³n que /api/clasificacion/set: evitar bug B1.
// Cuando se borra un CDG, tambiÃ©n se borra el contenedor Traslado que
// se creÃ³ a partir de Ã©l (decisiÃ³n Erick 23-may: limpio total).
//
// IdentificaciÃ³n del Traslado asociado:
//   - El CDG tiene state.cdg[contId].traslado (asignado en /api/cdg/finalizar)
//   - El Traslado tiene state.teorico[num].cdgRef === contId
//   - Borramos AMBOS con todas sus referencias (teorico + fisico + asignaciones)
app.post('/api/cdg/delete', async (req, res) => {
  const { contId, usuario } = req.body;
  if(!contId) return res.status(400).json({ ok:false, error:'falta contId' });
  if(!state.cdg[contId]) return res.status(404).json({ ok:false, error:'CDG no existe' });

  // Identificar el contenedor Traslado asociado
  var trasladoNum = state.cdg[contId].traslado || null;

  // FIX (miÃ© 28-may-2026, server v18): rollback ante fallo de persistencia.
  // ANTES de borrar nada, snapshot completo de TODO lo que vamos a borrar.
  // Esto incluye: el CDG, el Traslado (en sus 6 ubicaciones), y los traslados
  // detectados por fallback (cdgRef). El rollback restaura cada referencia.
  //
  // FIX (rev cruzada post-v18 borrador): tambiÃ©n capturamos version + historial
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
  // no borraba si el valor era null (comÃºn en contenedores no contados).
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

  // Fallback: por si el traslado no estaba en .traslado pero sÃ­ en teorico con cdgRef
  // (escenario raro pero defensivo)
  rollbackData.trasladosFallback.forEach(function(t){
    deleteTrasladoRefs(t.num);
  });

  addHistorial(usuario||'â€”', 'CDG eliminado (+ Traslado asociado)', contId + (trasladoNum ? ' â†’ ' + trasladoNum : ''));
  state.version++;

  // Cancelar debounced pending y await el save
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'CDG delete save'
    );
    console.log('CDG delete save: persisted to Supabase âœ“ (CDG ' + contId + ', Traslado ' + (trasladoNum||'â€”') + ')');
  } catch(saveErr) {
    console.log('CDG delete save FAILED:', saveErr.message);
    // Rollback completo: CDG, traslado primario, traslados fallback, historial, version.
    // Usa los flags tenia* para distinguir "propiedad existÃ­a con valor null/falsy"
    // vs "propiedad no existÃ­a" y restaurar el estado EXACTO previo.
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
      error: 'El CDG NO se borrÃ³ (memoria restaurada). ReintentÃ¡. (' + saveErr.message + ')'
    });
  }

  res.json({ ok:true, version:state.version, deletedCDG:contId, deletedTraslado:trasladoNum });
});

// FIX (dom 24-may-2026 PM, server v15): endpoint para "Sincronizar con WMS".
// Cuando un contenedor estÃ¡ congelado (tiene fÃ­sico contado) y el WMS tiene
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
// Permisos: solo supervisores (validado en el cliente; el server confÃ­a
// pero loggea quiÃ©n hizo la sincronizaciÃ³n).
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
  // que detectaron ambos validadores: el fisico se alinea por Ã­ndice posicional
  // con teorico.items[idx]. Si el WMS elimina/reordena lÃ­neas, los conteos
  // fÃ­sicos quedan apuntando a SKUs equivocados.
  //
  // Estrategia: indexar el fisico viejo por el SKU que tenÃ­a, y reconstruir
  // un fisico nuevo donde cada conteo se asigna al Ã­ndice del MISMO SKU en
  // la lista nueva. Los SKUs eliminados del WMS pierden su conteo (decisiÃ³n
  // operativa: el WMS dice "esa lÃ­nea ya no existe"). Los SKUs agregados
  // quedan con fÃ­sico null (sin contar todavÃ­a, correcto).
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

  // Aplicar items del WMS, preservar resto de propiedades (auditorÃ­a)
  state.teorico[cont] = Object.assign({}, prev, {
    items: alerta.itemsWMS,
    meta:  alerta.metaWMS || prev.meta || {}   // FIX (server v16): aplicar meta del WMS al sincronizar
  });
  // Aplicar fisico re-mapeado (solo si habÃ­a fisico antes)
  if(newFisico) state.fisico[cont] = newFisico;

  // Contar cuÃ¡ntos conteos se mantuvieron / perdieron por SKUs eliminados
  var conteosMantenidos = 0, conteosPerdidos = 0;
  if(prevFisico && newFisico) {
    prevFisico.forEach(function(f){
      if(f && f.fisico !== undefined && f.fisico !== null && f.fisico !== '') {
        // Estaba contado antes. Ver si quedÃ³ en el nuevo.
        var sigueEnNuevo = newFisico.some(function(nf){
          return nf === f;
        });
        if(sigueEnNuevo) conteosMantenidos++;
        else conteosPerdidos++;
      }
    });
  }

  // Snapshot version + historial ANTES de mutar (addHistorial los mutarÃ¡).
  // FIX (rev cruzada post-v18 borrador): el patrÃ³n original solo hacÃ­a version--
  // que revertÃ­a solo 1 de 2 incrementos y dejaba historial con la entrada
  // del intento fallido. Mismo fix que aplicamos a los otros 4 endpoints v18.
  var snapshotVersion   = state.version;
  var snapshotHistorial = state.historial.slice();

  // Eliminar la alerta
  delete state.alertasWMS[cont];

  addHistorial(usuario||'â€”', 'SincronizÃ³ con WMS: ' + alerta.totalDiffs + ' cambios aplicados (conteos: ' + conteosMantenidos + ' mantenidos, ' + conteosPerdidos + ' perdidos)', cont);
  state.version++;

  // Persistir await (mismo patrÃ³n que clasificacion/set)
  if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
  try {
    await withTimeout(
      dbSet('daily_state', buildDailyStatePayload()),
      15000,
      'Sincronizar WMS save'
    );
    console.log('Sincronizar WMS save: persisted to Supabase âœ“');
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
      error: 'La sincronizaciÃ³n NO quedÃ³ guardada. ReintentÃ¡. (' + saveErr.message + ')'
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

// â”€â”€ Chat â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// â”€â”€ Field locking â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
  // Preserve null/undefined distinction â€” only fall back to prev when the
  // new value wasn't sent (undefined). Don't coerce null to 0.
  state.fisico[cont][idx] = {
    fisico:    fisico    !== undefined ? fisico    : prev.fisico,
    daniado:   daniado   !== undefined ? daniado   : prev.daniado,
    cobertura: cobertura !== undefined ? cobertura : (prev.cobertura || 'En revisiÃ³n'),
    // FIX (mar 19-may-2026): el cliente nuevo ya no envÃ­a calcExpr (campo
    // CÃ¡lculo eliminado de la UI). Pero clientes con cachÃ© viejo todavÃ­a
    // pueden mandar calcExpr: ''. Tratamos null/undefined/'' como "no enviar"
    // y preservamos siempre el histÃ³rico previo. Solo aceptamos valores
    // no-vacÃ­os (compatibilidad con clientes legacy que aÃºn calculan).
    calcExpr:  (calcExpr != null && calcExpr !== '') ? calcExpr : ((prev && prev.calcExpr) || ''),
    quien:     usuario,
    ts:        new Date().toLocaleString('es'),
    lastUser:  usuario,
    lastAt:    Date.now()
  };
  state.version++;
  // Debounced (1.5s) â€” agrupa multiples tecleos seguidos del mismo o varios
  // usuarios. Antes hacÃ­a saveDailyState() sync por cada keystroke, lo que
  // serializaba el state completo (~7MB) en cada request y agotaba el heap.
  scheduleSave();
  res.json({ ok:true, version:state.version });
});

// â”€â”€ Upload teÃ³rico â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// FIX CRÃTICO (miÃ© 20-may-2026): el endpoint era sync con respuesta antes
// del save real. Eso causÃ³ pÃ©rdida silenciosa de contenedores ayer (HP26-0357,
// HP26-0591 entre otros). Dos problemas que se combinaban:
//   1) saveDailyState() retornaba Promesa pero NO se awaiteaba. La respuesta
//      {ok:true} se enviaba antes que Supabase confirmara. Si Render
//      reiniciaba en esos ms, el upload se perdÃ­a.
//   2) Si habÃ­a un scheduleSave debounced pendiente de otro usuario, los
//      dos saves competÃ­an en paralelo. El debounced podÃ­a ganar y pisar
//      el upload con un state pre-upload en memoria.
// Fix: (a) cancelar saveTimer pendiente antes del upload, (b) hacer await
// del save crÃ­tico, (c) responder con error si el save falla.
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    const usuario = req.body.usuario || 'â€”';

    // FIX (miÃ© 28-may-2026, server v18): rollback ante fallo de persistencia.
    // mergeSheet muta state.teorico, state.fisico y state.alertasWMS reemplazando
    // entradas por contenedor. Snapshot somero de esos 3 objetos preserva las
    // referencias previas; al revertir, los contenedores modificados vuelven a
    // apuntar a sus objetos originales (los no tocados nunca cambiaron).
    // Costo: tres spreads superficiales (~ms), no deep clone.
    //
    // LIMITACIÃ“N CONOCIDA del snapshot somero (deuda en backlog):
    // mergeSheet tambiÃ©n muta PROPIEDADES de objetos existentes â€” concretamente
    // hace `state.teorico[cont].cdgValidado = true` para contenedores que ya
    // tienen `fromCDG=true` (lÃ­nea ~1352). Como el snapshot guarda referencias,
    // esa mutaciÃ³n NO se revierte al hacer rollback: el contenedor preserva su
    // objeto original, pero ese objeto ahora tiene `cdgValidado=true` adentro.
    // Severidad baja:
    //   - el flag no destruye datos (solo marca como validado un contenedor
    //     que ya tenÃ­a fromCDG=true; semÃ¡nticamente coherente)
    //   - el prÃ³ximo upload exitoso lo normaliza igual
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
        addHistorial(usuario, 'TeÃ³rico cargado', type+' â€” '+count+' contenedores');
      }
    });
    if(loaded.length === 0) {
      const rows  = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:'', raw:false });
      const count = mergeSheet(rows, 'General');
      if(count > 0) loaded.push(wb.SheetNames[0]+'('+count+')');
      addHistorial(usuario, 'TeÃ³rico cargado', loaded.join());
    }
    state.version++;

    // Cancelar cualquier debounced save pendiente: si existe, contendrÃ­a
    // una copia del state PRE-upload y al ejecutarse pisarÃ­a el upload.
    if(saveTimer) { clearTimeout(saveTimer); saveTimer = null; }

    // Await: NO responder Ã©xito hasta que Supabase confirme.
    // Usar buildDailyStatePayload directo (no saveDailyState que swallow errores).
    // FIX (rev Claude2): timeout 15s. FIX (rev ChatGPT): mensaje operativo.
    try {
      await withTimeout(
        dbSet('daily_state', buildDailyStatePayload()),
        15000,
        'Upload save'
      );
      console.log('Upload save: persisted to Supabase âœ“');
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
        error: 'El archivo NO se guardÃ³ (memoria restaurada). VolvÃ© a subirlo. (' + saveErr.message + ')'
      });
    }

    res.json({ ok:true, loaded });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// â”€â”€ Fotos: upload a Supabase Storage â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Endpoint que recibe una foto en base64 desde el cliente y la sube al
// bucket 'app-fotos' de Supabase Storage. Devuelve la URL pÃºblica para
// que el cliente la guarde en el campo `foto` del hallazgo/CDG en vez de
// almacenar el base64 completo dentro del state.
//
// FIX (miÃ© 20-may-2026): migraciÃ³n fotos a Storage para resolver
// PayloadTooLargeError + state inflado por base64. Decisiones del diseÃ±o:
//   - bucket pÃºblico (URLs largas+random como protecciÃ³n por oscuridad)
//   - validaciÃ³n de tamaÃ±o max 5MB (cliente debe comprimir antes)
//   - validaciÃ³n de mime type permitidos
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
        // Construir URL pÃºblica. Para buckets pÃºblicos:
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
      return res.status(400).json({ ok:false, error:'kind invÃ¡lido (debe ser hallazgo|cdg-gral|cdg-sku)' });
    }
    // refId sanitizado: solo alfanumÃ©rico, guiones, slashes, puntos.
    if(!/^[A-Za-z0-9_\-./]+$/.test(String(refId))) {
      return res.status(400).json({ ok:false, error:'refId invÃ¡lido (solo A-Z 0-9 _ - . /)' });
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
      return res.status(400).json({ ok:false, error:'base64 invÃ¡lido' });
    }

    if(buffer.length > MAX_PHOTO_BYTES) {
      return res.status(413).json({
        ok: false,
        error: 'foto muy grande (' + Math.round(buffer.length/1024) + ' KB). MÃ¡ximo: ' + Math.round(MAX_PHOTO_BYTES/1024) + ' KB. ComprimÃ­ antes de subir.'
      });
    }

    // Construir path en el bucket. ExtensiÃ³n a partir del mime.
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
        error: 'No se pudo subir la foto al Storage. ReintentÃ¡. (' + uploadErr.message + ')'
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

// â”€â”€ Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
       r.some(h => h === 'numero' || h === 'nÃºmero')) {
      hdrRowIdx = ri;
      break;
    }
  }
  const hdr      = rows[hdrRowIdx].map(h => norm(h));
  const dataRows = rows.slice(hdrRowIdx + 1);
  const colCont = findCol(hdr, ['numero','nÃºmero']);
  const colSku  = findCol(hdr, ['sku']);
  const colQty  = findCol(hdr, ['cant.','cant','cantidad']);
  let colDesc = -1;
  for(let i = colSku + 1; i < hdr.length; i++) {
    if(hdr[i].includes('nombre') || hdr[i].includes('descripcion')) { colDesc = i; break; }
  }
  if(colDesc < 0) colDesc = findCol(hdr, ['nombre','descripcion']);
  if(colCont < 0 || colSku < 0 || colDesc < 0 || colQty < 0) return 0;

  const g = (row, i) => i >= 0 ? row[i] : '';
  // FIX (sÃ¡b 23-may-2026): el WMS de Embarques trae DOS columnas relevantes para
  // proveedor: 'Proveedor' (cÃ³digo, ej. HPP001199) y 'Nombre' (razÃ³n social, ej.
  // LONGTAI TRADING FZCO). Antes solo se capturaba el cÃ³digo en codProv, asÃ­ que
  // raw.nomProv quedaba undefined y getOrigenDesc (export) recibÃ­a el cÃ³digo â†’
  // 'DescripciÃ³n de origen' salÃ­a como el cÃ³digo en vez de Truper China/MÃ©xico.
  // Capturamos el 'Nombre' que aparece JUSTO DESPUÃ‰S de la columna Proveedor,
  // para no confundirlo con el 'Nombre' de descripciÃ³n del SKU (que va despuÃ©s de SKU).
  const cCodProvIdx = findCol(hdr, ['proveedor','cÃ³digo de proveedor','codigo de proveedor']);
  let cNomProvIdx = -1;
  if(cCodProvIdx >= 0) {
    for(let i = cCodProvIdx + 1; i < hdr.length; i++) {
      if(hdr[i].includes('nombre')) { cNomProvIdx = i; break; }
      // No buscar mÃ¡s allÃ¡ del SKU: el 'Nombre' de descripciÃ³n va despuÃ©s del SKU
      if(hdr[i] === 'sku' || hdr[i].includes('sku')) break;
    }
  }
  const ex = {
    cFecha:    findCol(hdr, ['fecha']),
    cCodProv:  cCodProvIdx,
    cNomProv:  cNomProvIdx,
    cLineas:   findCol(hdr, ['lineas','lÃ­neas']),
    cStatus:   findCol(hdr, ['status']),
    cOC:       findCol(hdr, ['orden de compra','# orden']),
    cIngr:     findCol(hdr, ['ingresado']),
    cColoc:    findCol(hdr, ['colocado']),
    cFalt:     findCol(hdr, ['faltantes']),
    cSobr:     findCol(hdr, ['sobrantes']),
    cDan:      findCol(hdr, ['daÃ±ado','danado']),
    cOrigen:   findCol(hdr, ['origen']),
    cIngreso:  findCol(hdr, ['# ingreso','ingreso']),
    cDocSap:   findCol(hdr, ['doc. sap','doc sap']),
    cTipo:     findCol(hdr, ['tipo']),
    cUnidad:   findCol(hdr, ['unidad']),
    cUnidades: findCol(hdr, ['unidades']),
    cDestino:  findCol(hdr, ['destino'])
  };
  const newConts = {};
  // FIX (dom 25-may-2026, server v16): OPTIMIZACIÃ“N DE TAMAÃ‘O DEL STATE.
  // Antes, cada item guardaba 18 campos en raw. Medido contra el WMS real,
  // 10 de esos campos (fecha, codProv, nomProv, lineas, status, origen,
  // ingreso, docSap, tipo, destino) son CONSTANTES dentro de un mismo
  // contenedor â€” se repetÃ­an idÃ©nticos en cada una de sus ~150 lÃ­neas.
  // Ahora esos campos van UNA sola vez a newMeta[cont] (nivel contenedor),
  // y raw queda solo con los que VARÃAN por lÃ­nea (oc, ingresado, colocado,
  // faltantes, sobrantes, daniado, unidad, unidades). ReducciÃ³n ~70% del teorico.
  // Compatibilidad: el cliente lee raw heredando de meta (Object.assign), asÃ­
  // que los exports siguen funcionando y los contenedores VIEJOS (con raw
  // completo y sin meta) tambiÃ©n, porque su raw viejo gana sobre el meta ausente.
  const newMeta = {};
  dataRows
    .filter(r => r && r.some(c => String(c).trim() !== ''))
    .forEach(row => {
      const cont = String(row[colCont]||'').trim();
      if(!cont) return;
      // FIX (mar 26-may-2026, server v17): LISTA BLANCA de nombre de contenedor.
      // El export de Power BI trae texto al pie ("Filtros aplicados: ... Embarque
      // es HP26-XXXX Aplica es 1", "Total general", etc.) que el parser leÃ­a como
      // contenedores basura. Ahora SOLO se acepta una fila si su nombre tiene el
      // formato real: empieza con H o U, dÃ­gitos, guion, dÃ­gitos, con sufijo
      // opcional (ej. /CS). HP26-0540, H274-0123, U25-161410, U147-0089.
      // Cualquier otra cosa (texto largo, notas, totales) se ignora automÃ¡ticamente.
      if(!/^[HU][A-Z0-9]*-[0-9]+(\/[A-Z0-9]+)?$/i.test(cont)) return;
      // FIX (mar 26-may-2026, server v17): los contenedores terminados en /CS
      // existen en la base del WMS pero el equipo NO los usa. Antes habÃ­a que
      // borrarlos a mano tras cada carga del maestro porque se volvÃ­an a meter.
      // Ahora se ignoran automÃ¡ticamente al cargar.
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
          // Solo campos que VARÃAN por lÃ­nea:
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
  // Contexto: el equipo va a subir el Excel del Power BI varias veces al dÃ­a.
  // El WMS puede cambiar cantidades o eliminar lÃ­neas de un contenedor que
  // YA empezamos a contar. Si dejÃ¡ramos que el upload sobreescriba, el equipo
  // perderÃ­a referencia (fÃ­sico contado vs teÃ³rico modificado).
  //
  // Regla: si state.fisico[cont] tiene AL MENOS 1 lÃ­nea con valor (contada),
  // el contenedor se considera "frozen" y NO se actualiza su teorico desde el
  // upload. EN CAMBIO, se registra una alerta en state.alertasWMS para que el
  // supervisor decida si "descongelar" y sincronizar manualmente.
  //
  // Helper: detectar si un contenedor ya tiene fÃ­sico contado
  function hasPhysicalCounted(contKey) {
    var f = state.fisico && state.fisico[contKey];
    if(!Array.isArray(f)) return false;
    for(var i = 0; i < f.length; i++) {
      var x = f[i];
      if(x && x.fisico !== undefined && x.fisico !== null && x.fisico !== '') return true;
    }
    return false;
  }

  // Helper: comparar items prev vs new para detectar diferencias especÃ­ficas
  function compareItems(prevItems, newItems) {
    var prev = Array.isArray(prevItems) ? prevItems : [];
    var nw = Array.isArray(newItems) ? newItems : [];
    var diffs = [];
    // Index por SKU para comparaciÃ³n
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

    // â”€â”€ FREEZE check (server v15) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    // Si el contenedor ya tiene fÃ­sico contado, NO actualizar items.
    // Pero SÃ comparar y registrar alerta si hay cambios.
    var isFrozen = prev.items && hasPhysicalCounted(cont);

    if(isFrozen) {
      congeladosCount++;
      var diffs = compareItems(prev.items, newConts[cont]);
      if(diffs.length > 0) {
        congeladosConCambio++;
        // Guardar alerta CON los items nuevos del WMS. Esto permite al
        // supervisor "Sincronizar con WMS" despuÃ©s, aplicando estos items.
        // NO acumular alertas viejas: cada upload reemplaza la alerta con
        // el snapshot mÃ¡s reciente del WMS, para no inflar el state.
        state.alertasWMS[cont] = {
          detectadoEn: hoy,
          itemsAntes:  prev.items.length,
          itemsNuevos: newConts[cont].length,
          diffs:       diffs.slice(0, 30),  // lÃ­mite defensivo: top 30 diffs
          totalDiffs:  diffs.length,
          // Snapshot completo de items del WMS para "Sincronizar" despuÃ©s
          // (necesario porque sin esto el supervisor no podrÃ­a aplicar el WMS
          // ya que el upload solo se guarda si NO estÃ¡ congelado)
          itemsWMS:    newConts[cont],
          metaWMS:     newMeta[cont] || {}   // FIX (server v16): meta para aplicar al sincronizar
        };
      } else {
        // Sin diferencias: limpiar cualquier alerta vieja para este contenedor
        if(state.alertasWMS[cont]) delete state.alertasWMS[cont];
      }
      // Solo actualizar metadatos no-crÃ­ticos del raw del primer item
      // (status del WMS, etc.) â€” pero NO items
      // Por simplicidad y seguridad: NO tocamos nada cuando estÃ¡ frozen.
      // El supervisor puede sincronizar manualmente vÃ­a /api/wms/sincronizar.
      return;
    }

    // Contenedor SIN fÃ­sico contado: comportamiento normal (preservar auditorÃ­a)
    state.teorico[cont] = Object.assign({}, prev, {
      items: newConts[cont],
      meta:  newMeta[cont] || prev.meta || {},   // FIX (server v16): meta a nivel contenedor
      type,
      fechaCarga: fechaCargaPrev || hoy
    });
    // Si habÃ­a una alerta vieja, limpiarla (el upload nuevo se aplicÃ³ OK)
    if(state.alertasWMS && state.alertasWMS[cont]) delete state.alertasWMS[cont];
    // Preserve existing fisico data â€” never overwrite conteo work
    if(!state.fisico[cont]) state.fisico[cont] = null;
  });

  console.log('mergeSheet:', Object.keys(newConts).length, 'contenedores procesados,',
              congeladosCount, 'congelados (' + congeladosConCambio + ' con cambios del WMS)');
  return Object.keys(newConts).length;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// CDG v2 â€” MÃ³dulo colaborativo (server v19, 29-may-2026)
// Adaptado sÃ¡b 30-may-2026: corregido comentario service_roleâ†’anon,
// rutas sku-catalog/* movidas antes de /:id (bug de routing Express).
//
// Arquitectura:
//   cdg_meta_{id}  â†’ registro en app_state con metadata, usuarios, bitÃ¡cora
//   cdg_lineas     â†’ tabla Supabase independiente, una fila por lÃ­nea
//
// Esto evita reescribir el array completo de lÃ­neas en cada operaciÃ³n.
// Dos usuarios agregando lÃ­neas simultÃ¡neamente hacen INSERTs independientes,
// sin conflicto de versiÃ³n.
//
// IMPORTANTE: estos endpoints NO tocan state (el monolito Hamilton).
// NO aparecen en buildDailyStatePayload() ni en publicState().
// La red de seguridad saveDirectToSupabase del cliente sigue operando
// sobre daily_state como antes â€” no se ve afectada.
//
// Backup del mÃ³dulo CDG v1 (legacy): los endpoints /api/cdg/* originales
// se conservan sin cambios mÃ¡s abajo para rollback en caso necesario.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â”€â”€ Helpers CDG v2 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

// Genera un UUID v4 simple (sin dependencias externas)
// FIX (lun 1-jun-2026, v19): lock en memoria por licencia para operaciones CDG v2.
// Resuelve la race condition donde un POST /linea en vuelo puede insertar despuÃ©s
// de que el cierre ya leyÃ³ las lÃ­neas para Hamilton.
// Estructura: { [licenciaId]: { activeWrites: 0, closing: false } }
var cdgLocks = {};
function cdgLockAcquire(licenciaId) {
  if(!cdgLocks[licenciaId]) cdgLocks[licenciaId] = { activeWrites: 0, closing: false };
  if(cdgLocks[licenciaId].closing) return false; // rechazar â€” cierre en progreso
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
    // FIX: lanzar error en lugar de proceder â€” el cierre no puede correr con
    // writes activos. El caller limpia el lock y responde error al cliente.
    throw new Error('Timeout esperando que terminen ' + l.activeWrites + ' escritura(s) activa(s). ReintentÃ¡ en unos segundos.');
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
  // claves duplicadas por espacios o mayÃºsculas (iPad autocapitaliza).
  return await dbGet('cdg_meta_' + cdgNormId(licenciaId));
}

// Guarda la metadata de una licencia CDG en app_state
async function cdgSaveMeta(licenciaId, meta) {
  await dbSet('cdg_meta_' + cdgNormId(licenciaId), meta);
}

// FIX (lun 1-jun-2026, v19): helper central de normalizaciÃ³n.
// Todos los endpoints v2 usan cdgNormId() al extraer req.params.id y al
// construir licencia_id en filas de cdg_lineas. Garantiza que meta y lÃ­neas
// usen siempre la misma clave canÃ³nica (sin espacios, mayÃºsculas).
function cdgNormId(id) {
  return String(id || '').trim().toUpperCase();
}

// Lee las lÃ­neas de una licencia desde cdg_lineas (tabla Supabase).
// Carga inicial (sin desde): solo lÃ­neas NO eliminadas.
// Delta (con desde): todas las lÃ­neas modificadas despuÃ©s de ese momento,
// INCLUYENDO las eliminadas â€” necesario para que el polling propague deletes
// a otras tablets. El cliente descarta las que tengan eliminada=true.
async function cdgGetLineas(licenciaId, desde) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return [];
  var nId   = cdgNormId(licenciaId);
  var query = '?licencia_id=eq.' + encodeURIComponent(nId)
            + '&order=ts_creacion.asc';
  if(desde) {
    // Delta: incluir eliminadas para que las tablets las descarten localmente
    // FIX (lun 1-jun-2026, v19): usar gte. (mayor o igual) en lugar de gt.
    // Con gt. estricto, dos lÃ­neas con el mismo ts_modif pierden la del borde
    // del cursor â€” nunca llega en el delta. gte. la reenvÃ­a, y el merge del
    // cliente la descarta como duplicado (findIndex por id). Sin pÃ©rdida de datos.
    query += '&ts_modif=gte.' + encodeURIComponent(desde);
  } else {
    // Carga completa: solo activas
    query += '&eliminada=eq.false';
  }
  var rows = await supabase('GET', 'cdg_lineas', null, query);
  return Array.isArray(rows) ? rows : [];
}

// Inserta una lÃ­nea nueva en cdg_lineas
async function cdgInsertLinea(linea) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var rows = await supabase('POST', 'cdg_lineas', linea, '');
  return Array.isArray(rows) ? rows[0] : rows;
}

// Actualiza una lÃ­nea existente (solo campos permitidos)
async function cdgUpdateLinea(lineaId, patch, autorEsperado) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  patch.ts_modif = new Date().toISOString();
  var query = '?id=eq.' + encodeURIComponent(lineaId)
            + '&autor=eq.' + encodeURIComponent(autorEsperado)
            + '&eliminada=eq.false';
  var rows = await supabase('PATCH', 'cdg_lineas', patch, query);
  return Array.isArray(rows) ? rows[0] : rows;
}

// Soft-delete de una lÃ­nea (solo el autor puede borrar la suya)
async function cdgSoftDeleteLinea(lineaId, autorEsperado) {
  if(!SUPABASE_URL || !SUPABASE_KEY) return null;
  var patch = { eliminada: true, ts_modif: new Date().toISOString() };
  var query = '?id=eq.' + encodeURIComponent(lineaId)
            + '&autor=eq.' + encodeURIComponent(autorEsperado)
            + '&eliminada=eq.false';
  var rows = await supabase('PATCH', 'cdg_lineas', patch, query);
  return Array.isArray(rows) ? rows[0] : rows;
}

// Actualiza tsUltimaModifLineas en la metadata cuando cambian las lÃ­neas
// FIX (sÃ¡b 30-may-2026, server v19): re-leer meta fresco desde DB antes de bump.
// SIN este fix: si usuario A leyÃ³ meta con estado='activo' y luego usuario B
// cerrÃ³ la licencia (estado='cerrado'), el write de A sobreescribe con su meta
// stale y reabre la licencia â€” riesgo operativo real.
// CON el fix: siempre escribimos sobre la base mÃ¡s reciente de Supabase,
// preservando estado/finalizador/bitacora actuales. Solo propagamos lastActivity
// del usuario local (campo seguro de pisar â€” a lo sumo queda unos segundos atrÃ¡s).
// deltaTotalLineas: +1 al insertar, -1 al eliminar, 0 o undefined al editar.
// Sin deltaTotalLineas el contador no se actualiza (no lo podemos inferir).
// Costo: 1 GET extra por operaciÃ³n de lÃ­nea. Aceptable para equipos de 4-6.
async function cdgBumpLineasTs(licenciaId, metaLocal, deltaTotalLineas) {
  var metaFresh = await cdgGetMeta(licenciaId);
  var base = metaFresh || metaLocal; // fallback si la licencia fue borrada
  base.tsUltimaModifLineas = new Date().toISOString();
  base.version = (base.version || 0) + 1;
  // Aplicar delta de totalLineas sobre el valor FRESCO (no el stale de metaLocal)
  if(deltaTotalLineas !== undefined && deltaTotalLineas !== 0) {
    base.totalLineas = Math.max(0, (base.totalLineas || 0) + deltaTotalLineas);
  }
  // Propagar solo lastActivity del usuario que operÃ³ (nunca estado de licencia)
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
  // mientras la operaciÃ³n estaba en vuelo y pueda responder 409 al cliente.
  return { estadoActual: base.estado };
}

// Agrega una entrada a la bitÃ¡cora de la licencia
function cdgBitacora(meta, usuario, accion, detalle) {
  if(!meta.bitacora) meta.bitacora = [];
  meta.bitacora.push({
    ts:      new Date().toISOString(),
    usuario: usuario || 'â€”',
    accion:  accion,
    detalle: detalle || ''
  });
  // LÃ­mite defensivo: mÃ¡x 200 entradas en bitÃ¡cora
  if(meta.bitacora.length > 200) meta.bitacora = meta.bitacora.slice(-200);
}

// Verifica si una licencia acepta nuevas lÃ­neas
function cdgAceptaLineas(meta) {
  return meta && meta.estado === 'activo';
}

// â”€â”€ GET /api/cdg/v2/listar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Lista licencias CDG v2.
// Sin parÃ¡metros: activas, pausadas y cerradas en Ãºltimas 48h (para captura/manifiesto).
// ?historico=true: todas las licencias cerradas sin lÃ­mite de fecha (para Reportes).
//   En modo histÃ³rico, enriquece cada licencia con totalUnidades y costoTotal
//   calculados desde cdg_lineas vigentes (eliminada=false).
// FIX (lun 1-jun-2026, server v20): modo histÃ³rico para Reportes/Historial CDG.
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
          // Modo histÃ³rico: incluir todas (activas, pausadas, cerradas sin lÃ­mite de fecha)
          incluir = true;
        } else {
          // Modo operativo: activas, pausadas, cerradas recientes (Ãºltimas 48h)
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
            // Agregados: se llenan en modo histÃ³rico; null en modo operativo
            totalUnidades: null,
            costoTotal:    null
          });
        }
      } catch(e) { /* skip malformed */ }
    });

    // Modo histÃ³rico: enriquecer con agregados desde cdg_lineas
    if(historico && licencias.length > 0) {
      try {
        // Una sola query: todas las lÃ­neas vigentes de todas las licencias encontradas
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

// â”€â”€ POST /api/cdg/v2/crear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Crea una nueva licencia CDG.
// Body: { id, tipo, usuario }
// id: correlativo U25-XXXX. Si no se pasa, se genera automÃ¡ticamente.
app.post('/api/cdg/v2/crear', async (req, res) => {
  var { id, tipo, usuario } = req.body;
  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });

  // Generar ID si no se pasÃ³
  if(!id) {
    var fecha = new Date();
    var mm = String(fecha.getMonth() + 1).padStart(2, '0');
    var dd = String(fecha.getDate()).padStart(2, '0');
    var rand = Math.floor(Math.random() * 9000) + 1000;
    id = 'U25-' + mm + dd + rand;
  }
  // FIX (lun 1-jun-2026, v19): normalizar id con trim().toUpperCase() para evitar
  // duplicados por espacios o mayÃºsculas (iPad autocapitaliza la primera letra).
  id = cdgNormId(id); // usar helper central

  // Verificar que no exista ya
  try {
    var existing = await cdgGetMeta(id);
    if(existing) {
      return res.status(409).json({ ok: false, error: 'Ya existe una licencia con ese ID. UsÃ¡ /api/cdg/v2/' + id + ' para unirte.' });
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
    res.status(500).json({ ok: false, error: 'No se pudo crear la licencia. ReintentÃ¡. (' + e.message + ')' });
  }
});

// FIX (sÃ¡b 30-may-2026, server v19): rutas /sku-catalog/* ANTES de /:id.
// Express hace match en orden de definiciÃ³n. Si /:id va primero, captura
// "sku-catalog" como parÃ¡metro :id y las rutas de catÃ¡logo nunca se alcanzan.
// â”€â”€ GET /api/cdg/v2/sku-catalog/sugerir â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Devuelve hasta 15 SKUs que matcheen el texto ingresado (sku o descripcion).
// Usado para autocomplete en el form CDG v2. Sin cargar los 41k SKUs al cliente.
// Query: ?q=texto (mÃ­nimo 2 chars para evitar queries demasiado amplias)
// FIX (lun 1-jun-2026, server v19): endpoint nuevo para autocomplete.
app.get('/api/cdg/v2/sku-catalog/sugerir', async (req, res) => {
  var q = String(req.query.q || '').trim();
  if(q.length < 2) return res.json({ ok: true, resultados: [] });

  try {
    if(!SUPABASE_URL || !SUPABASE_KEY) {
      return res.json({ ok: true, resultados: [] });
    }
    // Buscar por SKU exacto primero (si es numÃ©rico) O por descripciÃ³n ilike
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

// â”€â”€ GET /api/cdg/v2/sku-catalog/buscar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Busca la descripciÃ³n y costo de un SKU en el catÃ¡logo.
// Query: ?sku=XXXX
// FIX (lun 1-jun-2026, server v19): incluir costo en select y respuesta.
// Antes solo devolvÃ­a sku,descripcion â€” el cliente no podÃ­a autollenar el costo.
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
      // costo: devolver como nÃºmero o null (nunca como string con sÃ­mbolo)
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

// â”€â”€ POST /api/cdg/v2/sku-catalog/upload â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Carga o actualiza el catÃ¡logo de SKUs desde un archivo Excel/CSV.
// Hace UPSERT por SKU (actualiza si existe, inserta si no).
// Multer field: 'file'. Columnas esperadas: SKU + DescripciÃ³n (flexible).
app.post('/api/cdg/v2/sku-catalog/upload', upload.single('file'), async (req, res) => {
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    var wb   = XLSX.read(req.file.buffer, { type: 'buffer', raw: false });
    var sn   = wb.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: false });

    if(rows.length < 2) {
      return res.status(400).json({ ok: false, error: 'El archivo estÃ¡ vacÃ­o o solo tiene encabezado.' });
    }

    var hdr  = rows[0].map(function(h) { return norm(String(h)); });
    var cSku  = findCol(hdr, ['sku', 'articulo', 'artÃ­culo', 'codigo', 'cÃ³digo']);
    var cDesc = findCol(hdr, ['descripcion', 'descripciÃ³n', 'desc', 'nombre', 'producto']);

    if(cSku < 0 || cDesc < 0) {
      return res.status(400).json({
        ok: false,
        error: 'No se encontraron columnas SKU y DescripciÃ³n. Encabezados detectados: ' + rows[0].join(', ')
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
      return res.status(400).json({ ok: false, error: 'No se encontraron filas vÃ¡lidas (SKU + DescripciÃ³n requeridos).' });
    }

    // Upsert en lotes de 500 para no superar lÃ­mites de Supabase
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

// â”€â”€ GET /api/cdg/v2/:id â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Polling del estado de la licencia. El cliente pasa su Ãºltimo timestamp
// conocido de lÃ­neas para recibir solo el delta.
// Query params: lineasDesde (ISO timestamp, opcional)
app.get('/api/cdg/v2/:id', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineasDesde = req.query.lineasDesde || null;

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });

    // FIX (sÃ¡b 30-may-2026, server v19): siempre consultar cdg_lineas cuando
    // el cliente manda lineasDesde. La optimizaciÃ³n anterior (skip si
    // tsUltimaModifLineas no cambiÃ³) ocultaba lÃ­neas cuando el bump de metadata
    // fallaba â€” la lÃ­nea existÃ­a en cdg_lineas pero ninguna tablet la veÃ­a
    // hasta hacer un refresh completo.
    // Sin la optimizaciÃ³n: 1 query extra a Supabase por poll (aceptable).
    var lineas = await cdgGetLineas(licenciaId, lineasDesde);
    var respuesta = { ok: true, meta: meta, lineas: lineas };

    res.json(respuesta);
  } catch(e) {
    console.log('CDG v2 GET', licenciaId, 'error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// â”€â”€ POST /api/cdg/v2/:id/linea â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Agrega una lÃ­nea nueva a la licencia.
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

  // Lock: registrar escritura activa. Si la licencia estÃ¡ cerrando â†’ rechazar.
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se estÃ¡ cerrando. Ya no se pueden agregar lÃ­neas.' });
  }

  var fotosArr = Array.isArray(fotos) ? fotos : [];
  if(fotosArr.length > 3) {
    cdgLockRelease(licenciaId);
    return res.status(400).json({ ok: false, error: 'mÃ¡ximo 3 fotos por lÃ­nea' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(!cdgAceptaLineas(meta)) {
      return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ ' + meta.estado + '. No se pueden agregar lÃ­neas.' });
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
    // FIX (sÃ¡b 30-may-2026, server v19): bump de metadata en try separado.
    // Si cdgInsertLinea ya persistiÃ³ la lÃ­nea y cdgBumpLineasTs falla,
    // el catch externo responderÃ­a 500 â†’ el cliente reintentarÃ­a â†’ lÃ­nea duplicada.
    // Separando el try: si el bump falla, la lÃ­nea estÃ¡ guardada, loggeamos y
    // respondemos 200 igual. El contador totalLineas se puede recalcular desde
    // cdg_lineas; tsUltimaModifLineas se corrige en el prÃ³ximo bump exitoso.
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
      // Detectar si la licencia se cerrÃ³ mientras esta operaciÃ³n estaba en vuelo
      if(bumpResult && bumpResult.estadoActual === 'cerrado') {
        licenciaCerradaMidFlight = true;
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata FAILED (lÃ­nea ya insertada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea agregada:', licenciaId, 'sku:', sku, 'por:', usuario);
    if(licenciaCerradaMidFlight) {
      // La lÃ­nea quedÃ³ guardada en cdg_lineas, pero la licencia se cerrÃ³
      // mientras operabas. ok:true porque el dato estÃ¡ seguro.
      return res.status(409).json({
        ok: true,
        linea: lineaCreada,
        aviso: 'La licencia fue cerrada mientras agregabas la lÃ­nea. Tu lÃ­nea quedÃ³ guardada en el sistema.'
      });
    }
    res.json({ ok: true, linea: lineaCreada });
  } catch(e) {
    console.log('CDG v2 agregar linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo guardar la lÃ­nea. ReintentÃ¡. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId); // liberar escritura activa siempre
  }
});

// â”€â”€ PATCH /api/cdg/v2/:id/linea/:lineaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Edita una lÃ­nea propia (datos o fotos). Solo el autor puede editar su lÃ­nea.
// Body: { usuario, sku?, descripcion?, cantidad?, costoUnit?, fotos? }
app.patch('/api/cdg/v2/:id/linea/:lineaId', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineaId    = req.params.lineaId;
  var { usuario, sku, descripcion, cantidad, costoUnit, fotos } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se estÃ¡ cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(meta.estado === 'cerrado') {
      return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ cerrada. No se puede editar.' });
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
        return res.status(400).json({ ok: false, error: 'mÃ¡ximo 3 fotos por lÃ­nea' });
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
      return res.status(403).json({ ok: false, error: 'No se encontrÃ³ la lÃ­nea o no sos el autor.' });
    }

    // Actualizar tsUltimaModifLineas para que el prÃ³ximo poll la incluya.
    // FIX (sÃ¡b 30-may-2026, server v19): try separado â€” si la ediciÃ³n ya
    // persistiÃ³ y el bump falla, respondemos 200 igual (la lÃ­nea estÃ¡ guardada).
    try {
      if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
      var bumpEditResult = await withTimeout(cdgBumpLineasTs(licenciaId, meta, 0), 15000, 'CDG bump ts edit');
      if(bumpEditResult && bumpEditResult.estadoActual === 'cerrado') {
        console.log('CDG v2 linea editada (licencia cerrada mid-flight):', lineaId);
        return res.status(409).json({ ok: true, linea: lineaActualizada, aviso: 'La licencia fue cerrada mientras editabas. Tu cambio quedÃ³ guardado.' });
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata (edit) FAILED (lÃ­nea ya editada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea editada:', lineaId, 'por:', usuario);
    res.json({ ok: true, linea: lineaActualizada });
  } catch(e) {
    console.log('CDG v2 editar linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo editar la lÃ­nea. ReintentÃ¡. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// â”€â”€ DELETE /api/cdg/v2/:id/linea/:lineaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Soft-delete de una lÃ­nea. Solo el autor puede borrar su propia lÃ­nea.
// Body: { usuario }
app.delete('/api/cdg/v2/:id/linea/:lineaId', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var lineaId    = req.params.lineaId;
  var { usuario } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se estÃ¡ cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(meta.estado === 'cerrado') {
      return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ cerrada. No se puede eliminar.' });
    }

    var resultado = await withTimeout(
      cdgSoftDeleteLinea(lineaId, usuario),
      15000,
      'CDG delete linea'
    );

    if(!resultado) {
      return res.status(403).json({ ok: false, error: 'No se encontrÃ³ la lÃ­nea o no sos el autor.' });
    }

    // Actualizar metadata.
    // FIX (sÃ¡b 30-may-2026, server v19): try separado â€” si el soft-delete ya
    // persistiÃ³ y el bump falla, respondemos 200 igual (la lÃ­nea ya estÃ¡ eliminada).
    try {
      meta.totalLineas = Math.max(0, (meta.totalLineas || 1) - 1);
      if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();
      var bumpDelResult = await withTimeout(cdgBumpLineasTs(licenciaId, meta, -1), 15000, 'CDG bump ts delete');
      if(bumpDelResult && bumpDelResult.estadoActual === 'cerrado') {
        console.log('CDG v2 linea eliminada (licencia cerrada mid-flight):', lineaId);
        return res.status(409).json({ ok: true, aviso: 'La licencia fue cerrada mientras eliminabas. Tu cambio quedÃ³ guardado.' });
      }
    } catch(metaErr) {
      console.log('CDG v2 bump metadata (delete) FAILED (lÃ­nea ya eliminada, no bloqueante):', metaErr.message);
    }

    console.log('CDG v2 linea eliminada:', lineaId, 'por:', usuario);
    res.json({ ok: true });
  } catch(e) {
    console.log('CDG v2 delete linea FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudo eliminar la lÃ­nea. ReintentÃ¡. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// â”€â”€ POST /api/cdg/v2/:id/fotos-encabezado â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Agrega fotos de encabezado a la licencia (mÃ¡x 5).
// Body: { usuario, fotos[] } â€” fotos son paths en Storage ya subidos
app.post('/api/cdg/v2/:id/fotos-encabezado', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, fotos } = req.body;

  if(!usuario)               return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!Array.isArray(fotos) || fotos.length === 0) {
    return res.status(400).json({ ok: false, error: 'falta fotos' });
  }
  // Lock: registrar write de metadata. Si cerrando â†’ 423.
  if(!cdgLockAcquire(licenciaId)) {
    return res.status(423).json({ ok: false, error: 'La licencia se estÃ¡ cerrando.' });
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });
    if(!cdgAceptaLineas(meta)) {
      return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ ' + meta.estado });
    }

    var actuales = meta.fotosEncabezado || [];
    var total    = actuales.length + fotos.length;
    if(total > 5) {
      return res.status(400).json({
        ok: false,
        error: 'MÃ¡ximo 5 fotos de encabezado. Ya tenÃ©s ' + actuales.length + ', intentÃ¡s agregar ' + fotos.length + '.'
      });
    }

    meta.fotosEncabezado = actuales.concat(fotos);
    meta.version = (meta.version || 0) + 1;
    if(meta.usuarios[usuario]) meta.usuarios[usuario].lastActivity = new Date().toISOString();

    // FIX (lun 1-jun-2026, v19): re-leer meta fresco antes de guardar para
    // detectar cierre concurrente â€” mismo patrÃ³n que /accion bloque compartido.
    var metaFreshFotos = await cdgGetMeta(licenciaId);
    if(metaFreshFotos && metaFreshFotos.estado === 'cerrado') {
      return res.status(409).json({ ok: false, error: 'La licencia fue cerrada mientras subÃ­as la foto. No se guardÃ³.' });
    }
    await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG fotos encabezado save');
    console.log('CDG v2 fotos encabezado:', licenciaId, '+' + fotos.length + ' fotos');
    res.json({ ok: true, fotosEncabezado: meta.fotosEncabezado });
  } catch(e) {
    console.log('CDG v2 fotos encabezado FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'No se pudieron guardar las fotos. ReintentÃ¡. (' + e.message + ')' });
  } finally {
    cdgLockRelease(licenciaId);
  }
});

// â”€â”€ POST /api/cdg/v2/:id/accion â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Centraliza todas las operaciones de estado de licencia y usuarios.
// Body: { usuario, tipo, ...params }
//
// Tipos soportados:
//   unirse          â†’ usuario se une a la licencia
//   guardar_progreso â†’ registra lastActivity del usuario (sin cerrar)
//   pausar          â†’ usuario pasa a estado pausado
//   reanudar        â†’ usuario vuelve a activo
//   marcar_inactivo â†’ cliente informa que el usuario llegÃ³ a 45 min sin actividad
//   delegar         â†’ finalizador delega el rol a otro usuario activo
//                     params: { nuevoFinalizador }
//   cerrar          â†’ finalizador cierra la licencia definitivamente
//   desbloquear     â†’ supervisor reabre una licencia cerrada
//                     params: { supervisor: true }
app.post('/api/cdg/v2/:id/accion', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  var { usuario, tipo, nuevoFinalizador, supervisor } = req.body;

  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!tipo)    return res.status(400).json({ ok: false, error: 'falta tipo de acciÃ³n' });

  // Lock para acciones que escriben metadata (no-cerrar).
  // 'cerrar' gestiona el lock internamente con cdgLockStartClosing/WaitDrain.
  var lockAdquirido = false;
  if(tipo !== 'cerrar') {
    if(!cdgLockAcquire(licenciaId)) {
      return res.status(423).json({ ok: false, error: 'La licencia se estÃ¡ cerrando.' });
    }
    lockAdquirido = true;
  }

  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia no encontrada' });

    var now = new Date().toISOString();

    // â”€â”€ unirse â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    if(tipo === 'unirse') {
      if(meta.estado === 'cerrado') {
        return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ cerrada.' });
      }
      if(!meta.usuarios[usuario]) {
        meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
        cdgBitacora(meta, usuario, 'usuario_unido', '');
      } else {
        // Ya estaba registrado â€” reactivar
        meta.usuarios[usuario].estado       = 'activo';
        meta.usuarios[usuario].lastActivity = now;
        cdgBitacora(meta, usuario, 'usuario_reingreso', '');
      }
    }

    // â”€â”€ guardar_progreso â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'guardar_progreso') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].lastActivity = now;
      meta.usuarios[usuario].estado       = 'activo';
      cdgBitacora(meta, usuario, 'progreso_guardado', '');
    }

    // â”€â”€ pausar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'pausar') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado       = 'pausado';
      meta.usuarios[usuario].lastActivity = now;
      cdgBitacora(meta, usuario, 'usuario_pausado', '');
    }

    // â”€â”€ reanudar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'reanudar') {
      if(meta.estado === 'cerrado') {
        return res.status(403).json({ ok: false, error: 'La licencia estÃ¡ cerrada.' });
      }
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado       = 'activo';
      meta.usuarios[usuario].lastActivity = now;
      cdgBitacora(meta, usuario, 'usuario_reanudo', '');
    }

    // â”€â”€ marcar_inactivo â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'marcar_inactivo') {
      if(!meta.usuarios[usuario]) meta.usuarios[usuario] = { estado: 'activo', lastActivity: now };
      meta.usuarios[usuario].estado = 'inactivo';
      cdgBitacora(meta, usuario, 'usuario_inactivo_auto', '45 min sin actividad');

      // ReasignaciÃ³n automÃ¡tica lazy del finalizador si quedÃ³ inactivo
      if(meta.finalizador === usuario) {
        var candidatos = Object.keys(meta.usuarios).filter(function(u) {
          return u !== usuario && meta.usuarios[u].estado === 'activo';
        });
        if(candidatos.length > 0) {
          // Criterio: usuario con actividad mÃ¡s reciente
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

    // â”€â”€ delegar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'delegar') {
      if(meta.finalizador !== usuario) {
        return res.status(403).json({ ok: false, error: 'Solo el finalizador actual puede delegar.' });
      }
      if(!nuevoFinalizador) {
        return res.status(400).json({ ok: false, error: 'falta nuevoFinalizador' });
      }
      if(!meta.usuarios[nuevoFinalizador]) {
        return res.status(400).json({ ok: false, error: nuevoFinalizador + ' no estÃ¡ en esta licencia.' });
      }
      meta.finalizador = nuevoFinalizador;
      cdgBitacora(meta, usuario, 'finalizador_delegado',
        'De ' + usuario + ' a ' + nuevoFinalizador);
    }

    // â”€â”€ cerrar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'cerrar') {
      if(meta.estado === 'cerrado') {
        return res.status(400).json({ ok: false, error: 'La licencia ya estÃ¡ cerrada.' });
      }
      // FIX (lun 1-jun-2026, v19): permitir cierre a cualquier usuario activo
      // unido a la licencia, no solo al finalizador.
      // RazÃ³n operativa: en multiusuario el finalizador puede no estar disponible.
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
      // FIX (lun 1-jun-2026, v19): orden correcto â€” lock primero, re-leer fresco,
      // luego aplicar TODOS los campos de cierre al meta fresco para no perderlos.
      // 1. closing=true en memoria â†’ nuevos writes reciben 423 de inmediato
      // 2. drain â†’ esperar writes activos (hasta 15s; lanza si no drena)
      // 3. re-leer meta fresco â†’ tiene fotos/bitÃ¡cora de todos los writes anteriores
      // 4. aplicar TODOS los campos de cierre al fresco
      // 5. guardar en Supabase â†’ lock distribuido
      // 6. leer lÃ­neas con garantÃ­a
      cdgLockStartClosing(licenciaId);
      await cdgLockWaitDrain(licenciaId, 15000);

      var metaFresco = await cdgGetMeta(licenciaId);
      var base = metaFresco || meta;

      // Aplicar TODOS los campos de cierre sobre el meta mÃ¡s reciente
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
      // Persistir de forma directa y estricta. Si falla â†’ revertir y lanzar
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
      console.log('CDG v2 cerrar: entrada Hamilton creada para', licenciaId, '-', itemsV2.length, 'lÃ­neas');
      cdgLockClear(licenciaId); // limpiar lock â€” licencia cerrada, ya no necesita rastreo
    }

    // â”€â”€ desbloquear â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    else if(tipo === 'desbloquear') {
      if(!supervisor) {
        return res.status(403).json({ ok: false, error: 'Solo un supervisor puede desbloquear.' });
      }
      if(meta.estado !== 'cerrado') {
        return res.status(400).json({ ok: false, error: 'La licencia no estÃ¡ cerrada.' });
      }
      meta.estado      = 'activo';
      meta.fechaCierre = null;
      cdgBitacora(meta, usuario, 'licencia_desbloqueada', 'supervisor');
    }

    else {
      return res.status(400).json({ ok: false, error: 'Tipo de acciÃ³n desconocido: ' + tipo });
    }

    // Para 'cerrar': meta ya fue guardado con lock early (meta.version ya incrementado).
    // El bloque compartido harÃ­a un segundo increment y save redundante â€” lo saltamos.
    if(tipo !== 'cerrar') {
      // FIX (lun 1-jun-2026, v19): re-leer meta fresco antes de guardar.
      // Escenario sin este fix: guardar_progreso lee meta (activo), cierre guarda
      // meta cerrado + crea Hamilton, guardar_progreso termina y pisa con meta stale
      // â†’ licencia reabierta en Supabase con Hamilton ya creado.
      // EXCEPCIÃ“N: 'desbloquear' tiene que guardar estado='activo' sobre una licencia
      // cerrada â€” es exactamente su propÃ³sito, no debe ser bloqueado por esta guardia.
      var metaFresh = await cdgGetMeta(licenciaId);
      if(metaFresh && metaFresh.estado === 'cerrado' && tipo !== 'desbloquear') {
        return res.status(409).json({ ok: false, error: 'La licencia fue cerrada mientras procesabas esta acciÃ³n. No se guardaron cambios.' });
      }
      meta.version = (meta.version || 0) + 1;
      await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG accion save');
    }
    console.log('CDG v2 accion:', tipo, licenciaId, 'por:', usuario);
    res.json({ ok: true, meta: meta });

  } catch(e) {
    console.log('CDG v2 accion FAILED:', tipo, licenciaId, e.message);
    if(tipo === 'cerrar') cdgLockClear(licenciaId);
    res.status(500).json({ ok: false, error: 'No se pudo ejecutar la acciÃ³n. ReintentÃ¡. (' + e.message + ')' });
  } finally {
    if(lockAdquirido) cdgLockRelease(licenciaId);
  }
});

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// FIN CDG v2 â€” Los endpoints /api/cdg/* originales (v1 legacy) se conservan
// sin modificaciÃ³n arriba de este bloque para rollback si fuera necesario.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â”€â”€ Start â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const PORT = process.env.PORT || 3000;
// Always start server â€” even if Supabase fails
app.listen(PORT, () => console.log('Conteo app on port ' + PORT));
// Load state after server is up
loadState().catch(e => console.log('State load failed (non-fatal):', e.message));

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// CDG v2 â€” MANIFIESTO / VALIDACIÃ“N WMS
// FIX (lun 1-jun-2026, server v20): endpoints WMS para comparaciÃ³n CDG vs WMS.
// Tabla: cdg_wms (una fila por licencia, UPSERT al recargar).
// Routing: declarados ANTES de /:id para evitar captura por Express.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â”€â”€ Helpers WMS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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


// Helper: normalizar valor de columna NÃºmero del WMS.
// El Excel puede traer "U25-161959 Â· CDG" o "U25-161959 - CDG".
// Se extrae solo la parte izquierda del separador para comparar con licenciaId.
function cdgNormLicWMS(v) {
  var s = String(v || '').trim().toUpperCase();
  // Separador "Â·" (U+00B7, punto medio) con o sin espacios
  var dotIdx = s.indexOf('Â·');
  if(dotIdx >= 0) s = s.substring(0, dotIdx);
  // Separador " - " con espacios
  var dashIdx = s.indexOf(' - ');
  if(dashIdx >= 0) s = s.substring(0, dashIdx);
  return s.trim();
}

// Parseo robusto de cantidad: soporta 1234, "1,234", "1,234.00", espacios.
// Retorna null si no es numÃ©rico o <= 0.
function cdgParseQty(val) {
  if(val === null || val === undefined || val === '') return null;
  var s = String(val).trim().replace(/,/g, '');
  var n = Number(s);
  if(isNaN(n) || n <= 0) return null;
  return n;
}

// Normaliza un encabezado de columna para detecciÃ³n flexible:
// "Cant." â†’ "cant", "# Ingreso" â†’ "ingreso", "NÃºmero" â†’ "numero"
function cdgNormHdr(h) {
  return String(h || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quitar tildes
    .replace(/[^a-z0-9]/g, '');                        // solo alfanumÃ©rico
}

// FIX (mar 2-jun-2026, server v20): rutas fijas /bulk y /sincronizar-hamilton
// DEBEN ir ANTES de la ruta paramÃ©trica /:id. Express hace match en orden.
// Si /:id va primero, captura "bulk" como parÃ¡metro :id.

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// CDG WMS BULK â€” POST /api/cdg/wms/bulk
// FIX (lun 1-jun-2026, server v20): carga masiva de WMS para licencias antiguas.
// Recibe el mismo Excel WMS (hoja Traslados) y procesa TODAS las licencias.
// No requiere que la licencia exista como cdg_meta_* v2; sirve para CDG clÃ¡sico.
// No crea conteos, no modifica cdg_lineas, no toca Hamilton.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
app.post('/api/cdg/wms/bulk', upload.single('file'), async (req, res) => {
  var usuario = String(req.body && req.body.usuario ? req.body.usuario : '').trim();
  if(!usuario)  return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    // â”€â”€ 1. Parsear Excel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: 'buffer', raw: false, cellDates: true });
    } catch(e) {
      return res.status(400).json({ ok: false, error: 'El archivo no es un Excel vÃ¡lido. (' + e.message + ')' });
    }

    var advertencias = [];
    var sheetName = wb.SheetNames.find(function(n){ return n.trim().toLowerCase() === 'traslados'; });
    if(!sheetName) {
      sheetName = wb.SheetNames[0];
      advertencias.push('Hoja "Traslados" no encontrada; se usÃ³ la primera hoja: "' + sheetName + '".');
    }
    var ws   = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if(rows.length < 2) return res.status(400).json({ ok: false, error: 'El archivo estÃ¡ vacÃ­o o solo tiene encabezado.' });

    // â”€â”€ 2. Detectar columnas (reutiliza la misma lÃ³gica del endpoint individual) â”€
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
    if(cNum  < 0) faltantes.push('"NÃºmero"');
    if(cSku  < 0) faltantes.push('"SKU"');
    if(cCant < 0 && cUnis < 0) faltantes.push('"Cant." o "Unidades"');
    if(faltantes.length) {
      return res.status(400).json({
        ok: false,
        error: 'Columnas requeridas no encontradas: ' + faltantes.join(', ') + '. Encabezados: ' + rows[0].join(', ')
      });
    }

    // â”€â”€ 3. Agrupar filas por licencia â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var porLicencia = {};  // licNorm â†’ { skuMap:{}, encabezado, filaCount }
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
      return res.status(400).json({ ok: false, error: 'No se encontraron filas vÃ¡lidas en el archivo.' });
    }

    // â”€â”€ 4. Enriquecer descripciones vacÃ­as con sku_catalog (batch por licencia) â”€â”€
    // Recopilar todos los SKUs sin descripciÃ³n en un solo set
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

    // â”€â”€ 5. Upsert por licencia â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

      // Verificar si existe como v2 (meta) â€” solo informativo, no bloquea
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

    console.log('CDG WMS bulk:', usuario, 'â€”', licenciasActualizadas.length, 'licencias,', totalSkus, 'SKUs');
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// CDG WMS SINCRONIZAR HAMILTON â€” POST /api/cdg/wms/sincronizar-hamilton
// FIX (mar 2-jun-2026, server v20 rev2):
//   - Consolida por SKU antes de cruzar con WMS (fix bug lÃ­neas repetidas).
//     CDG clÃ¡sico puede tener N filas del mismo SKU; Hamilton necesita una sola.
//   - Fuente de consolidaciÃ³n: (1) cdg_lineas v2, (2) state.cdg[lic].items v1,
//     (3) teo.items existente como fallback.
//   - Remapea fisico existente al nuevo alineamiento por SKU consolidado.
//   - snapshot/rollback de teorico+fisico+version+historial.
//   - Busca por licId directo y por fallback cdgRef.
//   - Recalcula siempre con el WMS vigente.
//   - No crea conteos nuevos. No toca cdg_lineas. No toca Hamilton general.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
app.post('/api/cdg/wms/sincronizar-hamilton', async (req, res) => {
  var usuario = String((req.body && req.body.usuario) || '').trim();
  if(!usuario) return res.status(400).json({ ok: false, error: 'falta usuario' });

  try {
    // FIX (mar 2-jun-2026, server v20 rev): procesar cdg_wms pÃ¡gina por pÃ¡gina.
    // NO acumular en wmsAllRows â€” cada pÃ¡gina se procesa y descarta para no
    // retener skus jsonb completos de 200 licencias en memoria simultÃ¡neamente.
    var PAGE = 200;
    var totalProcesadas = 0;
    var huboWms = false;

    // â”€â”€ Snapshot ANTES de mutar â€” para rollback si el save falla â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var snapTeorico   = state.teorico ? JSON.parse(JSON.stringify(state.teorico)) : {};
    var snapFisico    = state.fisico   ? JSON.parse(JSON.stringify(state.fisico))  : {};
    var snapVersion   = state.version;
    var snapHistorial = state.historial ? state.historial.length : 0;

    var actualizadas   = [];
    var sinConteo      = [];
    var advertencias   = [];

    for(var pg = 0; pg < 10; pg++) {   // mÃ¡x 10 pÃ¡ginas = 2000 licencias
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

      // â”€â”€ Capturar items y fisico ANTES de cualquier mutaciÃ³n â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
      // CRÃTICO: viejoItems debe leerse ANTES de modificar teo.items.
      // Si se lee despuÃ©s de la consolidaciÃ³n, ya no tiene las N filas originales
      // del mismo SKU y el remap de fisico queda incorrecto (suma solo la primera fila).
      // FIX (mar 2-jun-2026, server v20 rev3): mover captura aquÃ­, antes del consolidado.
      var viejoFisico = state.fisico && state.fisico[teoKey]
        ? (Array.isArray(state.fisico[teoKey]) ? state.fisico[teoKey].slice() : null)
        : null;
      var viejoItems = (Array.isArray(teo.items) ? teo.items : [])
        .filter(function(it){ return !it._soloWMS; });
      try {
        skusWMS = typeof wmsRow.skus === 'string' ? JSON.parse(wmsRow.skus) : (wmsRow.skus || []);
      } catch(e) { skusWMS = []; }
      if(!skusWMS.length) continue;

      // Mapa WMS: SKU_UPPER â†’ { cantidad, descripcion }
      var wmsMap = {};
      var wmsDescMap = {};
      skusWMS.forEach(function(s){
        var sk = String(s.sku || '').trim().toUpperCase();
        wmsMap[sk] = (wmsMap[sk] || 0) + (Number(s.cantidad) || 0);
        if(!wmsDescMap[sk] && s.descripcion) wmsDescMap[sk] = s.descripcion;
      });

      // â”€â”€ Construir base consolidada por SKU â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
      // Prioridad:
      //   1. CDG v2: leer cdg_lineas vigentes y agrupar por SKU.
      //   2. CDG v1: leer state.cdg[lic].items y agrupar por SKU.
      //   3. Fallback: items existentes en teo (tambiÃ©n agrupar â€” puede tener dupes).
      // En todos los casos: UNA fila por SKU en la base consolidada.
      // { skUpper: { sku, desc, qty, costo, orden } }
      var cdgConsolidado = {};  // SK_UPPER â†’ { sku, desc, qty, costo, orden }
      var ordenSku = [];        // orden de primera apariciÃ³n

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

      // Fallback 3: items existentes en teo (sin _soloWMS), tambiÃ©n consolidando
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

      // â”€â”€ Remap fisico existente al nuevo Ã­ndice consolidado por SKU â”€â”€â”€â”€â”€â”€â”€â”€
      // fisMapeado usa viejoItems (capturado antes de mutar) y viejoFisico.
      // fisMapeado: SK_UPPER â†’ { fisico, daniado, quien, ts, cobertura }
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
            // Conservar Ãºltimo quien/ts no vacÃ­o
            if(sv.quien) fisMapeado[sk].quien = sv.quien;
            if(sv.ts)    fisMapeado[sk].ts    = sv.ts;
          } else {
            fisMapeado[sk] = {
              fisico:   fv,
              daniado:  dv,
              quien:    sv.quien    || '',
              ts:       sv.ts       || '',
              cobertura: sv.cobertura || 'En revisiÃ³n',
              calcExpr: sv.calcExpr || null
            };
          }
        });
      }

      // â”€â”€ Cruzar consolidado con WMS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

      // â”€â”€ Actualizar teorico â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
      teo.items = nuevosItems;

      // â”€â”€ Actualizar fisico si hay remap disponible â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
      var tieneFisicoRemap = Object.keys(fisMapeado).length > 0;
      if(tieneFisicoRemap) {
        // Construir nuevo array fisico alineado a nuevosItems (excluye _soloWMS)
        var nuevoFisico = nuevosItems.map(function(item){
          if(item._soloWMS) return null; // SKU solo-WMS: sin fÃ­sico
          var sk = String(item.sku || '').trim().toUpperCase();
          var fm = fisMapeado[sk];
          if(!fm) return null; // sin fÃ­sico registrado para este SKU
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

      actualizadas.push(lic + (teoKey !== lic ? 'â†’'+teoKey : '')
        + ' (' + nuevosItems.filter(function(i){ return !i._soloWMS; }).length + ' SKUs consolidados)');
      } // fin for wi (pageRows)

      if(pageRows.length < PAGE) break; // Ãºltima pÃ¡gina
    } // fin for pg (pÃ¡ginas)

    if(!huboWms) {
      return res.json({ ok: true, actualizadas: [], sinConteo: [], msg: 'No hay WMS cargados.' });
    }

    // â”€â”€ Persistir â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
          error: 'No se guardÃ³. Los datos en memoria fueron restaurados. ReintentÃ¡. (' + saveErr.message + ')'
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

// â”€â”€ cdgCargarWmsDesdeExcel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Helper que encapsula todo el parseo y guardado del WMS desde Excel.
// FIX (mar 2-jun-2026, server v20): extraÃ­do para reutilizar en el endpoint
// general POST /api/cdg/wms/:id (sin cdg_meta) y en POST /api/cdg/v2/:id/wms.
// ParÃ¡metros:
//   licenciaId: ya normalizado con cdgNormId
//   req: request de Express (req.file, req.body.usuario)
//   meta: cdg_meta ya cargado (o null si no existe)
//   res: response de Express
async function cdgCargarWmsDesdeExcel(licenciaId, req, meta, res) {
  var usuario = String(req.body && req.body.usuario ? req.body.usuario : '').trim();
  if(!usuario)  return res.status(400).json({ ok: false, error: 'falta usuario' });
  if(!req.file) return res.status(400).json({ ok: false, error: 'falta archivo' });

  try {
    // â”€â”€ Parsear Excel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: 'buffer', raw: false, cellDates: true });
    } catch(e) {
      return res.status(400).json({ ok: false, error: 'El archivo no es un Excel vÃ¡lido. (' + e.message + ')' });
    }

    var advertencia = null;
    var sheetName = wb.SheetNames.find(function(n) { return n.trim().toLowerCase() === 'traslados'; });
    if(!sheetName) {
      sheetName = wb.SheetNames[0];
      advertencia = 'Hoja "Traslados" no encontrada. Se usÃ³ la primera hoja: "' + sheetName + '". VerificÃ¡ que estÃ¡s cargando el archivo WMS correcto.';
    }
    var ws   = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    if(rows.length < 2) return res.status(400).json({ ok: false, error: 'El archivo estÃ¡ vacÃ­o o solo tiene encabezado.' });

    // â”€â”€ Detectar columnas â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
    if(cNum  < 0) faltantes.push('"NÃºmero"');
    if(cSku  < 0) faltantes.push('"SKU"');
    if(cCant < 0 && cUnis < 0) faltantes.push('"Cant." o "Unidades"');
    if(faltantes.length > 0) {
      return res.status(400).json({
        ok: false,
        error: 'Columnas requeridas no encontradas: ' + faltantes.join(', ') + '. '
             + 'Encabezados detectados: ' + rows[0].join(', ') + '. '
             + 'VerificÃ¡ que estÃ¡s cargando el archivo WMS correcto (hoja "Traslados" del export WMS).'
      });
    }

    // â”€â”€ Filtrar por licencia â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
             + (licEncontradas.length === 10 ? ' (y mÃ¡sâ€¦)' : '') + '.'
      });
    }

    // â”€â”€ Extraer encabezado â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var f0 = filasFiltradas[0];
    var encabezado = {
      fecha:     f0[cFecha]  || null,
      origen:    cOrigen  >= 0 ? (f0[cOrigen]  || null) : null,
      destino:   cDest    >= 0 ? (f0[cDest]    || null) : null,
      tipo:      cTipo    >= 0 ? (f0[cTipo]    || null) : null,
      status:    cStatus  >= 0 ? (f0[cStatus]  || null) : null,
      lineasWMS: cLineas  >= 0 ? Number(f0[cLineas] || 0) : 0
    };

    // â”€â”€ Procesar filas: parsear + deduplicar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // â”€â”€ Enriquecer descripciones vacÃ­as desde sku_catalog â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // â”€â”€ Construir array final + UPSERT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // â”€â”€ Actualizar meta.wms si existe cdg_meta â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    // Si meta es null (v1 o solo-WMS), no intentar cdgSaveMeta.
    var metaWarning = null;
    if(meta) {
      meta.wms = { cargadoPor: usuario, tsCarga: now, nombreArchivo, totalSkus, totalUnidades };
      cdgBitacora(meta, usuario, 'wms_cargado', nombreArchivo + ' â€” ' + totalSkus + ' SKUs, ' + totalUnidades + ' unidades');
      try {
        await withTimeout(cdgSaveMeta(licenciaId, meta), 15000, 'CDG WMS meta save');
      } catch(metaErr) {
        try {
          await withTimeout(cdgSaveMeta(licenciaId, meta), 10000, 'CDG WMS meta save retry');
        } catch(retryErr) {
          metaWarning = 'WMS guardado. No se pudo actualizar cachÃ© de metadata (' + retryErr.message + '). El WMS estarÃ¡ disponible igual.';
          console.log('CDG WMS meta save FAILED after retry:', retryErr.message);
        }
      }
    }

    var respAdvertencia = [advertencia, metaWarning].filter(Boolean).join(' | ') || undefined;
    console.log('CDG WMS cargado:', licenciaId, 'por', usuario, 'â€”', totalSkus, 'SKUs,', totalUnidades, 'unidades');
    res.json({
      ok: true, totalSkus, totalUnidades, tsCarga: now,
      ...(respAdvertencia ? { advertencia: respAdvertencia } : {})
    });

  } catch(e) {
    console.log('CDG WMS upload FAILED:', e.message);
    res.status(500).json({ ok: false, error: 'Error al procesar WMS: ' + e.message });
  }
}

// â”€â”€ POST /api/cdg/wms/:id â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Endpoint general WMS: NO exige cdg_meta v2.
// FIX (mar 2-jun-2026, server v20): permite cargar WMS para v1, v2 y solo-WMS.
// Si existe cdg_meta y estado === 'cerrado' â†’ rechaza.
// Si existe cdg_meta abierta â†’ guarda WMS + actualiza meta.wms.
// Si NO existe cdg_meta â†’ guarda solo en cdg_wms (sin meta).
app.post('/api/cdg/wms/:id', upload.single('file'), async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    var meta = null;
    try { meta = await cdgGetMeta(licenciaId); } catch(e) { /* no existe como v2 */ }
    if(meta && meta.estado === 'cerrado') {
      return res.status(400).json({ ok: false, error: 'La licencia estÃ¡ cerrada. No se puede cargar WMS.' });
    }
    await cdgCargarWmsDesdeExcel(licenciaId, req, meta, res);
  } catch(e) {
    console.log('CDG WMS POST general FAILED:', e.message);
    if(!res.headersSent) res.status(500).json({ ok: false, error: e.message });
  }
});

// â”€â”€ POST /api/cdg/v2/:id/wms â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Mantiene compatibilidad con el cliente existente de captura v2.
// Ahora delega a cdgCargarWmsDesdeExcel (misma lÃ³gica, sin cÃ³digo duplicado).
// FIX (mar 2-jun-2026, server v20): ya no duplica la lÃ³gica de parseo.
app.post('/api/cdg/v2/:id/wms', upload.single('file'), async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    var meta = await cdgGetMeta(licenciaId);
    if(!meta) return res.status(404).json({ ok: false, error: 'Licencia v2 no encontrada: ' + licenciaId });
    if(meta.estado === 'cerrado') {
      return res.status(400).json({ ok: false, error: 'La licencia estÃ¡ cerrada. No se puede cargar WMS.' });
    }
    await cdgCargarWmsDesdeExcel(licenciaId, req, meta, res);
  } catch(e) {
    console.log('CDG v2 WMS upload FAILED:', e.message);
    if(!res.headersSent) res.status(500).json({ ok: false, error: e.message });
  }
});

// â”€â”€ cdgResponderWms: helper compartido para devolver WMS de cdg_wms â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€ GET /api/cdg/wms/:id â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Endpoint general WMS: no exige cdg_meta v2.
// Sirve para v1 clÃ¡sico, v2 colaborativo y licencias "solo WMS".
// FIX (mar 2-jun-2026, server v20): nuevo endpoint sin restricciÃ³n de v2.
app.get('/api/cdg/wms/:id', async (req, res) => {
  var licenciaId = cdgNormId(req.params.id);
  try {
    await cdgResponderWms(licenciaId, res);
  } catch(e) {
    console.log('CDG WMS GET FAILED:', licenciaId, e.message);
    res.status(500).json({ ok: false, error: 'Error al obtener WMS: ' + e.message });
  }
});

// â”€â”€ GET /api/cdg/v2/:id/wms â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// BOD MODULE START â€” Bodega CDG / Armado de Tarimas â€” Fase 1
// FIX (mar 2-jun-2026, server v20): mÃ³dulo completamente aislado.
// Flag: BOD_ENABLED=true en variables de entorno de Render para activar.
// Con BOD_ENABLED=false (default), todos los endpoints responden 503.
// No toca: Hamilton, CDG, app_state, cdg_lineas, cdg_wms, sku_catalog,
//          teorico, fisico, endpoints /api/cdg/*, /api/state.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

var BOD_ENABLED = process.env.BOD_ENABLED === 'true';

// Middleware guard â€” aplica a todos los endpoints /api/bod/*
function bodGuard(req, res, next) {
  if(!BOD_ENABLED) return res.status(503).json({ ok: false, error: 'MÃ³dulo bodega no habilitado.' });
  next();
}

// Helper: normalizar nombre de tarima â†’ "A1", "B12", etc.
// "a 1" â†’ "A1", "A-1" â†’ "A1", "a1" â†’ "A1"
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

// â”€â”€ GET /api/bod/status â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Permite al cliente saber si el mÃ³dulo estÃ¡ habilitado.
app.get('/api/bod/status', function(req, res) {
  res.json({ ok: true, enabled: BOD_ENABLED });
});

// â”€â”€ POST /api/bod/sesion â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Crear o devolver sesiÃ³n existente (idempotente por licencia_id+fecha+tipo).
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

// â”€â”€ GET /api/bod/sesion?licencia_id=X&fecha=Y â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Buscar sesiÃ³n por licencia + fecha en lugar del ID interno.
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
      return res.status(404).json({ ok:false, error:'No se encontrÃ³ sesiÃ³n para la licencia "'+lic+'" en la fecha '+fecha+'.' });
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

// â”€â”€ GET /api/bod/sesion/:id â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Devolver sesiÃ³n + lÃ­neas vigentes + resumen por tarima.
app.get('/api/bod/sesion/:id', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var rows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(rows) || !rows.length) return res.status(404).json({ ok:false, error:'SesiÃ³n no encontrada' });
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

// â”€â”€ POST /api/bod/sesion/:id/linea â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // Verificar sesiÃ³n existe y estÃ¡ abierta
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'SesiÃ³n no encontrada' });
    if(sesRows[0].estado === 'cerrada') return res.status(400).json({ ok:false, error:'La sesiÃ³n estÃ¡ cerrada.' });

    // Completar descripciÃ³n desde sku_catalog si vacÃ­a
    var advertencia = null;
    if(!descripcion && SUPABASE_URL && SUPABASE_KEY) {
      try {
        var catRows = await supabase('GET', 'sku_catalog', null,
          '?sku=eq.' + encodeURIComponent(sku) + '&select=sku,descripcion&limit=1');
        if(Array.isArray(catRows) && catRows.length && catRows[0].descripcion)
          descripcion = catRows[0].descripcion;
      } catch(e) { advertencia = 'No se pudo completar descripciÃ³n desde catÃ¡logo.'; }
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

// â”€â”€ PATCH /api/bod/linea/:lineaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.patch('/api/bod/linea/:lineaId', bodGuard, async (req, res) => {
  try {
    var linId   = String(req.params.lineaId).trim();
    var { cantidad, sku, descripcion, tarima, usuario } = req.body || {};
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'LÃ­nea no encontrada' });
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

// â”€â”€ DELETE /api/bod/linea/:lineaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.delete('/api/bod/linea/:lineaId', bodGuard, async (req, res) => {
  try {
    var linId   = String(req.params.lineaId).trim();
    var { usuario, supervisor } = req.body || {};
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'LÃ­nea no encontrada' });
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
    } catch(e) { /* columnas aÃºn no existen â€” ignorar */ }
    await supabase('PATCH', 'bod_lineas', patch,
      '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// â”€â”€ GET /api/bod/sesion/:id/lineas â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€ GET /api/bod/barra/:barra â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.get('/api/bod/barra/:barra', bodGuard, async (req, res) => {
  try {
    var barra = String(req.params.barra).trim();
    var row   = await bodGetBarra(barra);
    if(!row) return res.json({ ok:true, encontrado:false });
    // Completar descripciÃ³n desde sku_catalog si vacÃ­a
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

// â”€â”€ POST /api/bod/barra-sku/upload â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/bod/barra-sku/upload', bodGuard, upload.single('file'), async (req, res) => {
  try {
    if(!req.file) return res.status(400).json({ ok:false, error:'falta archivo' });

    // â”€â”€ Parsear CSV o XLSX con XLSX (maneja ambos formatos) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var wb = XLSX.read(req.file.buffer, { type:'buffer', raw:false });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });
    if(rows.length < 2) return res.status(400).json({ ok:false, error:'Archivo vacÃ­o o solo encabezado.' });

    // â”€â”€ Detectar columnas â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    var hdr   = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var cBarra = hdr.findIndex(function(h){ return h==='barra'||h.includes('barcode')||h.includes('codigo'); });
    var cSku   = hdr.findIndex(function(h){ return h==='sku'||h.includes('articulo')||h.includes('artÃ­culo'); });
    var cDesc  = hdr.findIndex(function(h){ return h.includes('desc')||h.includes('nombre'); });
    if(cBarra < 0 || cSku < 0) {
      return res.status(400).json({
        ok:false,
        error:'No se encontraron columnas Barra y SKU. Encabezados detectados: '+rows[0].join(', ')
      });
    }

    // â”€â”€ Deduplicar en memoria â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    // Reglas: ignorar barra vacÃ­a, sku vacÃ­o, barra con < 6 caracteres.
    // Si barra aparece duplicada, conservar la Ãºltima fila.
    var dedup = {};        // barra â†’ { barra, sku, descripcion }
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
        error:'No se encontraron filas vÃ¡lidas. RevisÃ¡ que las columnas barra y sku tengan datos.',
        omitidas: omitidas
      });
    }

    // â”€â”€ Upsert por chunks de 500 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
        console.log('BOD barra-sku chunk', chunks, 'fallÃ³:', e.message);
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

// â”€â”€ POST /api/bod/linea/:lineaId/auditar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/bod/linea/:lineaId/auditar', bodGuard, async (req, res) => {
  try {
    var linId = String(req.params.lineaId).trim();
    var { usuario, supervisor, auditor, cantidad_audit } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden auditar.' });
    if(cantidad_audit === undefined || cantidad_audit === null || !(Number(cantidad_audit) > 0))
      return res.status(400).json({ ok:false, error:'cantidad_audit debe ser mayor a 0.' });
    var rows = await supabase('GET', 'bod_lineas', null, '?id=eq.' + encodeURIComponent(linId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'LÃ­nea no encontrada' });
    var now = new Date().toISOString();
    var patch = { auditado:true, auditado_por:usuario, ts_auditado:now, ts_modif:now,
                  cantidad_audit: Number(cantidad_audit) };
    await supabase('PATCH', 'bod_lineas', patch, '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// â”€â”€ POST /api/bod/linea/:lineaId/unauditar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Revierte una lÃ­nea auditada a estado pendiente.
app.post('/api/bod/linea/:lineaId/unauditar', bodGuard, async (req, res) => {
  try {
    var linId = String(req.params.lineaId).trim();
    var { usuario, supervisor, auditor } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden anular auditorÃ­as.' });
    var now = new Date().toISOString();
    await supabase('PATCH', 'bod_lineas',
      { auditado:false, auditado_por:null, ts_auditado:null, cantidad_audit:null, ts_modif:now },
      '?id=eq.' + encodeURIComponent(linId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// â”€â”€ POST /api/bod/sesion/:id/cerrar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/bod/sesion/:id/cerrar', bodGuard, async (req, res) => {
  try {
    var sesId  = String(req.params.id).trim();
    var { usuario, supervisor, auditor } = req.body || {};
    var esAuditor = supervisor === true || auditor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden cerrar sesiones.' });
    var rows = await supabase('GET', 'bod_sesiones', null, '?id=eq.' + encodeURIComponent(sesId) + '&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'SesiÃ³n no encontrada' });
    await supabase('PATCH', 'bod_sesiones',
      { estado:'cerrada', ts_cierre:new Date().toISOString(), modificado_por:usuario },
      '?id=eq.' + encodeURIComponent(sesId));
    res.json({ ok:true });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// â”€â”€ GET /api/bod/sesion/:id/furgones â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Devuelve asignaciones bod_tarima_furgon + cargas logÃ­sticas bod_furgon_cierres
app.get('/api/bod/sesion/:id/furgones', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var [rows, cierres] = await Promise.all([
      supabase('GET', 'bod_tarima_furgon', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&order=furgon.asc,tarima.asc'),
      supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)).catch(function(){ return []; })
    ]);
    rows    = Array.isArray(rows)    ? rows    : [];
    cierres = Array.isArray(cierres) ? cierres : [];
    res.json({ ok:true, asignaciones: rows, cargas: cierres });
  } catch(e) {
    res.status(500).json({ ok:false, error: e.message });
  }
});

// â”€â”€ POST /api/bod/sesion/:id/carga-logistica â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Combina: asignar tarimas a furgÃ³n + crear carga logÃ­stica en bod_furgon_cierres.
// Reemplaza el flujo de 2 pasos (asignar + finalizar) por un Ãºnico paso.
app.post('/api/bod/sesion/:id/carga-logistica', bodGuard, async (req, res) => {
  try {
    var sesId = String(req.params.id).trim();
    var { tarimas, furgon, placa, marchamo, licencia_hija, destino_tr999,
          observacion, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden crear cargas logÃ­sticas.' });

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
    if(!Array.isArray(tarimas) || !tarimas.length)
      return res.status(400).json({ ok:false, error:'Debe seleccionar al menos una tarima.' });

    var tarimasNorm = tarimas.map(function(t){ return bodNormTarima(t); }).filter(Boolean);

    // Verificar sesiÃ³n
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'SesiÃ³n no encontrada' });
    var sesion = sesRows[0];
    if(sesion.estado === 'cerrada') return res.status(400).json({ ok:false, error:'La sesiÃ³n estÃ¡ cerrada.' });

    // Verificar que el furgÃ³n no tenga carga logÃ­stica ya en esta sesiÃ³n
    try {
      var existing = await supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon)+'&limit=1');
      if(Array.isArray(existing) && existing.length) {
        return res.status(409).json({ ok:false, error:'El furgÃ³n '+furgon+' ya tiene una carga logÃ­stica en esta sesiÃ³n.' });
      }
    } catch(e) { /* tabla puede no existir â€” continuar */ }

    // Verificar que las tarimas no estÃ©n ya en otro furgÃ³n diferente
    var asigExist = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId));
    asigExist = Array.isArray(asigExist) ? asigExist : [];
    var asigMap = {};
    asigExist.forEach(function(a){ asigMap[a.tarima] = a.furgon; });
    var bloqueadas = tarimasNorm.filter(function(t){ return asigMap[t] && asigMap[t] !== furgon; });
    if(bloqueadas.length)
      return res.status(409).json({ ok:false, error:'Tarimas ya en otro furgÃ³n: '+bloqueadas.join(', '), bloqueadas:bloqueadas });

    // 1. Asignar tarimas al furgÃ³n en bod_tarima_furgon
    var now = new Date().toISOString();
    var nuevas = tarimasNorm.filter(function(t){ return !asigMap[t]; });
    for(var i=0; i<nuevas.length; i++) {
      await supabase('POST', 'bod_tarima_furgon',
        { id:'btf-'+Date.now()+'-'+i+'-'+Math.random().toString(36).slice(2,5),
          sesion_id:sesId, tarima:nuevas[i], furgon:furgon,
          asignado_por:usuario, ts_asignacion:now },
        '');
    }

    // 2. Leer lÃ­neas de esas tarimas para resumen SKU
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
    var resumenSkus  = Object.values(skuMap);
    var totalUnidades = resumenSkus.reduce(function(s,x){ return s+x.unidades; }, 0);

    // 3. Crear carga logÃ­stica en bod_furgon_cierres
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
    await supabase('POST', 'bod_furgon_cierres', cierre, '');

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

// â”€â”€ POST /api/bod/sesion/:id/furgon â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.post('/api/bod/sesion/:id/furgon', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var { tarimas, furgon, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden asignar furgones.' });
    if(!furgon)    return res.status(400).json({ ok:false, error:'falta furgon' });
    if(!Array.isArray(tarimas)||!tarimas.length) return res.status(400).json({ ok:false, error:'tarimas debe ser array no vacÃ­o' });
    var normadas = tarimas.map(function(t){ return bodNormTarima(t); }).filter(Boolean);
    if(!normadas.length) return res.status(400).json({ ok:false, error:'ninguna tarima vÃ¡lida' });
    furgon = String(furgon).trim();

    // Leer asignaciones existentes de la sesiÃ³n
    var existing = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId));
    var asigMap = {};
    if(Array.isArray(existing)) existing.forEach(function(r){ asigMap[r.tarima] = r.furgon; });

    // Detectar bloqueadas (asignadas a otro furgÃ³n distinto)
    var bloqueadas = normadas.filter(function(t){ return asigMap[t] && asigMap[t] !== furgon; });
    if(bloqueadas.length) return res.status(409).json({ ok:false, error:'Algunas tarimas ya estÃ¡n asignadas.', bloqueadas: bloqueadas });

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

// â”€â”€ DELETE /api/bod/sesion/:id/furgon/:tarima â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€ GET /api/bod/sesion/:id/manifiesto?furgon=N â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // Mapa tarima â†’ furgon y cierre â†’ datos
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

// â”€â”€ POST /api/bod/sesion/:id/furgon/:furgon/finalizar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Registra el cierre de un furgÃ³n: licencia hija + destino TR999.
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

    // Verificar sesiÃ³n existe
    var sesRows = await supabase('GET', 'bod_sesiones', null, '?id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(sesRows)||!sesRows.length) return res.status(404).json({ ok:false, error:'SesiÃ³n no encontrada' });
    var sesion = sesRows[0];

    // Verificar que no estÃ© ya finalizado
    try {
      var existing = await supabase('GET', 'bod_furgon_cierres', null,
        '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon)+'&limit=1');
      if(Array.isArray(existing) && existing.length) {
        return res.status(409).json({ ok:false, error:'El furgÃ³n '+furgon+' ya fue finalizado con licencia hija '+existing[0].licencia_hija+'.' });
      }
    } catch(e) { /* tabla puede no existir aÃºn â€” continuar */ }

    // Leer tarimas y lÃ­neas
    var asig = await supabase('GET', 'bod_tarima_furgon', null,
      '?sesion_id=eq.'+encodeURIComponent(sesId)+'&furgon=eq.'+encodeURIComponent(furgon));
    asig = Array.isArray(asig) ? asig : [];
    if(!asig.length) return res.status(400).json({ ok:false, error:'El furgÃ³n '+furgon+' no tiene tarimas asignadas.' });

    var tarimas = asig.map(function(a){ return a.tarima; });
    var tarimasQ = '?sesion_id=eq.'+encodeURIComponent(sesId)
      +'&tarima=in.('+tarimas.map(encodeURIComponent).join(',')+')'
      +'&eliminada=eq.false';
    var lineas = await supabase('GET', 'bod_lineas', null, tarimasQ);
    lineas = Array.isArray(lineas) ? lineas : [];
    if(!lineas.length) return res.status(400).json({ ok:false, error:'El furgÃ³n '+furgon+' no tiene lÃ­neas capturadas.' });

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

// â”€â”€ PATCH /api/bod/sesion/:id/carga-logistica/:cargaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.patch('/api/bod/sesion/:id/carga-logistica/:cargaId', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { placa, marchamo, licencia_hija, destino_tr999, observacion, usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden editar cargas logÃ­sticas.' });

    destino_tr999 = String(destino_tr999 || '').trim().toUpperCase();
    licencia_hija = String(licencia_hija || '').trim().toUpperCase();
    if(!licencia_hija) return res.status(400).json({ ok:false, error:'falta licencia_hija' });
    if(!destino_tr999) return res.status(400).json({ ok:false, error:'falta destino_tr999' });
    if(!/^TR999\.\d{3}\.\d{2}$/i.test(destino_tr999))
      return res.status(400).json({ ok:false, error:'El destino debe tener formato TR999.xxx.xx (ej: TR999.001.01). Recibido: '+destino_tr999 });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Carga no encontrada en esta sesiÃ³n.' });
    if(rows[0].estado === 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga estÃ¡ cerrada. Solo un supervisor puede reabrirla.' });

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

// â”€â”€ DELETE /api/bod/sesion/:id/carga-logistica/:cargaId â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
app.delete('/api/bod/sesion/:id/carga-logistica/:cargaId', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores pueden eliminar cargas logÃ­sticas.' });

    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length) return res.status(404).json({ ok:false, error:'Carga no encontrada en esta sesiÃ³n.' });

    var carga  = rows[0];
    if(carga.estado === 'cerrada')
      return res.status(409).json({ ok:false, error:'La carga estÃ¡ cerrada. Solo un supervisor puede reabrirla antes de eliminar.' });
    var tarimas = Array.isArray(carga.tarimas) ? carga.tarimas : [];

    // Eliminar la carga logÃ­stica
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

// â”€â”€ PATCH /api/bod/sesion/:id/carga-logistica/:cargaId/cerrar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
      return res.status(409).json({ ok:false, error:'La carga ya estÃ¡ cerrada.' });

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

    res.json({ ok:true, estado:'cerrada', cerrado_por: usuario, ts_cierre: now });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// â”€â”€ PATCH /api/bod/sesion/:id/carga-logistica/:cargaId/reabrir â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
      return res.status(409).json({ ok:false, error:'La carga no estÃ¡ cerrada.' });

    await supabase('PATCH', 'bod_furgon_cierres',
      { estado:'abierta' },
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));

    res.json({ ok:true, estado:'abierta', reabierto_por: usuario });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});

// -- POST /api/bod/sesion/:id/carga-logistica/:cargaId/enviar-hamilton --
// Envia una carga cerrada de Bodega CDG a Hamilton -> Traslados CDG.
app.post('/api/bod/sesion/:id/carga-logistica/:cargaId/enviar-hamilton', bodGuard, async (req, res) => {
  try {
    var sesId   = String(req.params.id).trim();
    var cargaId = String(req.params.cargaId).trim();
    var { usuario, auditor, supervisor } = req.body || {};
    var esAuditor = auditor === true || supervisor === true;
    if(!esAuditor) return res.status(403).json({ ok:false, error:'Solo auditores o supervisores pueden enviar manifiestos a Hamilton.' });

    usuario = String(usuario || '').trim();
    var rows = await supabase('GET', 'bod_furgon_cierres', null,
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId)+'&limit=1');
    if(!Array.isArray(rows)||!rows.length)
      return res.status(404).json({ ok:false, error:'Carga no encontrada.' });

    var cg = rows[0];
    if(cg.estado !== 'cerrada')
      return res.status(409).json({ ok:false, error:'Primero cerra la carga logistica antes de enviarla a Hamilton.' });
    if(cg.enviado_hamilton === true)
      return res.status(409).json({ ok:false, error:'Esta carga ya fue enviada a Hamilton como '+(cg.hamilton_contenedor||cg.licencia_hija||'licencia hija')+'.' });

    var faltantes = [];
    if(!String(cg.placa||'').trim())          faltantes.push('Placa');
    if(!String(cg.marchamo||'').trim())       faltantes.push('Marchamo');
    if(!String(cg.licencia_hija||'').trim())  faltantes.push('Licencia hija');
    if(!String(cg.destino_tr999||'').trim())  faltantes.push('Destino TR999');
    var tarimas = Array.isArray(cg.tarimas) ? cg.tarimas.filter(Boolean) : [];
    if(!tarimas.length) faltantes.push('Tarimas');
    var resumen = Array.isArray(cg.resumen_skus) ? cg.resumen_skus : [];
    if(!resumen.length) faltantes.push('Resumen de SKUs');
    if(faltantes.length)
      return res.status(400).json({ ok:false, error:'Faltan datos para enviar a Hamilton: '+faltantes.join(', ')+'.' });

    var licenciaHija = String(cg.licencia_hija||'').trim().toUpperCase();
    var destino      = String(cg.destino_tr999||'').trim().toUpperCase();
    var nowIso       = new Date().toISOString();
    var fechaEs      = new Date().toLocaleDateString('es');

    var items = resumen.map(function(s){
      var sku = String(s.sku || '').trim();
      var unidades = Number(s.unidades != null ? s.unidades : (s.qty != null ? s.qty : s.cantidad)) || 0;
      return {
        sku: sku,
        desc: s.descripcion || s.desc || '',
        qty: unidades,
        teoricoWMS: unidades,
        raw: {
          origen: 'Bodega CDG',
          status: 'Manifiesto Bodega CDG',
          tipo: 'CDG',
          licencia_padre: cg.licencia_padre || '',
          licencia_hija: licenciaHija,
          destino_tr999: destino,
          furgon: cg.furgon || '',
          placa: cg.placa || '',
          marchamo: cg.marchamo || '',
          tarimas: tarimas.join(', '),
          fecha: fechaEs
        }
      };
    }).filter(function(it){ return it.sku && it.qty > 0; });

    if(!items.length)
      return res.status(400).json({ ok:false, error:'El resumen de SKUs no tiene unidades validas para Hamilton.' });

    if(!state.cdg) state.cdg = {};
    if(!state.teorico) state.teorico = {};
    if(!state.fisico) state.fisico = {};

    var snapCdg          = JSON.parse(JSON.stringify(state.cdg || {}));
    var snapTeorico      = JSON.parse(JSON.stringify(state.teorico || {}));
    var snapFisico       = JSON.parse(JSON.stringify(state.fisico || {}));
    var snapVersion      = state.version;
    var snapHistorialLen = (state.historial||[]).length;

    state.cdg[licenciaHija] = {
      items: items,
      status: 'locked',
      autor: usuario,
      fecha: fechaEs,
      tipo: 'CDG',
      bloqueado: true,
      lastEditor: usuario,
      fromBodegaCDG: true,
      licenciaPadre: cg.licencia_padre || '',
      furgon: cg.furgon || '',
      placa: cg.placa || '',
      marchamo: cg.marchamo || '',
      destino_tr999: destino,
      tarimas: tarimas
    };

    state.teorico[licenciaHija] = {
      items: items,
      type: 'Traslados',
      fromCDG: true,
      fromBodegaCDG: true,
      cdgRef: licenciaHija,
      cdgValidado: true,
      cdgBloqueado: true,
      cdgTipo: 'CDG',
      fechaCarga: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Guatemala' }),
      meta: {
        origen: 'Bodega CDG',
        licencia_padre: cg.licencia_padre || '',
        licencia_hija: licenciaHija,
        furgon: cg.furgon || '',
        placa: cg.placa || '',
        marchamo: cg.marchamo || '',
        destino_tr999: destino,
        tarimas: tarimas
      }
    };
    state.fisico[licenciaHija] = null;

    addHistorial(usuario || '-', 'Bodega CDG enviado a Hamilton', (cg.licencia_padre||sesId)+' -> '+licenciaHija);
    state.version++;

    try {
      if(typeof saveDailyStateStrict === 'function') {
        await withTimeout(saveDailyStateStrict('BOD enviar Hamilton'), 20000, 'BOD enviar Hamilton save');
      } else {
        scheduleSave();
      }
    } catch(saveErr) {
      state.cdg      = snapCdg;
      state.teorico  = snapTeorico;
      state.fisico   = snapFisico;
      state.version  = snapVersion;
      if(state.historial) state.historial.length = snapHistorialLen;
      return res.status(500).json({ ok:false, error:'No se pudo guardar el traslado Hamilton: '+saveErr.message });
    }

    await supabase('PATCH', 'bod_furgon_cierres',
      { enviado_hamilton:true, enviado_hamilton_por:usuario, ts_envio_hamilton:nowIso, hamilton_contenedor:licenciaHija },
      '?id=eq.'+encodeURIComponent(cargaId)+'&sesion_id=eq.'+encodeURIComponent(sesId));

    res.json({ ok:true, licencia_hija:licenciaHija, hamilton_contenedor:licenciaHija, items:items.length, unidades:items.reduce(function(a,b){ return a + (Number(b.qty)||0); }, 0) });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
  }
});
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// BOD MODULE END
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// BOD WMS / TRAZABILIDAD â€” POST /api/bod/wms/upload
// Parsea Excel WMS con columnas: NÃºmero, Fecha, Status, SKU, Nombre,
// Unidades, Origen, Destino. Clasifica movimientos y guarda en bod_wms_movimientos.
// FIX (mar 2-jun-2026, server v20): mÃ³dulo de trazabilidad BOD Fase 2.
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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
    if(rows.length < 2) return res.status(400).json({ ok:false, error:'Archivo vacÃ­o.' });

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
      var norm = l.split(/[\sÂ·\-Â·]+/)[0];
      var esp  = licenciaId.split(/[\sÂ·\-Â·]+/)[0];
      return norm !== esp;
    });
    if(licsDistintas.length > 0) {
      return res.status(400).json({
        ok:false,
        error:'El archivo contiene licencias: '+Object.keys(licenciasEnArchivo).join(', ')+
              '. La sesiÃ³n activa es "'+licenciaId+'". VerificÃ¡ el archivo o la sesiÃ³n.'
      });
    }

    // Modo reemplazo selectivo segÃºn tipo_licencia
    try {
      if(tipoLicencia === 'hija_salida') {
        // Reemplazar solo movimientos de esta licencia hija + furgÃ³n
        await supabase('DELETE', 'bod_wms_movimientos', null,
          '?licencia_id=eq.'+encodeURIComponent(licenciaId)
          +'&furgon_relacionado=eq.'+encodeURIComponent(furgonRelacionado)
          +'&tipo_licencia=eq.hija_salida');
      } else {
        // Reemplazar movimientos iniciales de esta licencia
        await supabase('DELETE', 'bod_wms_movimientos', null,
          '?licencia_id=eq.'+encodeURIComponent(licenciaId)+'&tipo_licencia=eq.inicial');
        // Fallback: si la tabla no tiene la col aÃºn, borrar todos los de esta licencia
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

    if(!movimientos.length) return res.status(400).json({ ok:false, error:'Sin filas vÃ¡lidas.' });

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
      advertencia = 'El archivo fue cargado, pero no contiene movimientos con destino 952.006.01. WMS vs App mostrarÃ¡ cero teÃ³rico.';
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

// â”€â”€ Helpers para construir filtro de licencias (punto 5) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€ GET /api/bod/reportes/wms-vs-app â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Acepta: licencia_id | fecha | licencias (CSV)
// â”€â”€ Helper: construir mapa SKU â†’ contexto logÃ­stico â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

  // Mapas: tarima â†’ furgon, tarima â†’ cierre
  var tarimaFurgon = {}, tarimaCierre = {};
  asigs.forEach(function(a){ tarimaFurgon[a.tarima] = a.furgon; });
  cierres.forEach(function(cg){
    var ts = Array.isArray(cg.tarimas) ? cg.tarimas : [];
    ts.forEach(function(t){ tarimaCierre[t] = cg; });
  });

  // Cruzar lÃ­neas â†’ tarima â†’ furgon/cierre â†’ skuLog
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

// â”€â”€ GET /api/bod/reportes/wms-vs-app â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€ GET /api/bod/reportes/flujo-bolson â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    // Salidas exclusivamente desde cierres logÃ­sticos (bod_furgon_cierres.resumen_skus)
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

// â”€â”€ GET /api/bod/reportes/remanentes â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

      // Causa operativa enriquecida con contexto logÃ­stico
      var causa;
      if(d.entradas === 0 && d.destinos_distintos)
        causa = 'Pendiente por ubicaciÃ³n distinta a 952';
      else if(d.entradas > 0 && app === 0)
        causa = 'Pendiente de entarimar';
      else if(d.entradas === 0 && app > 0)
        causa = 'Entarimado sin WMS';
      else if(app > 0 && !furgon)
        causa = 'Pendiente de asignaciÃ³n de furgÃ³n';
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

// â”€â”€ GET /api/bod/reportes/furgones â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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


