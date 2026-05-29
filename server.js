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
const SUPABASE_KEY = process.env.SUPABASE_KEY;  // service_role key

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
  res.json(publicState());
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
  if(!state.cdg[contId]) {
    state.cdg[contId] = {
      items:  [],
      status: 'open',
      autor:  usuario,
      fecha:  new Date().toLocaleDateString('es')
    };
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
    state.teorico[num] = {
      type:         'Traslados',
      fromCDG:      true,
      cdgRef:       contId,
      cdgValidado:  true,
      cdgBloqueado: true,
      cdgTipo:      tipoEfectivo,
      items: items.map(i => ({
        sku:  i.sku,
        desc: i.desc,
        qty:  i.qty,
        raw:  { origen:tipoEfectivo, status:tipoEfectivo+' Validado', fecha:new Date().toLocaleDateString('es') }
      }))
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
//
// NOTA sobre el patrón de rollback (server v18, 28-may-2026):
// Originalmente este endpoint fue el "modelo" de rollback. En v18 ese patrón
// se generalizó y refinó: los 5 endpoints con await (`fecha-carga`,
// `clasificacion/set`, `cdg/delete`, `upload`, `wms/sincronizar`) ahora
// comparten el mismo patrón uniforme — snapshot de version + historial
// ANTES de mutar, restauración completa en catch. Ver comentarios in-line
// "FIX (mié 28-may-2026, server v18)" en cada endpoint para el detalle.
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

  // Persistir await (patrón v18: snapshot + try/await/catch + rollback completo)
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

// ── Start ─────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
// Always start server — even if Supabase fails
app.listen(PORT, () => console.log('Conteo app on port ' + PORT));
// Load state after server is up
loadState().catch(e => console.log('State load failed (non-fatal):', e.message));
