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
    state.teorico[num] = {
      type:         'Traslados',
      fromCDG:      true,
      cdgRef:       contId,
      cdgValidado:  true,
      cdgBloqueado: true,
      items: items.map(i => ({
        sku:  i.sku,
        desc: i.desc,
        qty:  i.qty,
        raw:  { origen:'CDG', status:'CDG Validado', fecha:new Date().toLocaleDateString('es') }
      }))
    };
    state.fisico[num] = null;
  } else {
    state.teorico[num].cdgValidado = true;
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
app.post('/api/teorico/fecha-carga', (req, res) => {
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
  state.teorico[cont].fechaCarga = fechaCarga || null;
  addHistorial(usuario||'—', 'Cambió fecha de trabajo a ' + (fechaCarga || '(vacío)'), cont);
  state.version++;
  res.json({ ok:true, version:state.version, fechaCarga: state.teorico[cont].fechaCarga });
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
      return res.status(500).json({
        ok: false,
        error: 'El archivo sí se leyó, pero NO quedó guardado. No cierres la app. Volvé a subirlo. (' + saveErr.message + ')'
      });
    }

    res.json({ ok:true, loaded });
  } catch(e) {
    res.status(500).json({ ok:false, error:e.message });
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
  const ex = {
    cFecha:    findCol(hdr, ['fecha']),
    cCodProv:  findCol(hdr, ['proveedor','código de proveedor']),
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
  dataRows
    .filter(r => r && r.some(c => String(c).trim() !== ''))
    .forEach(row => {
      const cont = String(row[colCont]||'').trim();
      if(!cont) return;
      if(!newConts[cont]) newConts[cont] = [];
      newConts[cont].push({
        sku:  String(row[colSku] ||'').trim(),
        desc: String(row[colDesc]||'').trim(),
        qty:  parseFloat(String(row[colQty]).replace(',','.')) || 0,
        raw: {
          fecha:     g(row, ex.cFecha),
          codProv:   g(row, ex.cCodProv),
          lineas:    g(row, ex.cLineas),
          status:    g(row, ex.cStatus),
          oc:        g(row, ex.cOC),
          ingresado: g(row, ex.cIngr),
          colocado:  g(row, ex.cColoc),
          faltantes: g(row, ex.cFalt),
          sobrantes: g(row, ex.cSobr),
          daniado:   g(row, ex.cDan),
          origen:    g(row, ex.cOrigen),
          ingreso:   g(row, ex.cIngreso),
          docSap:    g(row, ex.cDocSap),
          tipo:      g(row, ex.cTipo),
          unidad:    g(row, ex.cUnidad),
          unidades:  g(row, ex.cUnidades),
          destino:   g(row, ex.cDestino)
        }
      });
    });

  Object.keys(newConts).forEach(cont => {
    // Don't overwrite CDG-validated containers from teorico upload
    if(state.teorico[cont] && state.teorico[cont].fromCDG) {
      state.teorico[cont].cdgValidado = true;
      return;
    }
    // FIX (mar 19-may-2026): preservar fechaCarga si el contenedor ya
    // existía. Solo asignar fecha de hoy a contenedores NUEVOS.
    // Esto evita que al re-subir el Excel maestro (que incluye contenedores
    // previamente cargados), las fechas de los viejos se actualicen.
    //
    // FIX (rev Claude2): usar timezone Guatemala (America/Guatemala) en vez
    // de UTC. Sin esto, uploads del Excel a las 18:00+ hora local quedaban
    // con fecha del día siguiente.
    var fechaCargaPrev = (state.teorico[cont] && state.teorico[cont].fechaCarga) || null;
    var hoy = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Guatemala' });
    state.teorico[cont] = {
      items: newConts[cont],
      type,
      fechaCarga: fechaCargaPrev || hoy
    };
    // Preserve existing fisico data — never overwrite conteo work
    if(!state.fisico[cont]) state.fisico[cont] = null;
  });
  return Object.keys(newConts).length;
}

// ── Start ─────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
// Always start server — even if Supabase fails
app.listen(PORT, () => console.log('Conteo app on port ' + PORT));
// Load state after server is up
loadState().catch(e => console.log('State load failed (non-fatal):', e.message));
