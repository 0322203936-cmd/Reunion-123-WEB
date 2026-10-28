const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { Pool } = require('pg');
const { runActualizarData } = require('./actualizar_data');
const { runActualizarInventario } = require('./actualizar_inventario');
const CosechaCor = require('./cosechaCor');
const AppendBuilder = require('./appendBuilder');
const crypto = require('crypto');

function loadLocalEnv() {
  const envPath = path.join(__dirname, '.env');
  if (!fs.existsSync(envPath)) return;
  for (const rawLine of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const separator = line.indexOf('=');
    if (separator < 1) continue;
    const key = line.slice(0, separator).trim();
    let value = line.slice(separator + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (process.env[key] === undefined) process.env[key] = value;
  }
}

loadLocalEnv();

const PORT = Number(process.env.PORT || 3070);
const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, 'data');
const DATABASE_URL = process.env.DATABASE_URL;
const GITHUB_ACTIONS_TOKEN = String(process.env.GITHUB_ACTIONS_TOKEN || '').trim();
const GITHUB_REPOSITORY = String(process.env.GITHUB_REPOSITORY || '0322203936-cmd/Reunion-123-WEB').trim();
const REUNION_SYNC_TOKEN = String(process.env.REUNION_SYNC_TOKEN || '').trim();

// Each workbook ("apartado") has its own files and its own database tables.
// "reunion" keeps exactly the files/tables it always used.
const WORKBOOKS = {
  reunion: {
    name: 'Reunion 123',
    initial: path.join(DATA_DIR, 'initial_workbook.json'),
    state: path.join(DATA_DIR, 'workbook_state.json'),
    homolog: path.join(DATA_DIR, 'homologaciones.json'),
    prefix: 'reunion123'
  },
  cosecha: {
    name: 'Plan de cosecha',
    initial: path.join(DATA_DIR, 'cosecha_initial_workbook.json'),
    state: path.join(DATA_DIR, 'cosecha_workbook_state.json'),
    homolog: path.join(DATA_DIR, 'cosecha_homologaciones.json'),
    prefix: 'cosecha'
  },
  rvp: {
    name: 'Requerimiento Vs Proyeccion',
    initial: path.join(DATA_DIR, 'rvp_initial_workbook.json'),
    state: path.join(DATA_DIR, 'rvp_workbook_state.json'),
    homolog: path.join(DATA_DIR, 'rvp_homologaciones.json'),
    prefix: 'rvp'
  }
};
const DEFAULT_WB = 'reunion';

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

let pool = null;
// Render should start serving the page even if Neon is waking up. Until the
// database finishes initializing, read-only workbook requests use the local
// snapshot and do not cache it over the database-backed workbook.
let persistenceReady = !DATABASE_URL;
const caches = {}; // wbKey -> full workbook kept in memory; PATCH merges into it
const workbookLoads = {}; // wbKey -> shared in-flight Neon read
const gzCache = {}; // wbKey -> gzipped JSON of the cached workbook (dropped on every change)
const gzPending = {};
const workbookGeneration = {};
let actualizarDataRunning = false;
let actualizarInventarioRunning = false;
const actualizarDataState = {
  status: 'idle',
  startedAt: null,
  finishedAt: null,
  rows: null,
  lastRow: null,
  range: null,
  dateRange: null,
  error: null
};

const actualizarInventarioState = {
  status: 'idle',
  startedAt: null,
  finishedAt: null,
  rows: null,
  lastRow: null,
  error: null
};

function publicActualizarDataState() {
  return { ...actualizarDataState, ok: actualizarDataState.status === 'succeeded' };
}

function publicActualizarInventarioState() {
  return { ...actualizarInventarioState, ok: actualizarInventarioState.status === 'succeeded' };
}

function resetWorkbookCaches() {
  for (const key of Object.keys(WORKBOOKS)) {
    delete caches[key];
    invalidateGzip(key);
  }
}

function invalidateGzip(key) {
  delete gzCache[key];
  workbookGeneration[key] = (workbookGeneration[key] || 0) + 1;
}

async function compressedWorkbook(key, workbook) {
  if (gzCache[key]) return gzCache[key];
  const generation = workbookGeneration[key] || 0;
  if (!gzPending[key] || gzPending[key].generation !== generation) {
    const json = JSON.stringify(workbook);
    const promise = new Promise((resolve, reject) => {
      // Level 1 is intentionally used here: the response is only slightly
      // larger, while the free Render instance spends far less CPU after an edit.
      zlib.gzip(json, { level: 1 }, (error, data) => error ? reject(error) : resolve(data));
    });
    gzPending[key] = { generation, promise };
  }
  const entry = gzPending[key];
  const data = await entry.promise;
  if ((workbookGeneration[key] || 0) === entry.generation) gzCache[key] = data;
  if (gzPending[key] === entry) delete gzPending[key];
  return data;
}

function stateForOperation(operation) {
  return operation === 'inventario' ? actualizarInventarioState : actualizarDataState;
}

async function dispatchGitHubPosco(operation) {
  if (!GITHUB_ACTIONS_TOKEN) return false;
  const response = await fetch(`https://api.github.com/repos/${GITHUB_REPOSITORY}/dispatches`, {
    method: 'POST',
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${GITHUB_ACTIONS_TOKEN}`,
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'reunion-123'
    },
    body: JSON.stringify({ event_type: 'posco-update', client_payload: { operation } })
  });
  if (!response.ok) throw new Error(`GitHub rechazó el workflow (HTTP ${response.status}).`);
  return true;
}

async function receiveGitHubPoscoCallback(payload) {
  const operation = String(payload?.operation || '').trim().toLowerCase();
  if (!['data', 'inventario'].includes(operation)) throw new Error('Operación POSCO inválida.');
  const state = stateForOperation(operation);
  if (payload.status === 'succeeded') {
    // GitHub Actions escribió Neon desde fuera del proceso de Render.
    // Se limpian las cachés para que las pestañas reciban esos datos.
    resetWorkbookCaches();
    await syncCor();
    state.status = 'succeeded';
    state.rows = Number.isFinite(Number(payload.rows)) ? Number(payload.rows) : null;
    state.lastRow = Number.isFinite(Number(payload.lastRow)) ? Number(payload.lastRow) : null;
    state.range = payload.range || null;
    state.dateRange = payload.dateRange || null;
    state.error = null;
  } else {
    state.status = 'failed';
    state.error = payload.error || 'GitHub Actions no pudo completar el proceso POSCO.';
  }
  state.finishedAt = new Date().toISOString();
  if (operation === 'inventario') actualizarInventarioRunning = false;
  else actualizarDataRunning = false;
}

// ---------------------------------------------------------------- Plan de cosecha -> Cor de Reunion (real time)
// Whenever either book changes, the Cor of Reunion is rewritten from the last P Cosecha tab
// and every open Reunion page is told right away (server-sent events).
const corClients = new Set();
let corFeed = null;
let corTimer = null;
let appendTimer = null;
// change history of the Cor (last 5 changes per product + date), kept in data/cor_history.json
const COR_HISTORY_FILE = path.join(DATA_DIR, 'cor_history.json');
let corHistory = null;
let corHistoryTimer = null;
function loadCorHistory() {
  if (!corHistory) corHistory = readJson(COR_HISTORY_FILE, null) || { source: null, baseline: null, entries: {} };
  return corHistory;
}
function saveCorHistorySoon() {
  clearTimeout(corHistoryTimer);
  corHistoryTimer = setTimeout(() => {
    const tmp = `${COR_HISTORY_FILE}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(corHistory), 'utf8');
    fs.renameSync(tmp, COR_HISTORY_FILE);
  }, 300);
}
function publicCorFeed() {
  const h = loadCorHistory();
  return { ...corFeed, history: { source: h.source, entries: h.entries } };
}
function broadcastCor() {
  const msg = `event: cor\ndata: ${JSON.stringify(publicCorFeed())}\n\n`;
  for (const res of corClients) res.write(msg);
}
function broadcastWorkbookMutation(key, patch) {
  const hasStructuralChange = Object.keys((patch && patch.sheets) || {}).length ||
    ((patch && patch.deleted) || []).length;
  const event = hasStructuralChange ? 'workbook-refresh' : 'workbook-patch';
  const payload = hasStructuralChange
    ? { key }
    : { key, cellPatches: (patch && patch.cellPatches) || {} };
  const msg = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const res of corClients) res.write(msg);
}
function scheduleCorSync(delay = 150) {
  clearTimeout(corTimer);
  corTimer = setTimeout(() => { syncCor().catch(err => console.error('Sync Cor:', err)); }, delay);
}
function scheduleAppendSync(delay = 1200) {
  clearTimeout(appendTimer);
  appendTimer = setTimeout(async () => {
    try {
      await syncAppend(await loadWorkbook('reunion'), await loadWorkbook('cosecha'));
    } catch (err) {
      console.error('Sync Append:', err);
    }
  }, delay);
}
async function syncCor({ rebuildAppend = true } = {}) {
  const cosecha = await loadWorkbook('cosecha');
  const reunion = await loadWorkbook('reunion');
  const feed = CosechaCor.computeFeed(cosecha);
  const changedCells = {};
  const changed = CosechaCor.applyFeed(reunion, feed, changedCells);
  if (changed) {
    await patchWorkbook('reunion', {
      cellPatches: { Reunion: { set: changedCells, delete: [] } }
    });
  }
  const feedChanged = !corFeed || corFeed.sig !== feed.sig;
  corFeed = feed;
  if (CosechaCor.trackChanges(loadCorHistory(), feed, new Date().toISOString())) saveCorHistorySoon();
  if (changed || feedChanged) broadcastCor();
  if (rebuildAppend) scheduleAppendSync();
}

// Requerimiento Vs Proyeccion: sheet "Append" rebuilt from PegarData + Plan de cosecha
let appendSig = null;
async function syncAppend(reunion, cosecha) {
  const rvp = await loadWorkbook('rvp');
  const { sheet } = AppendBuilder.buildAppendSheet(reunion, cosecha, rvp);
  const md5 = obj => crypto.createHash('md5').update(JSON.stringify(obj)).digest('hex');
  const sig = md5(sheet.cells);
  // nothing new, and nobody overwrote it with an older copy
  if (sig === appendSig && rvp.sheets.Append && md5(rvp.sheets.Append.cells) === sig) return;
  const order = rvp.sheetOrder || [];
  const meta = order.includes('Append') ? {} : { sheetOrder: [...order, 'Append'] };
  const prev = rvp.sheets.Append;
  if (prev && prev.colWidths) sheet.colWidths = { ...sheet.colWidths, ...prev.colWidths }; // keep widths set by hand
  appendSig = sig;
  await patchWorkbook('rvp', { meta, sheets: { Append: sheet } });
  const msg = `event: append\ndata: ${JSON.stringify({ sig })}\n\n`;
  for (const res of corClients) res.write(msg);
}

async function runActualizarDataInBackground() {
  try {
    console.log(' Iniciando Actualizar Data local: POSCO -> PegarData...');
    const currentWorkbook = await loadWorkbook('reunion');
    const result = await runActualizarData(currentWorkbook);
    await saveWholeWorkbook('reunion', result.workbook);
    invalidateGzip('reunion');
    scheduleCorSync();
    actualizarDataState.status = 'succeeded';
    actualizarDataState.finishedAt = new Date().toISOString();
    actualizarDataState.rows = result.rows;
    actualizarDataState.lastRow = result.lastRow;
    actualizarDataState.range = `A1:W${result.lastRow}`;
    actualizarDataState.dateRange = {
      start: result.dateRange.start.toISOString().slice(0, 10),
      end: result.dateRange.end.toISOString().slice(0, 10)
    };
    actualizarDataState.error = null;
    console.log(` Actualizar Data completado: ${result.rows} filas en PegarData.`);
  } catch (error) {
    actualizarDataState.status = 'failed';
    actualizarDataState.finishedAt = new Date().toISOString();
    actualizarDataState.error = error.message || 'No se pudo actualizar PegarData.';
    console.error(' Actualizar Data falló:', error);
  } finally {
    actualizarDataRunning = false;
  }
}

async function runActualizarInventarioInBackground() {
  try {
    console.log(' Iniciando Actualizar Inventario local: POSCO -> Inventario...');
    const currentWorkbook = await loadWorkbook('reunion');
    const result = await runActualizarInventario(currentWorkbook);
    // the day the inventory was taken: its week (Sat..Fri) is the week the stock counts in Requerimiento Vs Proyeccion
    if (result.workbook.sheets && result.workbook.sheets.Inventario) result.workbook.sheets.Inventario.updatedAt = new Date().toISOString();
    await saveWholeWorkbook('reunion', result.workbook);
    invalidateGzip('reunion');
    scheduleCorSync();
    actualizarInventarioState.status = 'succeeded';
    actualizarInventarioState.finishedAt = new Date().toISOString();
    actualizarInventarioState.rows = result.rows;
    actualizarInventarioState.lastRow = result.lastRow;
    actualizarInventarioState.error = null;
    console.log(` Actualizar Inventario completado: ${result.rows} filas en Inventario.`);
  } catch (error) {
    actualizarInventarioState.status = 'failed';
    actualizarInventarioState.finishedAt = new Date().toISOString();
    actualizarInventarioState.error = error.message || 'No se pudo actualizar Inventario.';
    console.error(' Actualizar Inventario falló:', error);
  } finally {
    actualizarInventarioRunning = false;
  }
}

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch (error) {
    return fallback;
  }
}

function cfg(key) {
  return WORKBOOKS[key] || WORKBOOKS[DEFAULT_WB];
}

function initialWorkbook(key) {
  const c = cfg(key);
  return readJson(c.initial, { workbookName: c.name, sheetOrder: [], sheets: {} });
}

function initialHomologations(key) {
  return readJson(cfg(key).homolog, { version: 1, items: {} });
}

function splitWorkbook(workbook) {
  const { sheets = {}, ...meta } = workbook;
  return { meta, sheets };
}

function sheetWithoutCells(sheet) {
  if (!sheet || typeof sheet !== 'object') return sheet;
  const { cells, ...structure } = sheet;
  return structure;
}

function sheetCellRows(sheetName, sheet) {
  return Object.entries((sheet && sheet.cells) || {}).map(([coord, cell]) => ({
    sheet_name: sheetName,
    coord,
    cell
  }));
}

async function insertCellRows(client, table, rows) {
  for (let start = 0; start < rows.length; start += 2000) {
    const batch = rows.slice(start, start + 2000);
    if (!batch.length) continue;
    await client.query(
      `INSERT INTO ${table} (sheet_name, coord, cell)
       SELECT sheet_name, coord, cell
       FROM jsonb_to_recordset($1::jsonb) AS x(sheet_name text, coord text, cell jsonb)
       ON CONFLICT (sheet_name, coord) DO UPDATE SET cell = EXCLUDED.cell, updated_at = NOW()`,
      [JSON.stringify(batch)]
    );
  }
}

async function migrateWorkbookCells(key) {
  const p = cfg(key).prefix;
  const marker = `${p}_workbook_migrations`;
  const done = await pool.query(`SELECT 1 FROM ${marker} WHERE name = 'cells-v1'`);
  if (done.rowCount) return;

  // A workbook created by the new writer already has normalized cell rows.
  // Never delete those rows just because the migration marker is missing.
  const existingCells = await pool.query(`SELECT 1 FROM ${p}_workbook_cells LIMIT 1`);
  if (existingCells.rowCount) {
    await pool.query(`INSERT INTO ${marker} (name) VALUES ('cells-v1') ON CONFLICT (name) DO NOTHING`);
    return;
  }

  const sheetRes = await pool.query(`SELECT name, data FROM ${p}_workbook_sheets`);
  const rows = [];
  for (const row of sheetRes.rows) rows.push(...sheetCellRows(row.name, row.data));
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM ${p}_workbook_cells`);
    await insertCellRows(client, `${p}_workbook_cells`, rows);
    await client.query(`INSERT INTO ${marker} (name) VALUES ('cells-v1') ON CONFLICT (name) DO NOTHING`);
    await client.query('COMMIT');
    console.log(` Migración Neon ${key}: ${rows.length} celdas.`);
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally { client.release(); }
}

// ---------------------------------------------------------------- database
async function initDatabase() {
  if (!DATABASE_URL) return;

  pool = new Pool({
    connectionString: DATABASE_URL,
    ssl: { rejectUnauthorized: false },
    max: 5,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 10000,
    // Never let a slow Neon query hold a browser request indefinitely.
    statement_timeout: 15000,
    query_timeout: 20000
  });

  for (const key of Object.keys(WORKBOOKS)) {
    const p = cfg(key).prefix;
    await pool.query(`
      CREATE TABLE IF NOT EXISTS ${p}_workbook_meta (
        id SMALLINT PRIMARY KEY CHECK (id = 1),
        meta JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS ${p}_workbook_sheets (
        name TEXT PRIMARY KEY,
        data JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS ${p}_workbook_cells (
        sheet_name TEXT NOT NULL,
        coord TEXT NOT NULL,
        cell JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        PRIMARY KEY (sheet_name, coord)
      );

      CREATE TABLE IF NOT EXISTS ${p}_workbook_migrations (
        name TEXT PRIMARY KEY,
        completed_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS ${p}_homologations (
        id SMALLINT PRIMARY KEY CHECK (id = 1),
        state JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      );
    `);

    const existing = await pool.query(`SELECT 1 FROM ${p}_workbook_meta WHERE id = 1`);
    if (existing.rowCount === 0) {
      await writeWholeWorkbookToDb(key, readJson(cfg(key).state, initialWorkbook(key)));
    }
    await pool.query(
      `INSERT INTO ${p}_homologations (id, state) VALUES (1, $1::jsonb) ON CONFLICT (id) DO NOTHING`,
      [JSON.stringify(initialHomologations(key))]
    );
  }

  console.log(' Persistencia activa: Neon PostgreSQL');
}

async function migrateAllWorkbookCellsInBackground() {
  for (const key of Object.keys(WORKBOOKS)) {
    try {
      await migrateWorkbookCells(key);
    } catch (error) {
      console.error(`Migración Neon ${key} pendiente:`, error.message);
    }
  }
}

async function writeWholeWorkbookToDb(key, workbook) {
  const p = cfg(key).prefix;
  const { meta, sheets } = splitWorkbook(workbook);
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM ${p}_workbook_sheets`);
    await client.query(`DELETE FROM ${p}_workbook_cells`);
    for (const [name, data] of Object.entries(sheets)) {
      await client.query(
        `INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)`,
        [name, JSON.stringify(sheetWithoutCells(data))]
      );
    }
    const cellRows = [];
    for (const [name, data] of Object.entries(sheets)) cellRows.push(...sheetCellRows(name, data));
    await insertCellRows(client, `${p}_workbook_cells`, cellRows);
    await client.query(
      `INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb)
       ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`,
      [JSON.stringify(meta)]
    );
    await client.query('COMMIT');
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}

// ---------------------------------------------------------------- workbook persistence
async function loadWorkbook(key) {
  if (caches[key] && persistenceReady) return caches[key];
  if (workbookLoads[key]) return workbookLoads[key];
  workbookLoads[key] = loadWorkbookUncached(key).finally(() => {
    delete workbookLoads[key];
  });
  return workbookLoads[key];
}

async function loadWorkbookUncached(key) {
  if (caches[key] && persistenceReady) return caches[key];
  if (!pool || !persistenceReady) {
    const localWorkbook = caches[key] || readJson(cfg(key).state, null) || initialWorkbook(key);
    if (persistenceReady) caches[key] = localWorkbook;
    return localWorkbook;
  }
  const p = cfg(key).prefix;
  try {
    // A stalled Neon connection must not leave the browser request hanging.
    // The local snapshot is only a recovery path; Neon remains the source
    // whenever this read completes normally.
    const dbLoad = (async () => {
    const metaRes = await pool.query(`SELECT meta FROM ${p}_workbook_meta WHERE id = 1`);
    const sheetRes = await pool.query(`SELECT name, data FROM ${p}_workbook_sheets`);
    const cellRes = await pool.query(`SELECT sheet_name, coord, cell FROM ${p}_workbook_cells`);
    const meta = metaRes.rows[0]?.meta || splitWorkbook(initialWorkbook(key)).meta;
    const sheets = {};
    for (const row of sheetRes.rows) {
      // During the one-time migration, the legacy JSONB may still contain
      // cells. Keep them as a temporary fallback and overlay normalized rows.
      sheets[row.name] = { ...row.data, cells: { ...((row.data && row.data.cells) || {}) } };
    }
    for (const row of cellRes.rows) {
      if (!sheets[row.sheet_name]) sheets[row.sheet_name] = { cells: {} };
      sheets[row.sheet_name].cells[row.coord] = row.cell;
    }
    caches[key] = { ...meta, sheets };
    return caches[key];
    })();
    return await Promise.race([
      dbLoad,
      new Promise((_, reject) => setTimeout(() => reject(new Error('Neon tardó demasiado en responder')), 20000))
    ]);
  } catch (error) {
    // A temporary Neon failure must not leave the workbook request pending.
    // The local snapshot is the safe read fallback; the next request can try
    // Neon again after the service recovers.
    console.error(`No se pudo leer ${key} desde Neon; usando respaldo local:`, error.message);
    return caches[key] = readJson(cfg(key).state, null) || initialWorkbook(key);
  }
}

const fileWriteTimers = {};
function persistFileSoon(key) {
  clearTimeout(fileWriteTimers[key]);
  fileWriteTimers[key] = setTimeout(() => {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    const target = cfg(key).state;
    const tmp = `${target}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(caches[key]), 'utf8');
    fs.renameSync(tmp, target);
  }, 400);
}

async function saveWholeWorkbook(key, workbook) {
  caches[key] = workbook;
  invalidateGzip(key);
  if (!pool) persistFileSoon(key);
  else await writeWholeWorkbookToDb(key, workbook);
  broadcastWorkbookMutation(key, { sheets: workbook.sheets || {} });
}

async function patchWorkbook(key, patch) {
  const workbook = await loadWorkbook(key);
  invalidateGzip(key);
  const meta = patch.meta || {};
  for (const k of ['workbookName', 'sheetOrder', 'tables', 'activeSheet', 'pivots']) {
    if (meta[k] !== undefined) workbook[k] = meta[k];
  }
  const changed = patch.sheets || {};
  const cellPatches = patch.cellPatches && typeof patch.cellPatches === 'object'
    ? patch.cellPatches : {};
  const deleted = Array.isArray(patch.deleted) ? patch.deleted : [];
  for (const [name, data] of Object.entries(changed)) workbook.sheets[name] = data;
  for (const [name, cellPatch] of Object.entries(cellPatches)) {
    const sheet = workbook.sheets[name];
    if (!sheet || !cellPatch || typeof cellPatch !== 'object') continue;
    if (!sheet.cells || typeof sheet.cells !== 'object') sheet.cells = {};
    const set = cellPatch.set && typeof cellPatch.set === 'object' ? cellPatch.set : {};
    for (const [coord, cell] of Object.entries(set)) sheet.cells[coord] = cell;
    for (const coord of (Array.isArray(cellPatch.delete) ? cellPatch.delete : [])) {
      delete sheet.cells[coord];
    }
  }
  for (const name of deleted) delete workbook.sheets[name];

  // In local mode the snapshot is the persistence layer. When Neon is active,
  // do not stringify/write the entire workbook on every cell edit; that can
  // block the event loop and make the next reload appear stuck.
  if (!pool) {
    persistFileSoon(key);
    broadcastWorkbookMutation(key, patch);
    return;
  }

  const p = cfg(key).prefix;
  const onlyCellChanges = !Object.keys(changed).length && !deleted.length && !Object.keys(meta).length;
  const client = await pool.connect();
  if (onlyCellChanges) {
    try {
      const setRows = [];
      const deleteRows = [];
      for (const [name, cellPatch] of Object.entries(cellPatches)) {
        for (const [coord, cell] of Object.entries((cellPatch && cellPatch.set) || {})) {
          setRows.push({ sheet_name: name, coord, cell });
        }
        for (const coord of (Array.isArray(cellPatch && cellPatch.delete) ? cellPatch.delete : [])) {
          deleteRows.push({ sheet_name: name, coord });
        }
      }
      await insertCellRows(client, `${p}_workbook_cells`, setRows);
      if (deleteRows.length) {
        await client.query(
          `DELETE FROM ${p}_workbook_cells c
           USING jsonb_to_recordset($1::jsonb) AS x(sheet_name text, coord text)
           WHERE c.sheet_name = x.sheet_name AND c.coord = x.coord`,
          [JSON.stringify(deleteRows)]
        );
      }
      broadcastWorkbookMutation(key, patch);
      return;
    } finally { client.release(); }
  }
  try {
    await client.query('BEGIN');
    for (const [name, data] of Object.entries(changed)) {
      await client.query(
        `INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)
         ON CONFLICT (name) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()`,
        [name, JSON.stringify(sheetWithoutCells(data))]
      );
      await client.query(`DELETE FROM ${p}_workbook_cells WHERE sheet_name = $1`, [name]);
      await insertCellRows(client, `${p}_workbook_cells`, sheetCellRows(name, data));
    }
    // Apply individual cells directly to Neon rows. A one-cell edit no longer
    // serializes or rewrites an entire sheet JSON document.
    for (const [name, cellPatch] of Object.entries(cellPatches)) {
      if (!workbook.sheets[name] || changed[name]) continue;
      const set = cellPatch && cellPatch.set && typeof cellPatch.set === 'object'
        ? cellPatch.set : {};
      for (const [coord, cell] of Object.entries(set)) {
        await client.query(
          `INSERT INTO ${p}_workbook_cells (sheet_name, coord, cell)
           VALUES ($1, $2, $3::jsonb)
           ON CONFLICT (sheet_name, coord) DO UPDATE SET cell = EXCLUDED.cell, updated_at = NOW()`,
          [name, coord, JSON.stringify(cell)]
        );
      }
      for (const coord of (Array.isArray(cellPatch && cellPatch.delete) ? cellPatch.delete : [])) {
        await client.query(
          `DELETE FROM ${p}_workbook_cells WHERE sheet_name = $1 AND coord = $2`,
          [name, coord]
        );
      }
    }
    for (const name of deleted) {
      await client.query(`DELETE FROM ${p}_workbook_sheets WHERE name = $1`, [name]);
    }
    if (Object.keys(meta).length) {
      const { meta: fullMeta } = splitWorkbook(workbook);
      await client.query(
        `INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb)
         ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`,
        [JSON.stringify(fullMeta)]
      );
    }
    await client.query('COMMIT');
    broadcastWorkbookMutation(key, patch);
  } catch (error) {
    await client.query('ROLLBACK');
    // Keep a recovery snapshot only when Neon fails, so a transient database
    // problem does not discard the user's latest local state.
    persistFileSoon(key);
    throw error;
  } finally {
    client.release();
  }
}

async function resetWorkbook(key) {
  const workbook = initialWorkbook(key);
  await saveWholeWorkbook(key, workbook);
  return workbook;
}

// ---------------------------------------------------------------- homologations
async function getHomologations(key) {
  if (!pool) return readJson(cfg(key).homolog, initialHomologations(key));
  const result = await pool.query(`SELECT state FROM ${cfg(key).prefix}_homologations WHERE id = 1`);
  return result.rows[0]?.state || initialHomologations(key);
}

async function saveHomologations(key, state) {
  if (!pool) {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.writeFileSync(cfg(key).homolog, JSON.stringify(state, null, 2), 'utf8');
    return;
  }
  await pool.query(
    `UPDATE ${cfg(key).prefix}_homologations SET state = $1::jsonb, updated_at = NOW() WHERE id = 1`,
    [JSON.stringify(state)]
  );
}

// ---------------------------------------------------------------- http
function sendJson(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(typeof body === 'string' ? body : JSON.stringify(body));
}

function readRequestBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', chunk => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(new Error('Payload demasiado grande'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

async function handleRequest(req, res) {
  const urlObj = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = decodeURIComponent(urlObj.pathname);
  const wbParam = urlObj.searchParams.get('wb');
  const key = WORKBOOKS[wbParam] ? wbParam : DEFAULT_WB;

  if (pathname === '/api/posco-callback' && req.method === 'POST') {
    if (!REUNION_SYNC_TOKEN || req.headers.authorization !== `Bearer ${REUNION_SYNC_TOKEN}`) {
      return sendJson(res, 401, { error: 'Callback no autorizado.' });
    }
    try {
      const payload = JSON.parse(await readRequestBody(req, 1024 * 1024));
      await receiveGitHubPoscoCallback(payload);
      return sendJson(res, 200, { ok: true });
    } catch (error) {
      console.error('Callback POSCO falló:', error);
      return sendJson(res, 400, { error: error.message || 'Callback inválido.' });
    }
  }

  if (pathname === '/api/workbooks' && req.method === 'GET') {
    return sendJson(res, 200, Object.entries(WORKBOOKS).map(([k, c]) => ({ key: k, name: c.name })));
  }

  if (pathname === '/api/actualizar-data' && (req.method === 'GET' || req.method === 'POST')) {
    if (key !== 'reunion') return sendJson(res, 400, { error: 'Actualizar Data solo está disponible para Reunion.' });
    if (req.method === 'GET') return sendJson(res, 200, publicActualizarDataState());
    if (actualizarDataRunning) return sendJson(res, 202, publicActualizarDataState());

    actualizarDataRunning = true;
    Object.assign(actualizarDataState, {
      status: 'running',
      startedAt: new Date().toISOString(),
      finishedAt: null,
      rows: null,
      lastRow: null,
      range: null,
      dateRange: null,
      error: null
    });
    if (GITHUB_ACTIONS_TOKEN) {
      try {
        await dispatchGitHubPosco('data');
      } catch (error) {
        actualizarDataRunning = false;
        actualizarDataState.status = 'failed';
        actualizarDataState.finishedAt = new Date().toISOString();
        actualizarDataState.error = error.message;
        return sendJson(res, 502, publicActualizarDataState());
      }
      return sendJson(res, 202, publicActualizarDataState());
    }
    void runActualizarDataInBackground();
    return sendJson(res, 202, publicActualizarDataState());
  }

  if (pathname === '/api/actualizar-inventario' && (req.method === 'GET' || req.method === 'POST')) {
    if (key !== 'reunion') return sendJson(res, 400, { error: 'Actualizar Inventario solo está disponible para Reunion.' });
    if (req.method === 'GET') return sendJson(res, 200, publicActualizarInventarioState());
    if (actualizarInventarioRunning) return sendJson(res, 202, publicActualizarInventarioState());

    actualizarInventarioRunning = true;
    Object.assign(actualizarInventarioState, {
      status: 'running',
      startedAt: new Date().toISOString(),
      finishedAt: null,
      rows: null,
      lastRow: null,
      error: null
    });
    if (GITHUB_ACTIONS_TOKEN) {
      try {
        await dispatchGitHubPosco('inventario');
      } catch (error) {
        actualizarInventarioRunning = false;
        actualizarInventarioState.status = 'failed';
        actualizarInventarioState.finishedAt = new Date().toISOString();
        actualizarInventarioState.error = error.message;
        return sendJson(res, 502, publicActualizarInventarioState());
      }
      return sendJson(res, 202, publicActualizarInventarioState());
    }
    void runActualizarInventarioInBackground();
    return sendJson(res, 202, publicActualizarInventarioState());
  }

  if (pathname === '/api/workbook' && req.method === 'GET') {
    const workbook = await loadWorkbook(key);
    if (String(req.headers['accept-encoding'] || '').includes('gzip')) {
      const compressed = await compressedWorkbook(key, workbook);
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-store' });
      return res.end(compressed);
    }
    return sendJson(res, 200, workbook);
  }

  if (pathname === '/api/workbook' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req, 300 * 1024 * 1024));
      if (!parsed || typeof parsed !== 'object' || !parsed.sheets) throw new Error('bad workbook');
      await saveWholeWorkbook(key, parsed);
      scheduleCorSync();
      return sendJson(res, 200, { ok: true });
    } catch (error) {
      return sendJson(res, 400, { error: 'JSON de libro inválido' });
    }
  }

  if (pathname === '/api/workbook' && req.method === 'PATCH') {
    let parsed;
    try {
      parsed = JSON.parse(await readRequestBody(req, 300 * 1024 * 1024));
      if (!parsed || typeof parsed !== 'object' || (parsed.sheets && typeof parsed.sheets !== 'object')) throw new Error('bad patch');
    } catch (error) {
      return sendJson(res, 400, { error: 'Cambio inválido' });
    }
    await patchWorkbook(key, parsed);
    if (key === 'cosecha') scheduleCorSync();
    else if (key === 'reunion') {
      const appendSheets = new Set(['PegarData', 'Inventario']);
      const affectsAppend = [...Object.keys(parsed.sheets || {}), ...Object.keys(parsed.cellPatches || {})]
        .some(name => appendSheets.has(name));
      const reunionNamesChanged = parsed.cellPatches && parsed.cellPatches.Reunion &&
        Object.keys(parsed.cellPatches.Reunion.set || {}).some(coord => /^A\d+$/.test(coord));
      if (affectsAppend || reunionNamesChanged) scheduleAppendSync();
    }
    return sendJson(res, 200, { ok: true });
  }

  // Cor of Reunion coming from Plan de cosecha: current feed + live updates
  if (pathname === '/api/cor-feed' && req.method === 'GET') {
    if (!corFeed) await syncCor();
    return sendJson(res, 200, publicCorFeed());
  }
  // "✓ Visto": { keys: [...] } or { all: true } removes the change marks (the history stays)
  if (pathname === '/api/cor-history/seen' && req.method === 'POST') {
    let parsed;
    try { parsed = JSON.parse(await readRequestBody(req, 1024 * 1024)); } catch (e) { return sendJson(res, 400, { error: 'JSON inválido' }); }
    const h = loadCorHistory();
    const keys = parsed && parsed.all ? Object.keys(h.entries) : (Array.isArray(parsed && parsed.keys) ? parsed.keys : []);
    for (const k of keys) if (h.entries[k]) h.entries[k].unseen = false;
    saveCorHistorySoon();
    broadcastCor();
    return sendJson(res, 200, { ok: true });
  }
  // Cor typed in Reunion: { writes: [{ product, serial, value }] } goes to Plan de cosecha (finca Isabela)
  if (pathname === '/api/cor-write' && req.method === 'POST') {
    let parsed;
    try { parsed = JSON.parse(await readRequestBody(req, 1024 * 1024)); } catch (e) { return sendJson(res, 400, { error: 'JSON inválido' }); }
    const cosecha = await loadWorkbook('cosecha');
    const errors = [];
    const changed = {};
    let sheetName = null;
    for (const w of (parsed && Array.isArray(parsed.writes) ? parsed.writes : [])) {
      const r = CosechaCor.writeCor(cosecha, CosechaCor.norm(w.product), Number(w.serial), Number(w.value) || 0);
      if (r.error) { errors.push(r.error); continue; }
      sheetName = r.sheet;
      Object.assign(changed, r.cells);
    }
    if (sheetName) {
      await patchWorkbook('cosecha', {
        cellPatches: { [sheetName]: { set: changed, delete: [] } }
      });
      const msg = `event: cosecha
data: ${JSON.stringify({ sheet: sheetName, cells: changed })}

`;
      for (const client of corClients) client.write(msg);
      clearTimeout(corTimer);
      await syncCor();
    }
    return sendJson(res, 200, { ok: !errors.length, errors });
  }
  // one sheet of a book (used to refresh "Append" without reloading the whole page)
  if (pathname === '/api/sheet' && req.method === 'GET') {
    const workbook = await loadWorkbook(key);
    const sheet = workbook.sheets[urlObj.searchParams.get('name') || ''];
    if (!sheet) return sendJson(res, 404, { error: 'Hoja no encontrada' });
    return sendJson(res, 200, sheet);
  }
  if (pathname === '/api/events' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', Connection: 'keep-alive' });
    res.write(': ok\n\n');
    corClients.add(res);
    const ping = setInterval(() => res.write(': ping\n\n'), 25000);
    req.on('close', () => { clearInterval(ping); corClients.delete(res); });
    return;
  }

  if (pathname === '/api/homologations' && req.method === 'GET') {
    return sendJson(res, 200, await getHomologations(key));
  }

  if (pathname === '/api/homologations' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req, 10 * 1024 * 1024));
      if (!parsed || typeof parsed.items !== 'object' || Array.isArray(parsed.items)) throw new Error('bad shape');
      await saveHomologations(key, parsed);
      return sendJson(res, 200, { ok: true });
    } catch (error) {
      return sendJson(res, 400, { error: 'JSON de homologaciones inválido' });
    }
  }

  if (pathname === '/api/reset' && req.method === 'POST') {
    return sendJson(res, 200, await resetWorkbook(key));
  }

  // Windows user running the server (used by the "Bitacora" log, like Environ("Username"))
  if (pathname === '/api/whoami' && req.method === 'GET') {
    let user = process.env.USERNAME || process.env.USER || '';
    try { user = os.userInfo().username || user; } catch (error) { /* keep env value */ }
    return sendJson(res, 200, { user });
  }

  const requestedPath = pathname === '/' ? '/index.html' : pathname;
  const filePath = path.resolve(ROOT_DIR, `.${requestedPath}`);
  if (filePath !== ROOT_DIR && !filePath.startsWith(`${ROOT_DIR}${path.sep}`)) {
    return res.writeHead(403).end('Forbidden');
  }
  if (filePath.startsWith(`${DATA_DIR}${path.sep}`) || filePath.startsWith(`${path.join(ROOT_DIR, 'tools')}${path.sep}`)) {
    return res.writeHead(403).end('Forbidden');
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return res.end('Archivo no encontrado');
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
      'Cache-Control': 'no-store'
    });
    res.end(content);
  });
}

async function start() {
  const server = http.createServer((req, res) => {
    handleRequest(req, res).catch(error => {
      console.error('Request error:', error);
      if (!res.headersSent) sendJson(res, 500, { error: 'Error interno del servidor' });
      else res.end();
    });
  });

  server.listen(PORT, '0.0.0.0', () => {
    console.log('====================================================');
    console.log(' SISTEMA EXCEL - Reunion 123 / Plan de cosecha ACTIVO');
    console.log(` URL: http://localhost:${PORT}`);
    console.log('====================================================');
  });

  try {
    await initDatabase();
    persistenceReady = true;
    // Do not delay the web service while the one-time cell migration runs.
    void migrateAllWorkbookCellsInBackground();
    await loadWorkbook(DEFAULT_WB);
    // Cor is initialized at startup, but Append does not need to be rebuilt
    // until a source workbook changes. Avoid competing with the first edits.
    await syncCor({ rebuildAppend: false });
    console.log(' Persistencia inicializada y sincronización lista');
  } catch (error) {
    // Keep the web service available with the local snapshot. The next
    // deployment/retry can initialize Neon without leaving the page blank.
    console.error('No se pudo inicializar la persistencia; se mantiene el respaldo local:', error.message);
  }
}

start();
