const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { Pool } = require('pg');
const { runActualizarData } = require('./actualizar_data');
const { runActualizarInventario, runActualizarInventarioViernes, FRIDAY_SHEET_NAME } = require('./actualizar_inventario');
const CosechaCor = require('./cosechaCor');
const AppendBuilder = require('./appendBuilder');
const NeonStore = require('./neonStore');
const SheetDelta = require('./sheetDelta');
const gzipAsync = require('util').promisify(zlib.gzip);
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
// Neon needs SSL; the PostgreSQL installed on this same PC (tools/instalar_postgres.ps1) doesn't use it
const DATABASE_SSL = /@(localhost|127\.0\.0\.1)[:/]/.test(DATABASE_URL || '') ? false : { rejectUnauthorized: false };
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
const gzCache = {}; // wbKey -> promise of the gzipped JSON of the cached workbook (rebuilt soon after a change)
const sheetJson = {}; // wbKey -> Map(sheet name -> its JSON): only the sheets that changed are converted again
const warmTimers = {};

// ---- change log: every change of a book gets a number (rev). A page that has a copy at rev N asks
//   only for what changed after N (/api/workbook-changes) instead of downloading the whole book.
//   The log is also kept on disk (changelog/<book>.ndjson), so a restart (updates) does not make every
//   page download everything again. An entry is written to disk only once the database has it.
//   After a restart, revs the old process may have given without saving them (a crash) fall in a "gap":
//   a page that has one of those downloads the whole book. Deleting the changelog folder makes every
//   page download the whole books once (do it after changing the database by hand).
const LOG_DIR = path.join(ROOT_DIR, 'changelog');
const LOG_MAX_ENTRIES = 100000;
const LOG_KEEP_MS = 4 * 24 * 3600 * 1000; // pages keep their copy up to 4 days (app.js BOOK_MAX_AGE)
const LOG_GAP = 1000000;
const changeLog = {}; // wbKey -> { id, rev, minRev, savedRev, entries: [{ rev, t, deltas, sheets, deleted } | { rev, t, gapFrom }] }
const logFile = key => path.join(LOG_DIR, `${key}.ndjson`);
function logOf(key) {
  if (changeLog[key]) return changeLog[key];
  const log = { id: '', rev: 0, minRev: 0, savedRev: 0, entries: [] };
  try {
    const lines = fs.readFileSync(logFile(key), 'utf8').split('\n').filter(Boolean);
    const head = JSON.parse(lines.shift());
    log.id = String(head.id || '');
    log.rev = log.minRev = Number(head.minRev) || 0;
    for (const line of lines) {
      let e;
      try { e = JSON.parse(line); } catch (_) { break; } // a half-written last line (power cut)
      log.rev = e.rev;
      if (e.full) { log.entries = []; log.minRev = e.rev; } else log.entries.push(e);
    }
  } catch (_) { log.id = ''; }
  if (!log.id) { log.id = crypto.randomBytes(4).toString('hex'); log.rev = log.minRev = 0; log.entries = []; }
  // old entries (pages never keep a copy that long) and the gap of this restart
  const cutoff = Date.now() - LOG_KEEP_MS;
  while (log.entries.length && log.entries[0].t < cutoff) log.minRev = log.entries.shift().rev;
  const savedUpTo = log.rev;
  log.rev += LOG_GAP;
  log.entries.push({ rev: log.rev, t: Date.now(), gapFrom: savedUpTo });
  log.savedRev = log.rev;
  changeLog[key] = log;
  // the file is written again compact (only what is kept)
  try {
    fs.mkdirSync(LOG_DIR, { recursive: true });
    const tmp = `${logFile(key)}.tmp`;
    fs.writeFileSync(tmp, [JSON.stringify({ id: log.id, minRev: log.minRev }), ...log.entries.map(e => JSON.stringify(e))].join('\n') + '\n');
    fs.renameSync(tmp, logFile(key));
  } catch (error) { console.error(`Registro de cambios (${key}):`, error.message); }
  return log;
}
function bookRev(key) { const log = logOf(key); return `${log.id}.${log.rev}`; }
// entry: { deltas: { sheet: delta }, sheets: [whole sheets replaced], deleted: [names] } or { full: true }
function logChange(key, entry) {
  const log = logOf(key);
  log.rev++;
  if (entry.full) {
    // older copies: whole book (written to the file with the rest, once the database has it)
    log.entries = [{ rev: log.rev, t: Date.now(), full: true }];
    log.minRev = log.rev;
    if (!pool) saveLog(key, log.rev);
    return;
  }
  log.entries.push({ rev: log.rev, t: Date.now(), deltas: entry.deltas || {}, sheets: entry.sheets || [], deleted: entry.deleted || [] });
  if (log.entries.length > LOG_MAX_ENTRIES) {
    const dropped = log.entries.splice(0, LOG_MAX_ENTRIES / 4);
    log.minRev = dropped[dropped.length - 1].rev;
  }
  if (!pool) saveLog(key, log.rev); // local files: written right away
}
// the entries up to `upto` are in the database: written to the log file too
function saveLog(key, upto, extra) {
  const log = logOf(key);
  const lines = (extra || log.entries.filter(e => e.rev > log.savedRev && e.rev <= upto)).map(e => JSON.stringify(e));
  if (upto > log.savedRev) log.savedRev = upto;
  if (!lines.length) return;
  try {
    fs.mkdirSync(LOG_DIR, { recursive: true });
    fs.appendFileSync(logFile(key), lines.join('\n') + '\n');
  } catch (error) { console.error(`Registro de cambios (${key}):`, error.message); }
}
// entries that never needed a database write (nothing really changed) are written once the book is idle
setInterval(() => {
  const cutoff = Date.now() - LOG_KEEP_MS;
  for (const [key, log] of Object.entries(changeLog)) {
    while (log.entries.length && log.entries[0].t < cutoff) log.minRev = log.entries.shift().rev;
    if (!pending[key] && !flushing[key] && log.rev > log.savedRev) saveLog(key, log.rev);
  }
}, 15000).unref();
// what changed after rev `since`, merged: null when the page must download the whole book
function changesSince(key, workbook, since) {
  const [id, n] = String(since || '').split('.');
  const s = Number(n);
  const log = logOf(key);
  if (id !== log.id || !Number.isInteger(s) || s < log.minRev || s > log.rev) return null;
  const deltas = {};
  const whole = new Set();
  const deleted = new Set();
  for (const e of log.entries) {
    if (e.rev <= s) continue;
    if (e.full) return null;
    if (e.gapFrom !== undefined) { if (s > e.gapFrom) return null; continue; } // a rev the old process never saved
    for (const name of e.deleted) { deleted.add(name); whole.delete(name); delete deltas[name]; }
    for (const name of e.sheets) { whole.add(name); deleted.delete(name); delete deltas[name]; }
    for (const [name, delta] of Object.entries(e.deltas)) {
      if (whole.has(name)) continue; // the whole sheet goes anyway, as it is now
      deltas[name] = SheetDelta.mergeDelta(deltas[name] || {}, JSON.parse(JSON.stringify(delta)));
    }
  }
  const sheetData = {};
  for (const name of whole) {
    if (workbook.sheets[name]) sheetData[name] = workbook.sheets[name];
    else deleted.add(name);
  }
  for (const name of Object.keys(deltas)) if (!workbook.sheets[name]) { delete deltas[name]; deleted.add(name); }
  const { workbookName, sheetOrder, tables, pivots } = workbook;
  return { rev: bookRev(key), deltas, sheetData, deleted: [...deleted], meta: { workbookName, sheetOrder, tables, pivots } };
}

// the book changed (`names`: only those sheets, or all when omitted): the compressed copy that pages
// download is prepared again in the background, so a reload right after typing finds it ready
function bookChanged(key, names) {
  delete gzCache[key];
  const cache = sheetJson[key];
  if (cache) {
    if (names) for (const n of names) cache.delete(n);
    else cache.clear();
  }
  clearTimeout(warmTimers[key]);
  warmTimers[key] = setTimeout(() => {
    if (key === 'rvp' && appendDirty) return; // Append is rebuilt (and the copy prepared) when RvP is opened
    if (caches[key] && !gzCache[key]) compressedBook(key, caches[key]).catch(() => {});
  }, 1500);
}
function workbookJson(key, workbook) {
  const cache = sheetJson[key] || (sheetJson[key] = new Map());
  const { sheets = {}, ...meta } = workbook;
  for (const n of cache.keys()) if (!(n in sheets)) cache.delete(n);
  const parts = [];
  for (const [name, sheet] of Object.entries(sheets)) {
    let json = cache.get(name);
    if (json === undefined) { json = JSON.stringify(sheet); cache.set(name, json); }
    parts.push(`${JSON.stringify(name)}:${json}`);
  }
  const head = JSON.stringify(meta).slice(0, -1);
  return `${head}${head.length > 1 ? ',' : ''}"sheets":{${parts.join(',')}}}`;
}
function compressedBook(key, workbook) {
  if (!gzCache[key]) {
    // fast level: it is redone after every change. The ETag (hash of the content) lets a page that
    // already has this exact copy skip the download (304)
    const rev = persistenceReady ? bookRev(key) : ''; // the rev of exactly this content
    const job = gzipAsync(workbookJson(key, workbook), { level: 1 })
      .then(body => ({ body, rev, etag: `"${crypto.createHash('md5').update(body).digest('hex')}"` }));
    gzCache[key] = job;
    job.catch(() => { if (gzCache[key] === job) delete gzCache[key]; });
  }
  return gzCache[key];
}
let actualizarDataRunning = false;
let actualizarInventarioRunning = false;
let actualizarInventarioViernesRunning = false;
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

// "Actualizar Inventario Viernes": same report, only for the Append (INVENTARIO VIERNES)
const actualizarInventarioViernesState = { status: 'idle', startedAt: null, finishedAt: null, rows: null, lastRow: null, error: null };
function publicActualizarInventarioViernesState() {
  return { ...actualizarInventarioViernesState, ok: actualizarInventarioViernesState.status === 'succeeded' };
}

function resetWorkbookCache(key) {
  logChange(key, { full: true });
  delete caches[key];
  bookChanged(key);
  delete stored[key];
  delete pending[key];
}

function stateForOperation(operation) {
  return operation === 'inventario' ? actualizarInventarioState : actualizarDataState;
}

async function dispatchGitHubPosco(operation) {
  if (!GITHUB_ACTIONS_TOKEN) return false;
  await flushAll(); // the runner reads Neon: give it the latest edits
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
    await flushBook('reunion');
    resetWorkbookCache('reunion');
    await syncCor();
    broadcastBook('reunion', 'reload', {});
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
  if (pool) {
    // Neon: grouped, at most one write every 15 s
    corHistoryTimer = setTimeout(() => {
      pool.query(
        `INSERT INTO app_state (key, value) VALUES ('cor_history', $1::jsonb)
         ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()`,
        [JSON.stringify(corHistory)]
      ).catch(err => console.error('Historial Cor:', err.message));
    }, 15000);
    return;
  }
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
// live sync: `event` to every page with book `key` open, except the one that made the change
// every event says which book it belongs to (wb): a page also follows the books it opened earlier
// (?also=...) to keep them up to date in the background
function broadcastBook(key, event, data, exceptClient) {
  const msg = `event: ${event}\ndata: ${JSON.stringify({ ...data, wb: key })}\n\n`;
  for (const res of corClients) {
    if (res.wbKey !== key && !(res.alsoKeys && res.alsoKeys.has(key))) continue;
    if (!exceptClient || res.clientId !== exceptClient) res.write(msg);
  }
}
function broadcastCor() {
  const msg = `event: cor\ndata: ${JSON.stringify(publicCorFeed())}\n\n`;
  for (const res of corClients) res.write(msg);
}
function scheduleCorSync(delay = 150) {
  clearTimeout(corTimer);
  corTimer = setTimeout(() => { syncCor().catch(err => console.error('Sync Cor:', err)); }, delay);
}
async function syncCor() {
  const cosecha = await loadWorkbook('cosecha');
  const reunion = await loadWorkbook('reunion');
  const feed = CosechaCor.computeFeed(cosecha);
  const coords = [];
  const changed = CosechaCor.applyFeed(reunion, feed, coords);
  if (changed) markCellsChanged('reunion', 'Reunion', coords);
  const feedChanged = !corFeed || corFeed.sig !== feed.sig;
  corFeed = feed;
  if (CosechaCor.trackChanges(loadCorHistory(), feed, new Date().toISOString())) saveCorHistorySoon();
  if (changed || feedChanged) broadcastCor();
  appendDirty = true;
  if (rvpViewers()) scheduleAppend();
}

// "Append" is heavy to build (~100,000 cells): it is rebuilt only while someone has
// Requerimiento Vs Proyeccion open (a few seconds after the last change) or when it is opened
let appendDirty = true;
let appendTimer = null;
let appendBuilding = null;
function rvpViewers() {
  let n = 0;
  for (const res of corClients) if (res.wbKey === 'rvp') n++;
  return n;
}
function scheduleAppend(delay = 5000) {
  clearTimeout(appendTimer);
  appendTimer = setTimeout(() => { ensureAppend().catch(err => console.error('Append:', err)); }, delay);
}
async function ensureAppend() {
  if (appendBuilding) return appendBuilding;
  if (!appendDirty) return;
  appendDirty = false;
  appendBuilding = (async () => {
    await syncAppend(await loadWorkbook('reunion'), await loadWorkbook('cosecha'));
  })().finally(() => { appendBuilding = null; });
  return appendBuilding;
}

// Requerimiento Vs Proyeccion: sheet "Append" rebuilt from PegarData + Plan de cosecha
let appendSig = null;
async function syncAppend(reunion, cosecha) {
  const rvp = await loadWorkbook('rvp');
  const { sheet } = AppendBuilder.buildAppendSheet(reunion, cosecha, rvp);
  const md5 = obj => crypto.createHash('md5').update(JSON.stringify(obj)).digest('hex');
  const sig = md5(sheet.cells);
  if (sig === appendSig && rvp.sheets.Append) return; // nothing new
  const order = rvp.sheetOrder || [];
  const meta = order.includes('Append') ? {} : { sheetOrder: [...order, 'Append'] };
  const prev = rvp.sheets.Append;
  if (prev && prev.colWidths) sheet.colWidths = { ...sheet.colWidths, ...prev.colWidths }; // keep widths set by hand
  // only the cells of Append that really changed (pages then download only those, not ~3 MB)
  const diff = prev && !meta.sheetOrder ? SheetDelta.diffSheet(SheetDelta.snapshotOf(prev), sheet) : undefined;
  if (diff === null) { appendSig = sig; return; } // same content
  if (diff && diff.size < 20000) await patchWorkbook('rvp', { deltas: { Append: diff.delta } });
  else await patchWorkbook('rvp', { meta, sheets: { Append: sheet } });
  appendSig = sig; // after the patch: patchWorkbook clears it (it treats Append edits as made by hand)
  const msg = `event: append\ndata: ${JSON.stringify({ sig })}\n\n`;
  for (const res of corClients) res.write(msg);
}

async function runActualizarDataInBackground() {
  try {
    console.log(' Iniciando Actualizar Data local: POSCO -> PegarData...');
    const currentWorkbook = await loadWorkbook('reunion');
    const result = await runActualizarData(currentWorkbook);
    await saveWholeWorkbook('reunion', result.workbook);
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

async function runActualizarInventarioViernesInBackground() {
  const state = actualizarInventarioViernesState;
  try {
    console.log(' Iniciando Actualizar Inventario Viernes: POSCO -> Inventario Viernes (solo Append)...');
    const result = await runActualizarInventarioViernes(await loadWorkbook('reunion'));
    // its week (Sat..Fri) is the week of the day it ran, like Actualizar Inventario
    result.workbook.sheets[FRIDAY_SHEET_NAME].updatedAt = new Date().toISOString();
    await saveWholeWorkbook('reunion', result.workbook);
    scheduleCorSync(); // marks the Append to be rebuilt
    Object.assign(state, { status: 'succeeded', finishedAt: new Date().toISOString(), rows: result.rows, lastRow: result.lastRow, error: null });
    console.log(` Actualizar Inventario Viernes completado: ${result.rows} filas.`);
  } catch (error) {
    Object.assign(state, { status: 'failed', finishedAt: new Date().toISOString(), error: error.message || 'No se pudo actualizar Inventario Viernes.' });
    console.error(' Actualizar Inventario Viernes falló:', error);
  } finally {
    actualizarInventarioViernesRunning = false;
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

// ---------------------------------------------------------------- database (Neon, see neonStore.js)
//   Consumption is kept low: the books live in memory, so the database is only read once at start;
//   a change writes only the cells that changed, several edits within a few seconds go in one write,
//   and data rebuilt by the server (the cells of "Append") is not stored at all.
const FLUSH_DEBOUNCE_MS = 2000;  // write when nobody has edited for 2 s...
const FLUSH_MAX_WAIT_MS = 10000; // ...or at most 10 s after the first pending change
const COMPACT_AFTER = 150;       // deltas of one sheet before it is written whole again
const BACKUP_EVERY_MS = 24 * 60 * 60 * 1000;
const stored = {};  // wbKey -> { sheets: { name: { known, knownProps } }, meta: JSON, deltaCounts: {} } (what Neon has)
const pending = {}; // wbKey -> { deltas: { name: delta }, snapshots: Set, deleted: Set, meta: bool, since }
const flushTimers = {};
let lastBackupAt = 0;

// what is stored of a sheet: "Append" is rebuilt from PegarData + Plan de cosecha, so only its settings
function persistForm(key, name, sheet) {
  if (key === 'rvp' && name === 'Append' && sheet) return { ...sheet, cells: {} };
  return sheet;
}
function persistBook(key, workbook) {
  const { meta, sheets } = splitWorkbook(workbook);
  const out = {};
  for (const [name, sheet] of Object.entries(sheets)) out[name] = persistForm(key, name, sheet);
  return { meta, sheets: out };
}
function rememberStored(key, meta, sheets, deltaCounts = {}) {
  const st = { sheets: {}, meta: JSON.stringify(meta), deltaCounts };
  stored[key] = st;
  for (const name of Object.keys(sheets)) learnSheet(key, name);
}

// what Neon has of a sheet: only its settings (colWidths, merges...). Its cells are not copied in memory
// (Render's free plan has 512 MB): the page says which cells changed and those are written as they are now;
// a whole sheet replaced without saying which cells changed is written whole ("partial" snapshot)
function learnSheet(key, name) {
  const st = stored[key];
  const sheet = persistForm(key, name, caches[key] ? caches[key].sheets[name] : null);
  if (!st || !sheet) return;
  st.sheets[name] = { known: new Map(), knownProps: SheetDelta.propsOf(sheet), partial: true };
}
function newPending() {
  // dirty: { name: 'all' | Set of touched cells }, compared with what Neon has when writing
  return { dirty: {}, deltas: {}, snapshots: new Set(), deleted: new Set(), meta: false, since: Date.now() };
}

async function initDatabase() {
  if (!DATABASE_URL) return;

  pool = new Pool({
    connectionString: DATABASE_URL,
    ssl: DATABASE_SSL,
    max: 3,
    idleTimeoutMillis: 10000, // idle connections close quickly so Neon can go to sleep
    connectionTimeoutMillis: 15000,
    allowExitOnIdle: true
  });
  pool.on('error', err => console.error('Neon:', err.message));

  await NeonStore.ensureSharedTables(pool);
  for (const key of Object.keys(WORKBOOKS)) {
    const p = cfg(key).prefix;
    await NeonStore.ensureTables(pool, p);
    const existing = await pool.query(`SELECT 1 FROM ${p}_workbook_meta WHERE id = 1`);
    if (existing.rowCount === 0) {
      // first start against an empty database: upload the local data once
      const local = readJson(cfg(key).state, null) || initialWorkbook(key);
      const { meta, sheets } = persistBook(key, local);
      await NeonStore.writeWholeBook(pool, p, meta, sheets);
      console.log(` Neon: ${cfg(key).name} subido desde los datos locales (${Object.keys(sheets).length} hojas)`);
    }
    await pool.query(
      `INSERT INTO ${p}_homologations (id, state) VALUES (1, $1::jsonb) ON CONFLICT (id) DO NOTHING`,
      [JSON.stringify(initialHomologations(key))]
    );
  }
  const corRow = await pool.query(`SELECT value FROM app_state WHERE key = 'cor_history'`);
  if (corRow.rows[0]) corHistory = corRow.rows[0].value;
  else { loadCorHistory(); saveCorHistorySoon(); }
  const last = await pool.query('SELECT MAX(created_at) AS at FROM app_backups');
  lastBackupAt = last.rows[0].at ? new Date(last.rows[0].at).getTime() : 0;

  console.log(` Persistencia activa: PostgreSQL ${DATABASE_SSL ? '(Neon, en la nube)' : '(en esta PC)'}`);
}

// ---------------------------------------------------------------- workbook persistence
const loadingBooks = {};
// every sheet the book has stays in its tab order: a page with an old copy of the book sends an
// order without the tabs created after it was loaded, which would hide them (they are only removed
// from the order when the sheet itself is deleted)
function keepEveryTab(workbook) {
  if (!workbook || !workbook.sheets) return;
  const order = Array.isArray(workbook.sheetOrder) ? workbook.sheetOrder.filter(n => workbook.sheets[n]) : [];
  for (const n of Object.keys(workbook.sheets)) if (!order.includes(n)) order.push(n);
  workbook.sheetOrder = order;
}

async function loadWorkbook(key) {
  if (caches[key] && persistenceReady) return caches[key];
  if (!pool || !persistenceReady) {
    const localWorkbook = caches[key] || readJson(cfg(key).state, null) || initialWorkbook(key);
    if (persistenceReady) caches[key] = localWorkbook;
    return localWorkbook;
  }
  if (!loadingBooks[key]) {
    loadingBooks[key] = (async () => {
      const book = await NeonStore.loadBook(pool, cfg(key).prefix);
      const meta = book ? book.meta : splitWorkbook(initialWorkbook(key)).meta;
      const sheets = book ? book.sheets : {};
      caches[key] = { ...meta, sheets };
      keepEveryTab(caches[key]);
      rememberStored(key, meta, sheets, book ? book.deltaCounts : {});
      return caches[key];
    })().finally(() => { delete loadingBooks[key]; });
  }
  return loadingBooks[key];
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

// the in-memory book changed: remember which sheets (and cells, when known) to compare at write time.
// `touched` = { name: [coords] } for sheets where only those cells can have changed.
function queueChanges(key, names, deletedNames, touched) {
  const st = stored[key];
  if (!st) return;
  const pd = pending[key] || (pending[key] = newPending());
  for (const name of deletedNames) {
    delete pd.dirty[name];
    if (!st.sheets[name]) continue;
    delete st.sheets[name];
    delete pd.deltas[name];
    pd.snapshots.delete(name);
    pd.deleted.add(name);
  }
  for (const name of names) {
    pd.deleted.delete(name);
    const coords = touched && touched[name];
    const cur = pd.dirty[name];
    if (cur === 'all' || !coords) { pd.dirty[name] = 'all'; continue; }
    const set = cur || new Set();
    for (const c of coords) set.add(c);
    pd.dirty[name] = set;
  }
  scheduleFlush(key);
}

// turns the dirty marks into deltas / full copies (what differs from what Neon has)
function collectChanges(key, pd) {
  const st = stored[key];
  const workbook = caches[key];
  for (const [name, mark] of Object.entries(pd.dirty)) {
    const sheet = persistForm(key, name, workbook.sheets[name]);
    if (!sheet) continue;
    const known = st.sheets[name];
    // a new sheet, or a whole sheet that can't be compared yet: written whole
    if (!known || (mark === 'all' && known.partial)) { pd.snapshots.add(name); delete pd.deltas[name]; continue; }
    const diff = SheetDelta.diffSheet(known, sheet, mark === 'all' ? null : mark);
    if (!diff || pd.snapshots.has(name)) continue; // nothing new, or the whole sheet is going anyway
    // a big change (e.g. Actualizar Data) is cheaper as a full copy than as a delta
    if (diff.size > Math.max(3000, known.known.size * 0.3)) { delete pd.deltas[name]; pd.snapshots.add(name); continue; }
    pd.deltas[name] = pd.deltas[name] ? SheetDelta.mergeDelta(pd.deltas[name], diff.delta) : diff.delta;
    known.known.clear();
  }
  pd.dirty = {};
  const metaJson = JSON.stringify(splitWorkbook(workbook).meta);
  if (metaJson !== st.meta) { st.meta = metaJson; pd.meta = true; }
  return Object.keys(pd.deltas).length || pd.snapshots.size || pd.deleted.size || pd.meta;
}

function scheduleFlush(key, delay = FLUSH_DEBOUNCE_MS) {
  const pd = pending[key];
  if (!pd) return;
  clearTimeout(flushTimers[key]);
  const wait = Math.max(0, Math.min(delay, pd.since + FLUSH_MAX_WAIT_MS - Date.now()));
  flushTimers[key] = setTimeout(() => { flushBook(key).catch(err => console.error(`Neon (${key}):`, err.message)); }, wait);
}

const flushing = {};
async function flushBook(key) {
  while (flushing[key]) await flushing[key].catch(() => {});
  const pd = pending[key];
  if (!pd || !pool) return;
  delete pending[key];
  clearTimeout(flushTimers[key]);
  const upto = logOf(key).rev; // every change up to here is in pd (logged and queued together)
  if (!collectChanges(key, pd)) { saveLog(key, upto); return; } // nothing really changed: Neon stays asleep
  const p = cfg(key).prefix;
  const st = stored[key];
  const workbook = caches[key];
  flushing[key] = NeonStore.transaction(pool, async client => {
    for (const name of pd.deleted) {
      await client.query(`DELETE FROM ${p}_workbook_sheets WHERE name = $1`, [name]);
      await client.query(`DELETE FROM ${p}_sheet_deltas WHERE name = $1`, [name]);
    }
    const snaps = {};
    for (const name of pd.snapshots) if (workbook.sheets[name]) snaps[name] = persistForm(key, name, workbook.sheets[name]);
    await NeonStore.writeSnapshots(client, p, snaps);
    for (const [name, delta] of Object.entries(pd.deltas)) {
      await client.query(`INSERT INTO ${p}_sheet_deltas (name, delta) VALUES ($1, $2::jsonb)`, [name, JSON.stringify(delta)]);
    }
    if (pd.meta) await NeonStore.writeMeta(client, p, splitWorkbook(workbook).meta);
  });
  try {
    await flushing[key];
  } catch (error) {
    // put it back (newer changes on top) and try again a bit later: Neon waking up, network...
    const newer = pending[key];
    if (newer) {
      for (const [name, delta] of Object.entries(newer.deltas)) pd.deltas[name] = pd.deltas[name] ? NeonStore.mergeDelta(pd.deltas[name], delta) : delta;
      newer.snapshots.forEach(n => { pd.snapshots.add(n); delete pd.deltas[n]; });
      Object.assign(pd.dirty, newer.dirty);
      newer.deleted.forEach(n => pd.deleted.add(n));
      pd.meta = pd.meta || newer.meta;
    }
    pending[key] = pd;
    scheduleFlush(key, 15000);
    throw error;
  } finally {
    delete flushing[key];
  }
  saveLog(key, upto);
  for (const name of pd.deleted) delete st.deltaCounts[name];
  for (const name of pd.snapshots) { st.deltaCounts[name] = 0; learnSheet(key, name); }
  for (const name of Object.keys(pd.deltas)) st.deltaCounts[name] = (st.deltaCounts[name] || 0) + 1;
  // sheets with many deltas are written whole again, so loading stays quick
  const heavy = Object.keys(st.deltaCounts).filter(n => st.deltaCounts[n] >= COMPACT_AFTER && workbook.sheets[n]);
  if (heavy.length) {
    const next = pending[key] || (pending[key] = newPending());
    for (const n of heavy) { delete next.deltas[n]; next.snapshots.add(n); }
    scheduleFlush(key);
  }
  backupIfDue().catch(err => console.error('Respaldo:', err.message));
}

async function flushAll() {
  for (const key of Object.keys(pending)) await flushBook(key).catch(err => console.error(`Neon (${key}):`, err.message));
}

// once a day (only while the database is awake for a write anyway): gzip copy of every book, last 7 kept
async function backupIfDue() {
  if (!pool || Date.now() - lastBackupAt < BACKUP_EVERY_MS) return;
  lastBackupAt = Date.now();
  let total = 0;
  for (const key of Object.keys(WORKBOOKS)) {
    if (caches[key]) total += await NeonStore.writeBackup(pool, key, caches[key]);
  }
  console.log(` Respaldo diario guardado en Neon (${(total / 1048576).toFixed(1)} MB comprimido)`);
}

// a whole new copy of a book (Actualizar Data / Inventario): which sheets really changed. Small changes
// go as cells; big ones as whole sheets (fetched compressed). null = can't tell (whole book).
const LIVE_MAX_CELLS = 3000;
function bookDiff(key, before, after) {
  if (!before || !before.sheets) return null;
  const names = Object.keys(after.sheets || {});
  const deltas = {};
  const sheets = [];
  const cache = sheetJson[key];
  for (const name of names) {
    const a = before.sheets[name];
    const b = after.sheets[name];
    if (!a) { sheets.push(name); continue; }
    if (((cache && cache.get(name)) || JSON.stringify(a)) === JSON.stringify(b)) continue;
    const diff = SheetDelta.diffSheet(SheetDelta.snapshotOf(a), b);
    if (!diff) continue;
    if (diff.size > LIVE_MAX_CELLS) sheets.push(name); else deltas[name] = diff.delta;
  }
  const deleted = Object.keys(before.sheets).filter(n => !(n in (after.sheets || {})));
  return { deltas, sheets, deleted };
}

async function saveWholeWorkbook(key, workbook) {
  const previous = caches[key];
  const before = previous ? Object.keys(previous.sheets || {}) : [];
  const change = persistenceReady ? bookDiff(key, previous, workbook) : null;
  caches[key] = workbook;
  if (change) {
    // like an edit from a page: only what changed is logged, sent live and written
    const changedNames = [...Object.keys(change.deltas), ...change.sheets];
    logChange(key, change);
    bookChanged(key, [...changedNames, ...change.deleted]);
    const { workbookName, sheetOrder, tables, pivots } = workbook;
    broadcastBook(key, 'live', { deltas: change.deltas, sheets: change.sheets, deleted: change.deleted, meta: { workbookName, sheetOrder, tables, pivots } });
    if (!pool) return persistFileSoon(key);
    queueChanges(key, changedNames, change.deleted);
    return;
  }
  logChange(key, { full: true });
  bookChanged(key);
  broadcastBook(key, 'reload', {});
  if (!pool) return persistFileSoon(key);
  const names = Object.keys(workbook.sheets || {});
  queueChanges(key, names, before.filter(n => !names.includes(n)));
}

async function patchWorkbook(key, patch) {
  const workbook = await loadWorkbook(key);
  const meta = patch.meta || {};
  for (const k of ['workbookName', 'sheetOrder', 'tables', 'activeSheet', 'pivots']) {
    if (meta[k] !== undefined) workbook[k] = meta[k];
  }
  const changed = patch.sheets || {};
  const deleted = Array.isArray(patch.deleted) ? patch.deleted : [];
  for (const [name, data] of Object.entries(changed)) workbook.sheets[name] = data;
  for (const name of deleted) delete workbook.sheets[name];
  keepEveryTab(workbook);
  // deltas (only the cells that changed); a sheet the server doesn't have is reported back
  const missing = [];
  const touched = {};
  for (const [name, delta] of Object.entries(patch.deltas || {})) {
    if (!workbook.sheets[name]) { missing.push(name); continue; }
    SheetDelta.applyDelta(workbook.sheets[name], delta);
    touched[name] = Object.keys(delta.cells || {});
  }
  if (key === 'rvp' && ('Append' in changed || 'Append' in touched)) appendSig = null; // formulas edited by hand
  const logged = {};
  for (const [name, delta] of Object.entries(patch.deltas || {})) if (!missing.includes(name)) logged[name] = delta;
  logChange(key, { deltas: logged, sheets: Object.keys(changed), deleted });
  bookChanged(key, [...Object.keys(changed), ...Object.keys(touched), ...deleted]);

  if (!pool) { persistFileSoon(key); return missing; }
  queueChanges(key, [...Object.keys(changed), ...Object.keys(touched)], deleted, touched);
  return missing;
}

// the server changed some cells of a sheet itself (Cor, Plan de cosecha written from Reunion)
function markCellsChanged(key, name, coords) {
  bookChanged(key, [name]);
  const sheet = caches[key] && caches[key].sheets[name];
  if (sheet && coords.length) {
    const cells = {};
    for (const c of coords) cells[c] = sheet.cells[c] || null;
    logChange(key, { deltas: { [name]: { cells: JSON.parse(JSON.stringify(cells)) } } });
    broadcastBook(key, 'live', { deltas: { [name]: { cells } }, sheets: [], deleted: [] });
  } else {
    logChange(key, { sheets: [name] });
  }
  if (!pool) return persistFileSoon(key);
  queueChanges(key, [name], [], { [name]: coords });
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

  // Progressive read-only stream: sends the workbook structure first, then one complete sheet
  // at a time. It never changes persistence; it only lets the browser render while the response
  // is still arriving.
  if (pathname === '/api/workbook-stream' && req.method === 'GET') {
    if (key === 'rvp' && persistenceReady) await ensureAppend();
    const workbook = await loadWorkbook(key);
    const streamRev = persistenceReady ? bookRev(key) : '';
    const orderedNames = [
      ...(Array.isArray(workbook.sheetOrder) ? workbook.sheetOrder : []),
      ...Object.keys(workbook.sheets || {})
    ].filter((name, index, all) => workbook.sheets[name] && all.indexOf(name) === index);
    const skeletons = {};
    for (const name of orderedNames) {
      const { cells, ...header } = workbook.sheets[name] || {};
      skeletons[name] = { ...header, cells: {} };
    }
    const meta = { ...workbook, sheets: skeletons };
    const names = orderedNames.slice();
    const active = names.indexOf(workbook.activeSheet);
    if (active > 0) names.unshift(...names.splice(active, 1));

    const acceptsGzip = String(req.headers['accept-encoding'] || '').includes('gzip');
    const gzip = acceptsGzip ? zlib.createGzip({ level: 1 }) : null;
    if (gzip) {
      res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-store' });
      gzip.pipe(res);
    } else {
      res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-store' });
    }
    const writeLine = async value => {
      const line = `${JSON.stringify(value)}\n`;
      const target = gzip || res;
      if (!target.write(line)) await new Promise(resolve => target.once('drain', resolve));
    };
    try {
      await writeLine({ type: 'meta', workbook: meta, totalSheets: names.length, rev: streamRev });
      const cellChunkSize = 3000;
      for (const name of names) {
        const source = workbook.sheets[name] || {};
        const { cells = {}, ...header } = source;
        const entries = Object.entries(cells);
        await writeLine({ type: 'sheet-start', name, sheet: { ...header, cells: {} }, totalCells: entries.length });
        for (let i = 0; i < entries.length; i += cellChunkSize) {
          await writeLine({ type: 'sheet-cells', name, cells: Object.fromEntries(entries.slice(i, i + cellChunkSize)) });
        }
        await writeLine({ type: 'sheet-done', name });
      }
      await writeLine({ type: 'done' });
      if (gzip) await new Promise(resolve => { gzip.once('end', resolve); gzip.end(); });
      else res.end();
    } catch (error) {
      if (!res.writableEnded) res.destroy(error);
    }
    return;
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

  if (pathname === '/api/actualizar-inventario-viernes' && (req.method === 'GET' || req.method === 'POST')) {
    if (key !== 'reunion') return sendJson(res, 400, { error: 'Actualizar Inventario Viernes solo está disponible para Reunion.' });
    if (req.method === 'GET' || actualizarInventarioViernesRunning) return sendJson(res, req.method === 'GET' ? 200 : 202, publicActualizarInventarioViernesState());
    actualizarInventarioViernesRunning = true;
    Object.assign(actualizarInventarioViernesState, { status: 'running', startedAt: new Date().toISOString(), finishedAt: null, rows: null, lastRow: null, error: null });
    void runActualizarInventarioViernesInBackground();
    return sendJson(res, 202, publicActualizarInventarioViernesState());
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

  // only what changed after the page's copy (?since=rev); { full: true } = download the whole book
  if (pathname === '/api/workbook-changes' && req.method === 'GET') {
    if (!persistenceReady) return sendJson(res, 200, { full: true });
    if (key === 'rvp') await ensureAppend(); // Append up to date before it is compared
    const workbook = await loadWorkbook(key);
    const changes = changesSince(key, workbook, urlObj.searchParams.get('since'));
    const json = JSON.stringify(changes || { full: true });
    if (json.length > 2048 && String(req.headers['accept-encoding'] || '').includes('gzip')) {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-store' });
      return res.end(await gzipAsync(json, { level: 1 }));
    }
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
    return res.end(json);
  }

  if (pathname === '/api/workbook' && req.method === 'GET') {
    if (key === 'rvp' && persistenceReady) await ensureAppend(); // Append up to date before it is sent
    const workbook = await loadWorkbook(key);
    if (String(req.headers['accept-encoding'] || '').includes('gzip')) {
      // compressed off the main thread, and kept until the book changes
      const job = await compressedBook(key, workbook);
      const { body, etag } = job;
      if (req.headers['if-none-match'] === etag) {
        res.writeHead(304, { ETag: etag, 'Cache-Control': 'no-cache' });
        return res.end();
      }
      const head = { 'Content-Type': 'application/json; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-cache', ETag: etag };
      if (job.rev) head['X-Book-Rev'] = job.rev;
      res.writeHead(200, head);
      return res.end(body);
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
    // a page that mixed up books (another book's name) must not write into this one
    const sentName = parsed.meta && parsed.meta.workbookName;
    if (sentName && sentName !== cfg(key).name && Object.values(WORKBOOKS).some(w => w.name === sentName)) {
      console.warn(` Guardado rechazado: llegó "${sentName}" para el libro ${key}`);
      return sendJson(res, 409, { error: 'Libro equivocado', wrongBook: true });
    }
    const missing = await patchWorkbook(key, parsed);
    scheduleCorSync();
    // the other pages get the same change: cells as deltas, whole sheets to be fetched, sheet order / pivots
    const from = String(req.headers['x-client-id'] || '');
    const deltas = {};
    for (const [name, delta] of Object.entries(parsed.deltas || {})) if (!missing.includes(name)) deltas[name] = delta;
    const live = { deltas, sheets: Object.keys(parsed.sheets || {}), deleted: parsed.deleted || [] };
    if (parsed.meta) {
      const { activeSheet, ...meta } = parsed.meta; // each page keeps its own active sheet
      live.meta = meta;
    }
    if (Object.keys(deltas).length || live.sheets.length || live.deleted.length) broadcastBook(key, 'live', live, from);
    if (missing && missing.length) return sendJson(res, 409, { error: 'Hojas desconocidas', missing });
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
      markCellsChanged('cosecha', sheetName, Object.keys(changed));
      const msg = `event: cosecha
data: ${JSON.stringify({ sheet: sheetName, cells: changed })}

`;
      for (const client of corClients) client.write(msg);
      clearTimeout(corTimer);
      await syncCor();
    }
    return sendJson(res, 200, { ok: !errors.length, errors });
  }
  // download a copy of the whole book (gzip JSON), straight from memory: costs nothing in Neon
  if (pathname === '/api/backup' && req.method === 'GET') {
    const workbook = await loadWorkbook(key);
    const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
    res.writeHead(200, {
      'Content-Type': 'application/gzip',
      'Content-Disposition': `attachment; filename="respaldo-${key}-${stamp}.json.gz"`,
      'Cache-Control': 'no-store'
    });
    return res.end(await gzipAsync(JSON.stringify(workbook), { level: 6 }));
  }
  // one sheet of a book (used to refresh "Append" without reloading the whole page)
  if (pathname === '/api/sheet' && req.method === 'GET') {
    if (key === 'rvp' && urlObj.searchParams.get('name') === 'Append' && persistenceReady) await ensureAppend();
    const workbook = await loadWorkbook(key);
    const sheet = workbook.sheets[urlObj.searchParams.get('name') || ''];
    if (!sheet) return sendJson(res, 404, { error: 'Hoja no encontrada' });
    const json = JSON.stringify(sheet);
    if (json.length > 2048 && String(req.headers['accept-encoding'] || '').includes('gzip')) {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-store' });
      return res.end(await gzipAsync(json, { level: 1 }));
    }
    return sendJson(res, 200, json);
  }
  if (pathname === '/api/events' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', Connection: 'keep-alive' });
    res.write(': ok\n\n');
    res.wbKey = key;
    res.alsoKeys = new Set(String(urlObj.searchParams.get('also') || '').split(',').filter(k => WORKBOOKS[k] && k !== key));
    res.clientId = urlObj.searchParams.get('client') || '';
    corClients.add(res);
    if (key === 'rvp' && appendDirty) scheduleAppend(500);
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
    for (const key of Object.keys(WORKBOOKS)) await loadWorkbook(key); // the only full read of Neon
    await syncCor();
    console.log(' Persistencia inicializada y sincronización lista');
  } catch (error) {
    // Keep the web service available with the local snapshot. The next
    // deployment/retry can initialize Neon without leaving the page blank.
    console.error('No se pudo inicializar la persistencia; se mantiene el respaldo local:', error.message);
  }
}

// Render stops the service with SIGTERM (deploys, sleeping): write what is still pending first
let stopping = false;
function stopCleanly() {
  if (stopping) return;
  stopping = true;
  const pendingSaves = Object.keys(pending).length;
  if (pendingSaves) console.log(` Guardando ${pendingSaves} libro(s) pendientes en Neon antes de salir...`);
  Promise.race([flushAll(), new Promise(r => setTimeout(r, 8000))])
    .then(() => { for (const key of Object.keys(changeLog)) if (!pending[key]) saveLog(key, changeLog[key].rev); })
    .finally(() => process.exit(0));
}
for (const signal of ['SIGTERM', 'SIGINT']) process.on(signal, stopCleanly);
// Windows has no SIGTERM for other programs: tools/servidor.ps1 (updates, "Apagar sistema") creates this file instead
const STOP_FLAG = path.join(__dirname, 'data', '.detener');
try { fs.unlinkSync(STOP_FLAG); } catch (_) {}
setInterval(() => { if (fs.existsSync(STOP_FLAG)) { try { fs.unlinkSync(STOP_FLAG); } catch (_) {} stopCleanly(); } }, 1000).unref();

start();
