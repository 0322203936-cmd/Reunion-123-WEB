const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const { Pool } = require('pg');
const { runActualizarData } = require('./actualizar_data');

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
const caches = {}; // wbKey -> full workbook kept in memory; PATCH merges into it
const gzCache = {}; // wbKey -> gzipped JSON of the cached workbook (dropped on every change)
let actualizarDataRunning = false;

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

// ---------------------------------------------------------------- database
async function initDatabase() {
  if (!DATABASE_URL) return;

  pool = new Pool({
    connectionString: DATABASE_URL,
    ssl: { rejectUnauthorized: false },
    max: 5,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 10000
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

async function writeWholeWorkbookToDb(key, workbook) {
  const p = cfg(key).prefix;
  const { meta, sheets } = splitWorkbook(workbook);
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM ${p}_workbook_sheets`);
    for (const [name, data] of Object.entries(sheets)) {
      await client.query(
        `INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)`,
        [name, JSON.stringify(data)]
      );
    }
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
  if (caches[key]) return caches[key];
  if (!pool) {
    caches[key] = readJson(cfg(key).state, null) || initialWorkbook(key);
    return caches[key];
  }
  const p = cfg(key).prefix;
  const metaRes = await pool.query(`SELECT meta FROM ${p}_workbook_meta WHERE id = 1`);
  const sheetRes = await pool.query(`SELECT name, data FROM ${p}_workbook_sheets`);
  const meta = metaRes.rows[0]?.meta || splitWorkbook(initialWorkbook(key)).meta;
  const sheets = {};
  for (const row of sheetRes.rows) sheets[row.name] = row.data;
  caches[key] = { ...meta, sheets };
  return caches[key];
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
  delete gzCache[key];
  if (!pool) return persistFileSoon(key);
  await writeWholeWorkbookToDb(key, workbook);
}

async function patchWorkbook(key, patch) {
  const workbook = await loadWorkbook(key);
  delete gzCache[key];
  const meta = patch.meta || {};
  for (const k of ['workbookName', 'sheetOrder', 'tables', 'activeSheet', 'pivots']) {
    if (meta[k] !== undefined) workbook[k] = meta[k];
  }
  const changed = patch.sheets || {};
  const deleted = Array.isArray(patch.deleted) ? patch.deleted : [];
  for (const [name, data] of Object.entries(changed)) workbook.sheets[name] = data;
  for (const name of deleted) delete workbook.sheets[name];

  if (!pool) return persistFileSoon(key);

  const p = cfg(key).prefix;
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (const [name, data] of Object.entries(changed)) {
      await client.query(
        `INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)
         ON CONFLICT (name) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()`,
        [name, JSON.stringify(data)]
      );
    }
    for (const name of deleted) {
      await client.query(`DELETE FROM ${p}_workbook_sheets WHERE name = $1`, [name]);
    }
    const { meta: fullMeta } = splitWorkbook(workbook);
    await client.query(
      `INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb)
       ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`,
      [JSON.stringify(fullMeta)]
    );
    await client.query('COMMIT');
  } catch (error) {
    await client.query('ROLLBACK');
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

  if (pathname === '/api/workbooks' && req.method === 'GET') {
    return sendJson(res, 200, Object.entries(WORKBOOKS).map(([k, c]) => ({ key: k, name: c.name })));
  }

  if (pathname === '/api/actualizar-data' && req.method === 'POST') {
    if (key !== 'reunion') return sendJson(res, 400, { error: 'Actualizar Data solo está disponible para Reunion.' });
    if (actualizarDataRunning) return sendJson(res, 409, { error: 'Actualizar Data ya se está ejecutando.' });
    actualizarDataRunning = true;
    try {
      console.log(' Iniciando Actualizar Data local: POSCO -> PegarData...');
      const currentWorkbook = await loadWorkbook('reunion');
      const result = await runActualizarData(currentWorkbook);
      await saveWholeWorkbook('reunion', result.workbook);
      delete gzCache.reunion;
      console.log(` Actualizar Data completado: ${result.rows} filas en PegarData.`);
      return sendJson(res, 200, {
        ok: true,
        rows: result.rows,
        lastRow: result.lastRow,
        range: `A1:W${result.lastRow}`,
        dateRange: {
          start: result.dateRange.start.toISOString().slice(0, 10),
          end: result.dateRange.end.toISOString().slice(0, 10)
        }
      });
    } catch (error) {
      console.error(' Actualizar Data falló:', error);
      return sendJson(res, 500, { error: error.message || 'No se pudo actualizar PegarData.' });
    } finally {
      actualizarDataRunning = false;
    }
  }

  if (pathname === '/api/workbook' && req.method === 'GET') {
    const workbook = await loadWorkbook(key);
    if (String(req.headers['accept-encoding'] || '').includes('gzip')) {
      if (!gzCache[key]) gzCache[key] = zlib.gzipSync(JSON.stringify(workbook), { level: 6 });
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Encoding': 'gzip', 'Cache-Control': 'no-store' });
      return res.end(gzCache[key]);
    }
    return sendJson(res, 200, workbook);
  }

  if (pathname === '/api/workbook' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req, 300 * 1024 * 1024));
      if (!parsed || typeof parsed !== 'object' || !parsed.sheets) throw new Error('bad workbook');
      await saveWholeWorkbook(key, parsed);
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
    return sendJson(res, 200, { ok: true });
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
  try {
    await initDatabase();
    await loadWorkbook(DEFAULT_WB);
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
  } catch (error) {
    console.error('No se pudo inicializar la persistencia:', error.message);
    process.exit(1);
  }
}

start();
