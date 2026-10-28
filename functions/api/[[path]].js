import { neon } from '@neondatabase/serverless';

// Cloudflare Pages API: Neon is the source of truth. Render is no longer used
// for workbook reads/writes. POSCO automation remains in GitHub Actions.
const BOOKS = {
  reunion: { name: 'Reunion 123', prefix: 'reunion123' },
  cosecha: { name: 'Plan de cosecha', prefix: 'cosecha' },
  rvp: { name: 'Requerimiento Vs Proyeccion', prefix: 'rvp' }
};
const DEFAULT_BOOK = 'reunion';
const json = (body, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
const rowsOf = result => Array.isArray(result) ? result : (result?.rows || []);
const bookFrom = request => { const key = new URL(request.url).searchParams.get('wb'); return BOOKS[key] ? key : DEFAULT_BOOK; };
function db(env) { const value = String(env.DATABASE_URL || '').trim(); if (!value) throw new Error('Falta configurar DATABASE_URL en Cloudflare.'); return neon(value); }

function mergeWorkbook(metaRows, sheetRows, cellRows) {
  const meta = metaRows[0]?.meta || { workbookName: '', sheetOrder: [], sheets: {} };
  const sheets = {};
  for (const row of sheetRows) sheets[row.name] = { ...(row.data || {}), cells: { ...((row.data && row.data.cells) || {}) } };
  for (const row of cellRows) { if (!sheets[row.sheet_name]) sheets[row.sheet_name] = { cells: {} }; sheets[row.sheet_name].cells[row.coord] = row.cell; }
  return { ...meta, sheets };
}
async function loadWorkbook(sql, book) {
  const p = BOOKS[book].prefix;
  const [meta, sheets, cells] = await Promise.all([
    sql.query(`SELECT meta FROM ${p}_workbook_meta WHERE id = 1`),
    sql.query(`SELECT name, data FROM ${p}_workbook_sheets`),
    // Aggregate cells in PostgreSQL instead of sending one result row per
    // cell to the Worker. The workbook stays Neon-backed, but the Worker does
    // far less JSON merging and avoids intermittent 503s on large books.
    sql.query(`SELECT sheet_name, jsonb_object_agg(coord, cell) AS cells FROM ${p}_workbook_cells GROUP BY sheet_name`)
  ]);
  const workbook = mergeWorkbook(rowsOf(meta), rowsOf(sheets), []);
  for (const row of rowsOf(cells)) {
    if (!workbook.sheets[row.sheet_name]) workbook.sheets[row.sheet_name] = { cells: {} };
    workbook.sheets[row.sheet_name].cells = row.cells || {};
  }
  return workbook;
}
async function loadWorkbookIndex(sql, book) {
  const p = BOOKS[book].prefix;
  const [meta, sheets] = await Promise.all([
    sql.query(`SELECT meta FROM ${p}_workbook_meta WHERE id = 1`),
    sql.query(`SELECT name, data FROM ${p}_workbook_sheets`)
  ]);
  return mergeWorkbook(rowsOf(meta), rowsOf(sheets), []);
}
async function loadSheet(sql, book, name) {
  const p = BOOKS[book].prefix;
  const [sheetResult, cellsResult] = await Promise.all([
    sql.query(`SELECT name, data FROM ${p}_workbook_sheets WHERE name = $1`, [name]),
    sql.query(`SELECT jsonb_object_agg(coord, cell) AS cells FROM ${p}_workbook_cells WHERE sheet_name = $1`, [name])
  ]);
  const sheetRow = rowsOf(sheetResult)[0];
  if (!sheetRow) return null;
  return { ...(sheetRow.data || {}), cells: rowsOf(cellsResult)[0]?.cells || {} };
}
function splitWorkbook(workbook) { const { sheets = {}, ...meta } = workbook || {}; return { meta, sheets }; }
function withoutCells(sheet) { if (!sheet || typeof sheet !== 'object') return sheet; const { cells, ...rest } = sheet; return rest; }
function cellRows(sheetName, sheet) { return Object.entries((sheet && sheet.cells) || {}).map(([coord, cell]) => ({ sheet_name: sheetName, coord, cell })); }
async function upsertCells(sql, table, rows) {
  // Keep the total Neon subrequests below Cloudflare's Worker invocation
  // limit when importing a large workbook.
  for (let i = 0; i < rows.length; i += 10000) {
    const batch = rows.slice(i, i + 10000);
    await sql.query(`INSERT INTO ${table} (sheet_name, coord, cell)
      SELECT sheet_name, coord, cell
      FROM jsonb_to_recordset($1::jsonb) AS x(sheet_name text, coord text, cell jsonb)
      ON CONFLICT (sheet_name, coord) DO UPDATE SET cell = EXCLUDED.cell, updated_at = NOW()`, [JSON.stringify(batch)]);
  }
}
async function savePatch(sql, book, patch) {
  const p = BOOKS[book].prefix;
  const changed = patch.sheets && typeof patch.sheets === 'object' ? patch.sheets : {};
  const cellPatches = patch.cellPatches && typeof patch.cellPatches === 'object' ? patch.cellPatches : {};
  const deleted = Array.isArray(patch.deleted) ? patch.deleted : [];
  const meta = patch.meta && typeof patch.meta === 'object' ? patch.meta : {};
  for (const [name, data] of Object.entries(changed)) {
    // A POSCO failure or a lazy-loaded workbook must never replace PegarData
    // with an empty sheet. Only accept a full replacement when it contains
    // actual cells; normal cell patches remain supported below.
    if (book === 'reunion' && name === 'PegarData' && (!data || typeof data !== 'object' || !data.cells || !Object.keys(data.cells).length)) {
      const existing = rowsOf(await sql.query(`SELECT 1 FROM reunion123_workbook_cells WHERE sheet_name = $1 LIMIT 1`, [name]));
      if (existing.length) throw new Error('Se rechazó reemplazar PegarData vacía para proteger los datos existentes.');
    }
    await sql.query(`INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb) ON CONFLICT (name) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()`, [name, JSON.stringify(withoutCells(data))]);
    await sql.query(`DELETE FROM ${p}_workbook_cells WHERE sheet_name = $1`, [name]);
    await upsertCells(sql, `${p}_workbook_cells`, cellRows(name, data));
  }
  for (const [name, change] of Object.entries(cellPatches)) {
    for (const [coord, cell] of Object.entries((change && change.set) || {})) await sql.query(`INSERT INTO ${p}_workbook_cells (sheet_name, coord, cell) VALUES ($1, $2, $3::jsonb) ON CONFLICT (sheet_name, coord) DO UPDATE SET cell = EXCLUDED.cell, updated_at = NOW()`, [name, coord, JSON.stringify(cell)]);
    for (const coord of (Array.isArray(change?.delete) ? change.delete : [])) await sql.query(`DELETE FROM ${p}_workbook_cells WHERE sheet_name = $1 AND coord = $2`, [name, coord]);
  }
  for (const name of deleted) await sql.query(`DELETE FROM ${p}_workbook_sheets WHERE name = $1`, [name]);
  if (Object.keys(meta).length) {
    const current = await loadWorkbook(sql, book);
    await sql.query(`INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb) ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`, [JSON.stringify({ ...splitWorkbook(current).meta, ...meta })]);
  }
}
async function saveWhole(sql, book, workbook) {
  const p = BOOKS[book].prefix; const { meta, sheets } = splitWorkbook(workbook);
  await sql.query(`DELETE FROM ${p}_workbook_sheets`); await sql.query(`DELETE FROM ${p}_workbook_cells`);
  for (const [name, sheet] of Object.entries(sheets)) { await sql.query(`INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)`, [name, JSON.stringify(withoutCells(sheet))]); await upsertCells(sql, `${p}_workbook_cells`, cellRows(name, sheet)); }
  await sql.query(`INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb) ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`, [JSON.stringify(meta)]);
}
async function ensureRuns(sql) {
  // Cloudflare can receive two requests at the same time (for example, a
  // status poll while a button starts a run). PostgreSQL creates an internal
  // row type for a table; concurrent CREATE TABLE IF NOT EXISTS calls can
  // therefore collide on pg_type_typname_nsp_index. Keep initialization
  // idempotent and swallow only that expected race.
  try {
    await sql.query(`CREATE TABLE IF NOT EXISTS reunion123_posco_runs (
      operation TEXT PRIMARY KEY,
      status TEXT NOT NULL,
      started_at TIMESTAMPTZ,
      finished_at TIMESTAMPTZ,
      rows INTEGER,
      last_row INTEGER,
      range TEXT,
      date_range JSONB,
      error TEXT
    )`);
  } catch (error) {
    const message = String(error?.message || error);
    const isConcurrentCreate = /pg_type_typname_nsp_index|duplicate key value|already exists|duplicate table/i.test(message);
    if (!isConcurrentCreate) throw error;
  }
}
async function readRun(sql, operation) {
  await ensureRuns(sql);
  const result = rowsOf(await sql.query(`SELECT operation, status, started_at AS "startedAt", finished_at AS "finishedAt", rows, last_row AS "lastRow", range, date_range AS "dateRange", error FROM reunion123_posco_runs WHERE operation = $1`, [operation]));
  const row = result[0];
  return row ? { ...row, ok: row.status === 'succeeded' } : { operation, status: 'idle', startedAt: null, finishedAt: null, rows: null, lastRow: null, range: null, dateRange: null, error: null, ok: false };
}
async function dispatchPosco(env, operation) {
  const token = String(env.GITHUB_ACTIONS_TOKEN || '').trim(); if (!token) return json({ error: 'Falta configurar GITHUB_ACTIONS_TOKEN en Cloudflare.' }, 500);
  const repository = String(env.GITHUB_REPOSITORY || '0322203936-cmd/Reunion-123-WEB').trim();
  const response = await fetch(`https://api.github.com/repos/${repository}/dispatches`, { method: 'POST', headers: { accept: 'application/vnd.github+json', authorization: `Bearer ${token}`, 'x-github-api-version': '2022-11-28', 'user-agent': 'reunion-123-cloudflare' }, body: JSON.stringify({ event_type: 'posco-update', client_payload: { operation } }) });
  if (!response.ok) return json({ error: `GitHub rechazó el workflow (HTTP ${response.status}).` }, 502);
  return json({ ok: true, status: 'running', operation }, 202);
}

export async function onRequest(context) {
  const { request, env } = context; const url = new URL(request.url); const pathname = decodeURIComponent(url.pathname); const book = bookFrom(request);
  if (!pathname.startsWith('/api/')) return context.next();
  try {
    const sql = db(env);
    if (pathname === '/api/workbooks' && request.method === 'GET') return json(Object.entries(BOOKS).map(([key, value]) => ({ key, name: value.name })));
    if (pathname === '/api/workbook' && request.method === 'GET') {
      return json(url.searchParams.get('index') === '1' ? await loadWorkbookIndex(sql, book) : await loadWorkbook(sql, book));
    }
    if (pathname === '/api/workbook' && request.method === 'PATCH') { await savePatch(sql, book, await request.json()); return json({ ok: true }); }
    if (pathname === '/api/workbook' && request.method === 'POST') { const workbook = await request.json(); if (!workbook || typeof workbook !== 'object' || !workbook.sheets) return json({ error: 'JSON de libro inválido' }, 400); await saveWhole(sql, book, workbook); return json({ ok: true }); }
    if (pathname === '/api/sheet' && request.method === 'GET') { const sheet = await loadSheet(sql, book, url.searchParams.get('name') || ''); return sheet ? json(sheet) : json({ error: 'Hoja no encontrada' }, 404); }
    if (pathname === '/api/homologations' && request.method === 'GET') { const p = BOOKS[book].prefix; const result = rowsOf(await sql.query(`SELECT state FROM ${p}_homologations WHERE id = 1`)); return json(result[0]?.state || { version: 1, items: {} }); }
    if (pathname === '/api/homologations' && request.method === 'POST') { const p = BOOKS[book].prefix; await sql.query(`UPDATE ${p}_homologations SET state = $1::jsonb, updated_at = NOW() WHERE id = 1`, [JSON.stringify(await request.json())]); return json({ ok: true }); }
    if (pathname === '/api/actualizar-data' && (request.method === 'GET' || request.method === 'POST')) {
      if (request.method === 'GET') return json(await readRun(sql, 'data'));
      await ensureRuns(sql);
      await sql.query(`INSERT INTO reunion123_posco_runs (operation, status, started_at, finished_at, error) VALUES ('data', 'running', NOW(), NULL, NULL) ON CONFLICT (operation) DO UPDATE SET status = 'running', started_at = NOW(), finished_at = NULL, rows = NULL, last_row = NULL, range = NULL, date_range = NULL, error = NULL`);
      const response = await dispatchPosco(env, 'data');
      if (!response.ok) await sql.query(`UPDATE reunion123_posco_runs SET status = 'failed', finished_at = NOW(), error = $1 WHERE operation = 'data'`, [await response.clone().text()]);
      return response;
    }
    if (pathname === '/api/actualizar-inventario' && (request.method === 'GET' || request.method === 'POST')) {
      if (request.method === 'GET') return json(await readRun(sql, 'inventario'));
      await ensureRuns(sql);
      await sql.query(`INSERT INTO reunion123_posco_runs (operation, status, started_at, finished_at, error) VALUES ('inventario', 'running', NOW(), NULL, NULL) ON CONFLICT (operation) DO UPDATE SET status = 'running', started_at = NOW(), finished_at = NULL, rows = NULL, last_row = NULL, range = NULL, date_range = NULL, error = NULL`);
      const response = await dispatchPosco(env, 'inventario');
      if (!response.ok) await sql.query(`UPDATE reunion123_posco_runs SET status = 'failed', finished_at = NOW(), error = $1 WHERE operation = 'inventario'`, [await response.clone().text()]);
      return response;
    }
    if (pathname === '/api/posco-callback' && request.method === 'POST') {
      const expected = String(env.REUNION_SYNC_TOKEN || '').trim();
      if (!expected || request.headers.get('authorization') !== `Bearer ${expected}`) return json({ error: 'Callback no autorizado.' }, 401);
      const payload = await request.json();
      if (!['data', 'inventario'].includes(String(payload.operation || ''))) return json({ error: 'Operación inválida.' }, 400);
      await ensureRuns(sql);
      await sql.query(`UPDATE reunion123_posco_runs SET status = $1, finished_at = NOW(), rows = $2, last_row = $3, range = $4, date_range = $5::jsonb, error = $6 WHERE operation = $7`, [payload.status === 'succeeded' ? 'succeeded' : 'failed', Number.isFinite(Number(payload.rows)) ? Number(payload.rows) : null, Number.isFinite(Number(payload.lastRow)) ? Number(payload.lastRow) : null, payload.range || null, JSON.stringify(payload.dateRange || null), payload.error || null, payload.operation]);
      return json({ ok: true });
    }
    if (pathname === '/api/events' && request.method === 'GET') { const stream = new ReadableStream({ start(controller) { const encoder = new TextEncoder(); controller.enqueue(encoder.encode(': connected\n\n')); const timer = setInterval(() => controller.enqueue(encoder.encode(': ping\n\n')), 25000); request.signal.addEventListener('abort', () => { clearInterval(timer); try { controller.close(); } catch {} }); } }); return new Response(stream, { headers: { 'content-type': 'text/event-stream; charset=utf-8', 'cache-control': 'no-store', connection: 'keep-alive' } }); }
    if (pathname === '/api/whoami' && request.method === 'GET') return json({ user: '' });
    return json({ error: 'API no encontrada' }, 404);
  } catch (error) { console.error('Cloudflare API error:', error); return json({ error: error.message || 'Error interno del servidor' }, 500); }
}
