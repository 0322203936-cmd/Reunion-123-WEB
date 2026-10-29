// ============================================================================
// NEON STORAGE (shared by server.js and github_actions_runner.js)
//   Every book has three tables (prefix = reunion123 / cosecha / rvp):
//     {p}_workbook_meta    one row: the book without its sheets
//     {p}_workbook_sheets  one row per sheet: its last full copy ("snapshot")
//     {p}_sheet_deltas     small changes written after the snapshot (only the cells that changed)
//   Reading a sheet = its snapshot + its deltas in order. When a sheet piles up many deltas it is
//   written whole again and its deltas are deleted ("compaction"), so reads stay fast.
//   A delta: { cells: { A1: cell | null }, props: { key: value }, removed: [key] }   (null = cell deleted)
// ============================================================================

const zlib = require('zlib');

async function ensureTables(pool, p) {
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
    CREATE TABLE IF NOT EXISTS ${p}_sheet_deltas (
      id BIGSERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      delta JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    CREATE INDEX IF NOT EXISTS ${p}_sheet_deltas_name ON ${p}_sheet_deltas (name, id);
    CREATE TABLE IF NOT EXISTS ${p}_homologations (
      id SMALLINT PRIMARY KEY CHECK (id = 1),
      state JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
}

// shared tables: small key/value state (Cor history) and daily backups
async function ensureSharedTables(pool) {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_state (
      key TEXT PRIMARY KEY,
      value JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    CREATE TABLE IF NOT EXISTS app_backups (
      id BIGSERIAL PRIMARY KEY,
      book TEXT NOT NULL,
      data BYTEA NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);
}

function applyDelta(sheet, delta) {
  if (delta.cells) {
    sheet.cells = sheet.cells || {};
    for (const [coord, cell] of Object.entries(delta.cells)) {
      if (cell === null) delete sheet.cells[coord];
      else sheet.cells[coord] = cell;
    }
  }
  if (delta.props) Object.assign(sheet, delta.props);
  for (const key of delta.removed || []) delete sheet[key];
  return sheet;
}

// { meta, sheets, deltaCounts: { name: n } } or null when the book was never stored
async function loadBook(pool, p) {
  const [metaRes, sheetRes, deltaRes] = await Promise.all([
    pool.query(`SELECT meta FROM ${p}_workbook_meta WHERE id = 1`),
    pool.query(`SELECT name, data FROM ${p}_workbook_sheets`),
    pool.query(`SELECT name, delta FROM ${p}_sheet_deltas ORDER BY id`)
  ]);
  if (!metaRes.rows[0]) return null;
  const sheets = {};
  for (const row of sheetRes.rows) sheets[row.name] = row.data;
  const deltaCounts = {};
  for (const row of deltaRes.rows) {
    if (!sheets[row.name]) continue; // sheet deleted after the delta
    applyDelta(sheets[row.name], row.delta);
    deltaCounts[row.name] = (deltaCounts[row.name] || 0) + 1;
  }
  return { meta: metaRes.rows[0].meta, sheets, deltaCounts };
}

// full copy of some sheets (their old deltas go away); `client` is inside a transaction
async function writeSnapshots(client, p, sheets) {
  for (const [name, data] of Object.entries(sheets)) {
    await client.query(
      `INSERT INTO ${p}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)
       ON CONFLICT (name) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()`,
      [name, JSON.stringify(data)]
    );
    await client.query(`DELETE FROM ${p}_sheet_deltas WHERE name = $1`, [name]);
  }
}

async function writeMeta(client, p, meta) {
  await client.query(
    `INSERT INTO ${p}_workbook_meta (id, meta) VALUES (1, $1::jsonb)
     ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`,
    [JSON.stringify(meta)]
  );
}

// the whole book replaces what was stored (first upload)
async function writeWholeBook(pool, p, meta, sheets) {
  await transaction(pool, async client => {
    await client.query(`DELETE FROM ${p}_workbook_sheets`);
    await client.query(`DELETE FROM ${p}_sheet_deltas`);
    await writeSnapshots(client, p, sheets);
    await writeMeta(client, p, meta);
  });
}

async function transaction(pool, fn) {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const out = await fn(client);
    await client.query('COMMIT');
    return out;
  } catch (error) {
    await client.query('ROLLBACK').catch(() => {});
    throw error;
  } finally {
    client.release();
  }
}

// ---------------------------------------------------------------- what changed in a sheet
// `known` (Map coord -> JSON of the cell as last stored) is updated to the new state.
// Returns null when nothing changed, or { delta, size } (size: number of changed entries).
function diffSheet(known, knownProps, sheet) {
  const cells = {};
  let n = 0;
  const seen = new Set();
  for (const [coord, cell] of Object.entries(sheet.cells || {})) {
    seen.add(coord);
    const s = JSON.stringify(cell);
    if (known.get(coord) === s) continue;
    known.set(coord, s);
    cells[coord] = cell;
    n++;
  }
  for (const coord of known.keys()) {
    if (seen.has(coord)) continue;
    known.delete(coord);
    cells[coord] = null;
    n++;
  }
  const props = {};
  const removed = [];
  const keys = new Set([...Object.keys(sheet), ...knownProps.keys()]);
  for (const key of keys) {
    if (key === 'cells') continue;
    if (!(key in sheet)) { knownProps.delete(key); removed.push(key); n++; continue; }
    const s = JSON.stringify(sheet[key]);
    if (knownProps.get(key) === s) continue;
    knownProps.set(key, s);
    props[key] = sheet[key];
    n++;
  }
  if (!n) return null;
  const delta = {};
  if (Object.keys(cells).length) delta.cells = cells;
  if (Object.keys(props).length) delta.props = props;
  if (removed.length) delta.removed = removed;
  return { delta, size: n };
}

// remembers a sheet as stored, for the next diff
function snapshotOf(sheet) {
  const known = new Map();
  for (const [coord, cell] of Object.entries(sheet.cells || {})) known.set(coord, JSON.stringify(cell));
  const knownProps = new Map();
  for (const [key, value] of Object.entries(sheet)) if (key !== 'cells') knownProps.set(key, JSON.stringify(value));
  return { known, knownProps };
}

// merges a newer delta into an older pending one (same sheet)
function mergeDelta(into, delta) {
  if (delta.cells) into.cells = { ...(into.cells || {}), ...delta.cells };
  if (delta.props) {
    into.props = { ...(into.props || {}), ...delta.props };
    if (into.removed) into.removed = into.removed.filter(k => !(k in delta.props));
  }
  if (delta.removed) {
    into.removed = [...new Set([...(into.removed || []), ...delta.removed])];
    if (into.props) for (const k of delta.removed) delete into.props[k];
  }
  return into;
}

// ---------------------------------------------------------------- backups (gzip, last N per book)
async function writeBackup(pool, book, workbook, keep = 7) {
  const data = zlib.gzipSync(JSON.stringify(workbook), { level: 9 });
  await pool.query('INSERT INTO app_backups (book, data) VALUES ($1, $2)', [book, data]);
  await pool.query(
    `DELETE FROM app_backups WHERE book = $1 AND id NOT IN (
       SELECT id FROM app_backups WHERE book = $1 ORDER BY id DESC LIMIT $2)`,
    [book, keep]
  );
  return data.length;
}

module.exports = {
  ensureTables, ensureSharedTables, loadBook, applyDelta, writeSnapshots, writeMeta, writeWholeBook,
  transaction, diffSheet, snapshotOf, mergeDelta, writeBackup
};
