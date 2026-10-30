// Puts back the formulas of the Cor columns of sheet "Reunion" (title rows such as =+E4+E5)
// when they were overwritten, e.g. by pasting numbers over them. The formulas come from the
// copy in GitHub (data/workbook_state.json); each title is found by its name in column A, so it
// still works if rows were moved. Numbers typed in the product rows are left as they are.
//
// Usage (with the system stopped):  node tools/restaurar_formulas_cor.js [reference.json]
const fs = require('fs');
const path = require('path');
const http = require('http');
const { Pool } = require('pg');
const NeonStore = require('../neonStore');

const ROOT = path.join(__dirname, '..');
const PREFIX = 'reunion123';
const SHEET = 'Reunion';
const REF_ARRAY = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;

function loadLocalEnv() {
  const envPath = path.join(ROOT, '.env');
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

const norm = v => String(v === undefined || v === null ? '' : v).trim().toUpperCase();
const label = (sheet, r) => norm((sheet.cells[`A${r}`] || {}).v);
function indexToCol(i) {
  let s = '';
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}
function corColumns(sheet) {
  const out = [];
  for (let c = 1; c <= (sheet.maxCol || 80); c++) {
    if (norm((sheet.cells[`${indexToCol(c)}2`] || {}).v) === 'COR') out.push(indexToCol(c));
  }
  return out;
}
// rows of each name in column A, in order (a name may appear more than once)
function rowsByLabel(sheet) {
  const out = new Map();
  for (const coord of Object.keys(sheet.cells)) {
    const m = /^A(\d+)$/.exec(coord);
    if (!m) continue;
    const name = label(sheet, Number(m[1]));
    if (!name) continue;
    if (!out.has(name)) out.set(name, []);
    out.get(name).push(Number(m[1]));
  }
  for (const rows of out.values()) rows.sort((a, b) => a - b);
  return out;
}

// changes `cur` in place; returns { restored: [coords], skipped: [reasons] }
function repair(ref, cur) {
  const restored = [];
  const skipped = [];
  const cols = corColumns(ref);
  const curCols = new Set(corColumns(cur));
  const refRows = rowsByLabel(ref);
  const curRows = rowsByLabel(cur);
  const byRow = new Map(); // ref row -> [{col, f}]
  for (const [coord, cell] of Object.entries(ref.cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    if (!m || !cell.f || !cols.includes(m[1]) || !curCols.has(m[1])) continue;
    const r = Number(m[2]);
    if (!byRow.has(r)) byRow.set(r, []);
    byRow.get(r).push({ col: m[1], f: cell.f });
  }
  // where row `r` of the copy is now: same name in column A (same occurrence); a row without a
  // name (the total at the bottom) keeps its distance to the nearest named row above it
  function mapRow(r) {
    const name = label(ref, r);
    if (name) {
      const nth = (refRows.get(name) || []).indexOf(r);
      return (curRows.get(name) || [])[nth];
    }
    let a = r - 1;
    while (a >= 1 && !label(ref, a)) a--;
    const anchor = a >= 1 ? mapRow(a) : 0;
    return anchor === undefined ? undefined : anchor + (r - Math.max(a, 0));
  }
  for (const [refRow, formulas] of byRow) {
    const name = label(ref, refRow) || `fila ${refRow}`;
    const curRow = mapRow(refRow);
    if (!curRow) { skipped.push(`${name}: ya no está en la hoja`); continue; }
    for (const { col, f } of formulas) {
      // every row the formula uses must still be in the sheet
      let ok = true;
      const formula = f.replace(REF_ARRAY, (all, d1, c, d2, row) => {
        const r = mapRow(Number(row));
        if (!r) { ok = false; return all; }
        return `${d1}${c}${d2}${r}`;
      });
      if (!ok) { skipped.push(`${name} ${col}${curRow}: las filas que suma cambiaron`); continue; }
      const coord = `${col}${curRow}`;
      const cell = cur.cells[coord];
      if (cell && cell.f === formula) continue;
      cur.cells[coord] = { ...(cell || {}), f: formula };
      delete cur.cells[coord].v;
      restored.push(coord);
    }
  }
  return { restored, skipped };
}

function serverRunning(port) {
  return new Promise(resolve => {
    const req = http.get({ host: '127.0.0.1', port, path: '/', timeout: 1500 }, res => { res.resume(); resolve(true); });
    req.on('error', () => resolve(false));
    req.on('timeout', () => { req.destroy(); resolve(false); });
  });
}

async function main() {
  loadLocalEnv();
  if (!process.env.DATABASE_URL) throw new Error('Falta DATABASE_URL en el archivo .env');
  const port = Number(process.env.PORT || 3070);
  if (await serverRunning(port)) {
    throw new Error(`El sistema está prendido (puerto ${port}). Apágalo con Ctrl + C y vuelve a correr esto.`);
  }
  const refPath = process.argv[2] || path.join(ROOT, 'data', 'workbook_state.json');
  const reference = JSON.parse(fs.readFileSync(refPath, 'utf8')).sheets[SHEET];
  if (!reference) throw new Error(`La copia ${refPath} no tiene la hoja ${SHEET}`);

  const local = /@(localhost|127\.0\.0\.1)[:/]/.test(process.env.DATABASE_URL);
  const pool = new Pool({ connectionString: process.env.DATABASE_URL, ssl: local ? false : { rejectUnauthorized: false }, max: 2 });
  try {
    const book = await NeonStore.loadBook(pool, PREFIX);
    if (!book || !book.sheets[SHEET]) throw new Error(`No se encontró la hoja ${SHEET} en Neon`);
    const sheet = book.sheets[SHEET];
    const { restored, skipped } = repair(reference, sheet);
    for (const s of skipped) console.log(' No se tocó:', s);
    if (!restored.length) { console.log(' No había fórmulas de Cor que restaurar.'); return; }
    await NeonStore.writeBackup(pool, 'reunion', { ...book.meta, sheets: book.sheets }); // copy before the fix
    await NeonStore.transaction(pool, client => NeonStore.writeSnapshots(client, PREFIX, { [SHEET]: sheet }));
    console.log(` Listo: ${restored.length} fórmulas de Cor restauradas (ejemplo: ${restored.slice(0, 5).join(', ')}).`);
  } finally {
    await pool.end();
  }
}

if (require.main === module) {
  main().catch(err => { console.error(' Error:', err.message); process.exit(1); });
}
module.exports = { repair };
