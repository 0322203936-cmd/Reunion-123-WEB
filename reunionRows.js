// ============================================================================
// REUNION: new flowers of Plan de cosecha get their own row (server.js)
//   * A "Flor 123" of the last P Cosecha tab (or of the tab of the week in progress) that has cuts,
//     is not a flower of sheet Reunion and whose "Flor" (column A of the plan, e.g. LIRIO) is a title
//     of Reunion, is inserted at the end of that title's block.
//   * Rows below move down one row and every formula / filter / conditional format is adjusted
//     (same rules as inserting a row by hand, app.js applyRowOp). Nothing is deleted.
//   * The new row copies the formulas (and formats, no typed numbers) of a sibling flower, and the
//     title's plain sums (=+E155+E156...) get "+E<new row>". Its other formulas are not touched.
// ============================================================================
const CosechaCor = require('./cosechaCor');

const SHEET = 'Reunion';
const norm = CosechaCor.norm;
const isNum = v => typeof v === 'number' && Number.isFinite(v);
const num = v => (isNum(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) ? Number(v) : 0));
const indexToCol = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; };
const colToIndex = col => col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

// ---------------------------------------------------------------- row references (as in app.js)
const RANGE_REF_RE = /(?<![A-Za-z0-9_.$!])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?)(\d+)(?::(\$?[A-Za-z]{1,3}\$?)(\d+))?(?![\w(])/g;
const SHIFT_REF_RE = /(?<![A-Za-z0-9_.$])((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/g;
const unquoteSheet = prefix => { const n = prefix.slice(0, -1); return n.startsWith("'") ? n.slice(1, -1).replace(/''/g, "'") : n; };
// insert op = { sheet, at, n }: rows >= at move down n
function adjustRowRefs(text, homeSheet, op) {
  if (!text) return text;
  return String(text).split(/("[^"]*")/).map(part => {
    if (part.startsWith('"')) return part;
    return part.replace(RANGE_REF_RE, (m, prefix, c1, r1, c2, r2) => {
      const target = prefix ? unquoteSheet(prefix) : homeSheet;
      if (target !== op.sheet) return m;
      const map = r => (r >= op.at ? r + op.n : r);
      if (c2 === undefined) return `${prefix || ''}${c1}${map(Number(r1))}`;
      return `${prefix || ''}${c1}${map(Number(r1))}:${c2}${map(Number(r2))}`;
    });
  }).join('');
}
function shiftFormula(formula, dRow) {
  if (!formula || !dRow) return formula;
  return formula.split(/("[^"]*")/).map(part => {
    if (part.startsWith('"')) return part;
    return part.replace(SHIFT_REF_RE, (m, sheetPrefix, colAbs, col, rowAbs, row) =>
      `${sheetPrefix || ''}${colAbs}${col.toUpperCase()}${rowAbs}${rowAbs ? Number(row) : Number(row) + dRow}`);
  }).join('');
}

// insert `n` empty rows at `at` in sheet SHEET of the book (formulas of every sheet are adjusted)
function insertRows(book, at, n = 1) {
  const op = { sheet: SHEET, at, n };
  const touched = new Set([SHEET]);
  for (const [name, sh] of Object.entries(book.sheets || {})) {
    for (const cell of Object.values(sh.cells || {})) {
      if (!cell || !cell.f) continue;
      const f = adjustRowRefs(cell.f, name, op);
      if (f !== cell.f) { cell.f = f; touched.add(name); }
    }
  }
  const sheet = book.sheets[SHEET];
  const mapRow = r => (r >= at ? r + n : r);
  const adjustList = list => (list || []).map(ref => adjustRowRefs(ref, SHEET, op));
  if (sheet.merges) sheet.merges = adjustList(sheet.merges);
  if (sheet.filters) sheet.filters.forEach(f => { if (f.ref) f.ref = adjustRowRefs(f.ref, SHEET, op); });
  if (sheet.condFormats) sheet.condFormats.forEach(cf => { cf.ranges = adjustList(cf.ranges); });
  for (const key of ['hiddenRows', 'filterHiddenRows']) if (Array.isArray(sheet[key])) sheet[key] = sheet[key].map(mapRow);
  if (sheet.rowHeights) {
    const h = {};
    for (const [r, v] of Object.entries(sheet.rowHeights)) h[mapRow(Number(r))] = v;
    sheet.rowHeights = h;
  }
  if (sheet.styleBounds && sheet.styleBounds.rows >= at) sheet.styleBounds.rows += n;
  for (const t of Object.values(book.tables || {})) if (t.sheet === SHEET && t.ref) t.ref = adjustRowRefs(t.ref, SHEET, op);
  for (const pv of book.pivots || []) {
    if (pv.source && pv.source.sheet === SHEET && pv.source.ref) pv.source.ref = adjustRowRefs(pv.source.ref, SHEET, op);
    if (pv.sheet === SHEET) for (const k of ['location', 'area', 'anchor']) if (pv[k]) pv[k] = adjustRowRefs(pv[k], SHEET, op);
  }
  // move the cells (bottom first)
  const moved = {};
  for (const [coord, cell] of Object.entries(sheet.cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    if (!m) { moved[coord] = cell; continue; }
    moved[`${m[1]}${mapRow(Number(m[2]))}`] = cell;
  }
  sheet.cells = moved;
  sheet.maxRow = (sheet.maxRow || at) + n;
  return touched;
}

// ---------------------------------------------------------------- what is missing
// flowers of the plan with cuts: Map normalized Flor 123 -> { name, group }
function planFlowers(cosechaBook) {
  const out = new Map();
  const main = CosechaCor.lastCosechaSheet(cosechaBook);
  if (!main) return out;
  const cw = CosechaCor.currentWeekSource(cosechaBook, main);
  for (const tab of [main, cw && cw.found]) {
    if (!tab) continue;
    const cells = tab.sheet.cells || {};
    let florCol = null;
    const valueCols = [];
    const groups = (tab.sheet.weekDays && tab.sheet.weekDays.groups) || {};
    for (const [coord, cell] of Object.entries(cells)) {
      const m = /^([A-Z]+)4$/.exec(coord);
      if (!m) continue;
      if (norm(cell.v) === 'FLOR' && !florCol) florCol = m[1];
      if (!cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) {
        valueCols.push(m[1]);
        if (groups[m[1]]) valueCols.push(...groups[m[1]].cols);
      }
    }
    if (!florCol) continue;
    let last = 5;
    for (const coord of Object.keys(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(coord);
      if (m && m[1] === tab.florCol && Number(m[2]) > last) last = Number(m[2]);
    }
    for (let r = 5; r <= last; r++) {
      const raw = (cells[`${tab.florCol}${r}`] || {}).v;
      const key = norm(raw);
      if (!key || out.has(key)) continue;
      const group = norm((cells[`${florCol}${r}`] || {}).v);
      if (!group) continue;
      if (!valueCols.some(col => num((cells[`${col}${r}`] || {}).v) > 0)) continue; // no cuts: no row
      out.set(key, { name: String(raw).replace(/\s+/g, ' ').trim(), group });
    }
  }
  return out;
}

// Reunion layout: its Cor columns, flowers and titles ({ row, children: [rows] } by name)
function reunionLayout(sheet) {
  const cells = sheet.cells || {};
  const corCols = [];
  for (let c = 1; c <= (sheet.maxCol || 80); c++) {
    if (norm((cells[`${indexToCol(c)}2`] || {}).v) === 'COR') corCols.push(indexToCol(c));
  }
  let last = Math.max(sheet.maxRow || 0, 3000);
  while (last > 3 && !cells[`A${last}`]) last--;
  const names = new Set();
  const titles = new Map();
  for (let r = 3; r <= last; r++) {
    const name = norm((cells[`A${r}`] || {}).v);
    if (!name) continue;
    names.add(name);
    const cor = corCols.length && cells[`${corCols[0]}${r}`];
    if (!cor || !cor.f || !isPlainSum(cor.f, corCols[0])) continue;
    const children = refsOf(cor.f).map(x => x.row).filter(x => x > r);
    if (children.length && !titles.has(name)) titles.set(name, { row: r, children });
  }
  return { corCols, names, titles, last };
}
function refsOf(f) {
  const out = [];
  String(f).replace(/\$?([A-Z]{1,3})\$?(\d+)/gi, (m, col, row) => { out.push({ col: col.toUpperCase(), row: Number(row) }); return m; });
  return out;
}
// =+E4+E5 / =sum(E16+E17) / =SUM(E4:E9) of one column only
function isPlainSum(f, col) {
  const refs = refsOf(f);
  if (!refs.length || refs.some(x => x.col !== col)) return false;
  return String(f).replace(/\$?[A-Z]{1,3}\$?\d+/gi, '').replace(/sum/gi, '').replace(/[=+\s():]/g, '') === '';
}

// [{ name, group, title }] flowers to add
function missingFlowers(reunionBook, cosechaBook) {
  const sheet = reunionBook && reunionBook.sheets && reunionBook.sheets[SHEET];
  if (!sheet || !cosechaBook) return [];
  const layout = reunionLayout(sheet);
  const out = [];
  for (const [key, info] of planFlowers(cosechaBook)) {
    if (layout.names.has(key)) continue;
    const title = layout.titles.get(info.group);
    if (title) out.push({ ...info, title: title.row });
  }
  return out;
}

// ---------------------------------------------------------------- add one flower
// returns { row, touched: Set(sheet names) } or null
function addFlower(reunionBook, flower) {
  const sheet = reunionBook.sheets[SHEET];
  const layout = reunionLayout(sheet);
  const title = layout.titles.get(flower.group);
  if (!title || layout.names.has(norm(flower.name))) return null;
  const cells = sheet.cells;
  const at = Math.max(...title.children) + 1;
  // the sibling to copy: a flower of the block whose Req looks for its color (PegarData column U), else the last one
  const usesColor = r => Object.entries(cells).some(([coord, c]) => c.f && new RegExp(`^[A-Z]+${r}$`).test(coord) && /PegarData!\$?U/i.test(c.f));
  const plain = title.children.filter(r => !layout.titles.has(norm((cells[`A${r}`] || {}).v)));
  const pool = plain.length ? plain : title.children;
  const model = [...pool].reverse().find(usesColor) || pool[pool.length - 1];
  const touched = insertRows(reunionBook, at, 1);
  const modelRow = model >= at ? model + 1 : model; // (model is above `at`)
  const dRow = at - modelRow;
  // new row: formulas and formats of the sibling, its name in column A
  for (const [coord, cell] of Object.entries({ ...sheet.cells })) {
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    if (!m || Number(m[2]) !== modelRow) continue;
    const copy = {};
    if (cell.s) copy.s = JSON.parse(JSON.stringify(cell.s));
    if (cell.f) copy.f = shiftFormula(cell.f, dRow);
    if (m[1] === 'A') copy.v = flower.name;
    if (Object.keys(copy).length) sheet.cells[`${m[1]}${at}`] = copy;
  }
  if (!sheet.cells[`A${at}`]) sheet.cells[`A${at}`] = { v: flower.name };
  // conditional formats of the sibling's row also cover the new row
  for (const cf of sheet.condFormats || []) {
    const add = [];
    for (const ref of cf.ranges || []) {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
      if (m && Number(m[2]) === modelRow && (!m[4] || Number(m[4]) === modelRow)) add.push(m[3] ? `${m[1]}${at}:${m[3]}${at}` : `${m[1]}${at}`);
    }
    if (add.length) cf.ranges = [...cf.ranges, ...add];
  }
  // the title's plain sums of its flowers get the new row
  const titleRow = title.row; // above `at`: did not move
  for (const [coord, cell] of Object.entries(sheet.cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    if (!m || Number(m[2]) !== titleRow || !cell.f || !isPlainSum(cell.f, m[1])) continue;
    const refRows = refsOf(cell.f).map(x => x.row);
    if (!refRows.some(r => r > titleRow && r < at)) continue; // not a sum of this block
    const range = /^(=\s*SUM\(\s*\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)(\s*\)\s*)$/i.exec(cell.f);
    if (range && Number(range[2]) === at - 1) cell.f = `${range[1]}${at}${range[3]}`;
    else if (/\)\s*$/.test(cell.f)) cell.f = cell.f.replace(/\)\s*$/, `+${m[1]}${at})`);
    else cell.f = `${cell.f}+${m[1]}${at}`;
    delete cell.v; // recalculated by the pages
  }
  return { row: at, touched };
}

// ---------------------------------------------------------------- titles
// flowers of Reunion named "<name> ..." with no title "<name>" above them: { children: [rows], model: title row }
function missingTitle(reunionBook, name) {
  const sheet = reunionBook && reunionBook.sheets && reunionBook.sheets[SHEET];
  if (!sheet) return null;
  const layout = reunionLayout(sheet);
  const key = norm(name);
  if (layout.names.has(key)) return null;
  const children = [];
  for (let r = 3; r <= layout.last; r++) if (norm((sheet.cells[`A${r}`] || {}).v).startsWith(`${key} `)) children.push(r);
  if (!children.length || children[children.length - 1] - children[0] !== children.length - 1) return null; // not one block
  // model: a title that adds plain flowers and looks its requirement up by FLOR (like LIRIO)
  const usesFlor = r => /PegarData!\$?E/i.test(((sheet.cells[`H${r}`] || {}).f) || '');
  const model = (layout.titles.get('LIRIO') && layout.titles.get('LIRIO').row)
    || [...layout.titles.values()].map(t => t.row).find(usesFlor);
  if (!model) return null;
  return { name: key, children, model };
}
// insert the title row above its flowers: the title's formulas (plain sums -> its flowers, the rest as the model)
function addTitle(reunionBook, info) {
  const sheet = reunionBook.sheets[SHEET];
  const at = Math.min(...info.children);
  const touched = insertRows(reunionBook, at, 1);
  const kids = info.children.map(r => r + 1);
  const modelRow = info.model >= at ? info.model + 1 : info.model;
  for (const [coord, cell] of Object.entries({ ...sheet.cells })) {
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    if (!m || Number(m[2]) !== modelRow) continue;
    const copy = {};
    if (cell.s) copy.s = JSON.parse(JSON.stringify(cell.s));
    if (cell.f) copy.f = isPlainSum(cell.f, m[1]) ? `=+${kids.map(r => `${m[1]}${r}`).join('+')}` : shiftFormula(cell.f, at - modelRow);
    if (m[1] === 'A') copy.v = info.name;
    if (Object.keys(copy).length) sheet.cells[`${m[1]}${at}`] = copy;
  }
  if (!sheet.cells[`A${at}`]) sheet.cells[`A${at}`] = { v: info.name };
  for (const cf of sheet.condFormats || []) {
    const add = [];
    for (const ref of cf.ranges || []) {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
      if (m && Number(m[2]) === modelRow && (!m[4] || Number(m[4]) === modelRow)) add.push(m[3] ? `${m[1]}${at}:${m[3]}${at}` : `${m[1]}${at}`);
    }
    if (add.length) cf.ranges = [...cf.ranges, ...add];
  }
  return { row: at, touched };
}

module.exports = { missingTitle, addTitle, missingFlowers, addFlower, planFlowers, reunionLayout, insertRows, adjustRowRefs, shiftFormula };
