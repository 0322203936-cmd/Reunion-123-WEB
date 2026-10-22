// ============================================================================
// EXCEL FORMULA ENGINE - REUNION 123
// Tokenizer + parser (AST cached per formula) + memoized evaluation with
// Excel semantics, cross-sheet refs, whole-column ranges, structured table
// refs and indexed SUMIF(S)/COUNTIF(S).
// ============================================================================

function colToIndex(colStr) {
  let idx = 0;
  const s = colStr.toUpperCase();
  for (let i = 0; i < s.length; i++) idx = idx * 26 + (s.charCodeAt(i) - 64);
  return idx;
}

function indexToCol(idx) {
  let s = '';
  let n = idx;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseCoord(coordStr) {
  const clean = String(coordStr).replace(/\$/g, '').trim();
  const m = clean.match(/^([A-Za-z]{1,3})(\d+)$/);
  if (!m) return null;
  return {
    colStr: m[1].toUpperCase(),
    col: colToIndex(m[1]),
    row: parseInt(m[2], 10),
    coord: `${m[1].toUpperCase()}${parseInt(m[2], 10)}`
  };
}

// ---------------------------------------------------------------- values
class XlError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
const ERR = {
  DIV0: new XlError('#DIV/0!'),
  VALUE: new XlError('#VALUE!'),
  REF: new XlError('#REF!'),
  NA: new XlError('#N/A'),
  NAME: new XlError('#NAME?'),
  NUM: new XlError('#NUM!')
};
const ERROR_CODES = new Set(Object.values(ERR).map(e => e.code).concat(['#NULL!', '#SPILL!', '#CALC!']));
const isErr = v => v instanceof XlError;
const isBlank = v => v === null || v === undefined || v === '';

// Range value passed to functions (lazy)
class RangeRef {
  constructor(sheet, r1, c1, r2, c2) {
    this.sheet = sheet;
    this.r1 = Math.min(r1, r2); this.r2 = Math.max(r1, r2);
    this.c1 = Math.min(c1, c2); this.c2 = Math.max(c1, c2);
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  key() { return `${this.sheet}!${this.r1}:${this.c1}:${this.r2}:${this.c2}`; }
}

// ---------------------------------------------------------------- tokenizer
const TOKEN_RE = {
  ws: /^\s+/,
  str: /^"((?:[^"]|"")*)"/,
  err: /^#(?:DIV\/0!|VALUE!|REF!|N\/A|NAME\?|NUM!|NULL!|SPILL!|CALC!)/i,
  num: /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/,
  // sheet prefix: 'Quoted Name'! or Name!
  sheet: /^(?:'((?:[^']|'')+)'|([A-Za-z_À-ɏ][\w.À-ɏ]*))!/,
  cellRange: /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\w(])/,
  colRange: /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(\d])/,
  rowRange: /^(\$?)(\d+):(\$?)(\d+)(?![\w(])/,
  func: /^([A-Za-z_][\w.]*)\s*\(/,
  name: /^([A-Za-z_À-ɏ][\w.À-ɏ]*)/,
  op: /^(<>|>=|<=|[-+*/^&=<>%])/
};

function readStructuredRef(src, i) {
  // src[i] === '['  -> read balanced brackets
  let depth = 0;
  let j = i;
  for (; j < src.length; j++) {
    if (src[j] === '[') depth++;
    else if (src[j] === ']') { depth--; if (depth === 0) break; }
  }
  return src.slice(i, j + 1);
}

function tokenize(src) {
  const tokens = [];
  let i = 0;
  while (i < src.length) {
    const rest = src.slice(i);
    let m;
    if ((m = TOKEN_RE.ws.exec(rest))) { i += m[0].length; continue; }
    const ch = src[i];
    if (ch === '(' || ch === ')' || ch === ',' || ch === ';') {
      tokens.push({ t: ch === ';' ? ',' : ch });
      i++; continue;
    }
    if ((m = TOKEN_RE.str.exec(rest))) { tokens.push({ t: 'str', v: m[1].replace(/""/g, '"') }); i += m[0].length; continue; }
    if ((m = TOKEN_RE.err.exec(rest))) { tokens.push({ t: 'err', v: m[0].toUpperCase() }); i += m[0].length; continue; }

    // references (optionally sheet-qualified)
    let sheet = null;
    let consumed = 0;
    if ((m = TOKEN_RE.sheet.exec(rest))) {
      sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
      consumed = m[0].length;
    }
    const afterSheet = rest.slice(consumed);
    if ((m = TOKEN_RE.colRange.exec(afterSheet))) {
      tokens.push({ t: 'range', sheet, c1: colToIndex(m[2]), c2: colToIndex(m[4]), ca1: !!m[1], ca2: !!m[3], wholeCol: true });
      i += consumed + m[0].length; continue;
    }
    if ((m = TOKEN_RE.rowRange.exec(afterSheet))) {
      tokens.push({ t: 'range', sheet, r1: Number(m[2]), r2: Number(m[4]), ra1: !!m[1], ra2: !!m[3], wholeRow: true });
      i += consumed + m[0].length; continue;
    }
    if ((m = TOKEN_RE.cellRange.exec(afterSheet))) {
      if (m[6]) {
        tokens.push({
          t: 'range', sheet,
          c1: colToIndex(m[2]), r1: Number(m[4]), c2: colToIndex(m[6]), r2: Number(m[8]),
          ca1: !!m[1], ra1: !!m[3], ca2: !!m[5], ra2: !!m[7]
        });
      } else {
        tokens.push({ t: 'ref', sheet, c: colToIndex(m[2]), r: Number(m[4]), ca: !!m[1], ra: !!m[3] });
      }
      i += consumed + m[0].length; continue;
    }
    if (sheet !== null) throw ERR.REF;

    if ((m = TOKEN_RE.num.exec(rest))) { tokens.push({ t: 'num', v: Number(m[0]) }); i += m[0].length; continue; }
    if ((m = TOKEN_RE.func.exec(rest))) {
      tokens.push({ t: 'func', v: m[1].toUpperCase().replace(/^_XLFN\.|^_XLWS\./, '') });
      i += m[0].length; continue;
    }
    if ((m = TOKEN_RE.name.exec(rest))) {
      const name = m[1];
      i += m[0].length;
      if (src[i] === '[') {
        const spec = readStructuredRef(src, i);
        tokens.push({ t: 'sref', table: name, spec });
        i += spec.length;
        continue;
      }
      const up = name.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') tokens.push({ t: 'bool', v: up === 'TRUE' });
      else tokens.push({ t: 'name', v: name });
      continue;
    }
    if (ch === '[') { // structured ref without table name (inside a table)
      const spec = readStructuredRef(src, i);
      tokens.push({ t: 'sref', table: null, spec });
      i += spec.length; continue;
    }
    if ((m = TOKEN_RE.op.exec(rest))) { tokens.push({ t: 'op', v: m[1] }); i += m[0].length; continue; }
    throw ERR.NAME;
  }
  return tokens;
}

// ---------------------------------------------------------------- parser (precedence climbing)
// comparison < & < +- < */ < ^ < unary < %
const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

function parse(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];

  function parsePrimary() {
    const tk = next();
    if (!tk) throw ERR.VALUE;
    switch (tk.t) {
      case 'num': return { k: 'num', v: tk.v };
      case 'str': return { k: 'str', v: tk.v };
      case 'bool': return { k: 'bool', v: tk.v };
      case 'err': return { k: 'err', v: tk.v };
      case 'ref': return { k: 'ref', sheet: tk.sheet, r: tk.r, c: tk.c, ra: tk.ra, ca: tk.ca };
      case 'range': return { k: 'range', ...tk };
      case 'sref': return { k: 'sref', table: tk.table, spec: tk.spec };
      case 'name': return { k: 'name', v: tk.v };
      case '(': {
        const e = parseExpr(0);
        if (!peek() || peek().t !== ')') throw ERR.VALUE;
        next();
        return e;
      }
      case 'func': {
        const args = [];
        if (peek() && peek().t === ')') { next(); return { k: 'fn', name: tk.v, args }; }
        for (;;) {
          if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ k: 'missing' });
          else args.push(parseExpr(0));
          const sep = next();
          if (!sep) throw ERR.VALUE;
          if (sep.t === ')') break;
          if (sep.t !== ',') throw ERR.VALUE;
        }
        return { k: 'fn', name: tk.v, args };
      }
      case 'op':
        if (tk.v === '-' || tk.v === '+') return { k: 'un', op: tk.v, e: parseUnaryOperand() };
        throw ERR.VALUE;
      default: throw ERR.VALUE;
    }
  }

  function parseUnaryOperand() {
    // unary binds tighter than ^ in Excel (-2^2 = 4)
    let e = parsePrimary();
    while (peek() && peek().t === 'op' && peek().v === '%') { next(); e = { k: 'pct', e }; }
    return e;
  }

  function parseExpr(minPrec) {
    let left = parseUnaryOperand();
    for (;;) {
      const tk = peek();
      if (!tk || tk.t !== 'op' || !(tk.v in BIN_PREC)) break;
      const prec = BIN_PREC[tk.v];
      if (prec < minPrec) break;
      next();
      const nextMin = tk.v === '^' ? prec : prec + 1;
      const right = parseExpr(nextMin);
      left = { k: 'bin', op: tk.v, l: left, r: right };
    }
    return left;
  }

  const ast = parseExpr(0);
  if (pos !== tokens.length) throw ERR.VALUE;
  return ast;
}

const AST_CACHE = new Map();
function compile(formula) {
  let ast = AST_CACHE.get(formula);
  if (ast) return ast;
  let src = String(formula).trim();
  if (src.startsWith('=')) src = src.slice(1);
  try {
    ast = parse(tokenize(src));
  } catch (e) {
    ast = { k: 'err', v: isErr(e) ? e.code : '#NAME?' };
  }
  if (AST_CACHE.size > 200000) AST_CACHE.clear();
  AST_CACHE.set(formula, ast);
  return ast;
}

// ---------------------------------------------------------------- coercions (Excel rules)
function toNum(v) {
  if (isErr(v)) return v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isBlank(v)) return 0;
  const s = String(v).trim();
  if (s === '') return ERR.VALUE;
  const pct = /^(-?[\d,]*\.?\d+)\s*%$/.exec(s);
  if (pct) return Number(pct[1].replace(/,/g, '')) / 100;
  const n = Number(s.replace(/^\$/, '').replace(/,/g, ''));
  return Number.isFinite(n) ? n : ERR.VALUE;
}

function numToText(n) {
  if (!Number.isFinite(n)) return '#NUM!';
  if (Number.isInteger(n)) return String(n);
  const s = String(Number(n.toPrecision(15)));
  return s;
}

function toText(v) {
  if (isErr(v)) return v;
  if (isBlank(v)) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numToText(v);
  return String(v);
}

function toBool(v) {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isBlank(v)) return false;
  const s = String(v).toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  return ERR.VALUE;
}

// Excel ordering: numbers < text < booleans; text is case-insensitive
function typeRank(v) {
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return 1;
  if (typeof v === 'boolean') return 2;
  return 0;
}
function compareXl(a, b) {
  if (isBlank(a)) a = typeof b === 'string' ? '' : (typeof b === 'boolean' ? false : 0);
  if (isBlank(b)) b = typeof a === 'string' ? '' : (typeof a === 'boolean' ? false : 0);
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 1) {
    const sa = a.toLowerCase(), sb = b.toLowerCase();
    return sa < sb ? -1 : sa > sb ? 1 : 0;
  }
  if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

function round15(n) {
  return typeof n === 'number' && Number.isFinite(n) ? Number(n.toPrecision(15)) : n;
}

// ---------------------------------------------------------------- criteria (SUMIF / COUNTIF family)
function wildcardToRegex(pattern) {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) { re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}

function buildCriterion(crit) {
  if (isErr(crit)) return { error: crit };
  if (typeof crit === 'number' || typeof crit === 'boolean') {
    return { eqKey: criterionKey(crit), test: v => criterionKey(v) === criterionKey(crit) };
  }
  let s = isBlank(crit) ? '' : String(crit);
  let op = '=';
  const m = /^(<>|>=|<=|=|>|<)/.exec(s);
  if (m) { op = m[1]; s = s.slice(m[1].length); }
  const n = s.trim() !== '' ? Number(s) : NaN;
  const isNumCrit = Number.isFinite(n);
  const upper = s.toUpperCase();
  const boolCrit = upper === 'TRUE' ? true : upper === 'FALSE' ? false : null;

  if (op === '=' || op === '<>') {
    let match;
    let eqKey = null;
    if (s === '') {
      match = v => isBlank(v);
    } else if (isNumCrit) {
      match = v => (typeof v === 'number' && v === n) || (typeof v === 'string' && v.trim() !== '' && Number(v) === n);
      eqKey = `n:${n}`;
    } else if (boolCrit !== null) {
      match = v => v === boolCrit;
      eqKey = `b:${boolCrit}`;
    } else if (/[*?~]/.test(s)) {
      const re = wildcardToRegex(s);
      match = v => typeof v === 'string' && re.test(v);
    } else {
      const low = s.toLowerCase();
      match = v => typeof v === 'string' && v.toLowerCase() === low;
      eqKey = `s:${low}`;
    }
    if (op === '<>') return { test: v => !match(v) };
    return { eqKey, test: match };
  }
  // relational
  if (isNumCrit) {
    return {
      test: v => {
        if (typeof v !== 'number') return false;
        return op === '>' ? v > n : op === '<' ? v < n : op === '>=' ? v >= n : v <= n;
      }
    };
  }
  const low = s.toLowerCase();
  return {
    test: v => {
      if (typeof v !== 'string') return false;
      const x = v.toLowerCase();
      return op === '>' ? x > low : op === '<' ? x < low : op === '>=' ? x >= low : x <= low;
    }
  };
}

function criterionKey(v) {
  if (typeof v === 'number') return `n:${v}`;
  if (typeof v === 'boolean') return `b:${v}`;
  if (isBlank(v)) return 'blank';
  const s = String(v);
  const n = s.trim() !== '' ? Number(s) : NaN;
  if (Number.isFinite(n)) return `n:${n}`;
  return `s:${s.toLowerCase()}`;
}

// ---------------------------------------------------------------- engine
class FormulaEngine {
  constructor(workbook) {
    this.workbook = workbook;
    this.resetPass();
  }

  resetPass() {
    this.memo = new Map();       // "Sheet!A1" -> value
    this.visiting = new Set();
    this.indexCache = new Map(); // range key -> Map(criterionKey -> row offsets)
    this.colCache = new Map();   // range key -> values array
    this.usedRows = new Map();
  }

  // ---- cell access
  sheetOf(name) { return this.workbook.sheets[name]; }

  usedMaxRow(sheetName) {
    if (this.usedRows.has(sheetName)) return this.usedRows.get(sheetName);
    const sheet = this.sheetOf(sheetName);
    let max = 0;
    if (sheet && sheet.cells) {
      for (const coord of Object.keys(sheet.cells)) {
        const m = /(\d+)$/.exec(coord);
        if (m) { const r = Number(m[1]); if (r > max) max = r; }
      }
    }
    this.usedRows.set(sheetName, max);
    return max;
  }

  usedMaxCol(sheetName) {
    const sheet = this.sheetOf(sheetName);
    let max = 0;
    if (sheet && sheet.cells) {
      for (const coord of Object.keys(sheet.cells)) {
        const p = parseCoord(coord);
        if (p && p.col > max) max = p.col;
      }
    }
    return max;
  }

  cellValue(sheetName, row, col) {
    const sheet = this.sheetOf(sheetName);
    if (!sheet) return ERR.REF;
    const coord = indexToCol(col) + row;
    const cell = sheet.cells && sheet.cells[coord];
    if (!cell) return null;
    if (!cell.f) {
      const v = cell.v;
      if (typeof v === 'string' && v.startsWith('#') && cell.e) return new XlError(v);
      return v === undefined ? null : v;
    }
    const key = `${sheetName}!${coord}`;
    if (this.memo.has(key)) return this.memo.get(key);
    if (this.visiting.has(key)) return 0; // circular reference -> 0 like Excel
    this.visiting.add(key);
    let val;
    try {
      val = this.evalAst(compile(cell.f), sheetName, row, col);
      if (val instanceof RangeRef) val = this.implicitIntersect(val, row, col);
      if (Array.isArray(val)) val = val[0];
      if (val === null || val === undefined) val = 0;
      val = round15(val);
    } catch (e) {
      val = isErr(e) ? e : ERR.VALUE;
    } finally {
      this.visiting.delete(key);
    }
    this.memo.set(key, val);
    return val;
  }

  // ---- public API
  getCellValue(sheetName, coordStr) {
    const p = parseCoord(coordStr);
    if (!p) return '';
    const v = this.cellValue(sheetName, p.row, p.col);
    return isErr(v) ? v.code : (v === null ? '' : v);
  }

  evaluateFormula(formulaStr, sheetName, coord = 'A1') {
    const p = parseCoord(coord) || { row: 1, col: 1 };
    let v;
    try {
      v = this.evalAst(compile(formulaStr), sheetName, p.row, p.col);
      if (v instanceof RangeRef) v = this.implicitIntersect(v, p.row, p.col);
    } catch (e) {
      v = isErr(e) ? e : ERR.VALUE;
    }
    return isErr(v) ? v.code : (v === null || v === undefined ? 0 : round15(v));
  }

  // Evaluate a formula written for (anchorRow, anchorCol) as if it were at (row, col):
  // relative references shift, `$` ones stay (conditional formatting semantics).
  evaluateRelative(formula, sheetName, anchorRow, anchorCol, row, col) {
    const dr = row - anchorRow;
    const dc = col - anchorCol;
    const shift = node => {
      if (!node || typeof node !== 'object') return node;
      if (node.k === 'ref') return { ...node, r: node.ra ? node.r : node.r + dr, c: node.ca ? node.c : node.c + dc };
      if (node.k === 'range') {
        const out = { ...node };
        if (!node.wholeCol) { out.r1 = node.ra1 ? node.r1 : node.r1 + dr; out.r2 = node.ra2 ? node.r2 : node.r2 + dr; }
        if (!node.wholeRow) { out.c1 = node.ca1 ? node.c1 : node.c1 + dc; out.c2 = node.ca2 ? node.c2 : node.c2 + dc; }
        return out;
      }
      if (node.k === 'bin') return { ...node, l: shift(node.l), r: shift(node.r) };
      if (node.k === 'un' || node.k === 'pct') return { ...node, e: shift(node.e) };
      if (node.k === 'fn') return { ...node, args: node.args.map(shift) };
      return node;
    };
    let v;
    try {
      v = this.evalAst(shift(compile(String(formula).startsWith('=') ? formula : `=${formula}`)), sheetName, row, col);
      if (v instanceof RangeRef) v = this.implicitIntersect(v, row, col);
    } catch (e) {
      v = isErr(e) ? e : ERR.VALUE;
    }
    return v;
  }

  recalculateAll() {
    this.resetPass();
    for (const [sheetName, sheet] of Object.entries(this.workbook.sheets)) {
      if (!sheet.cells) continue;
      for (const [coord, cell] of Object.entries(sheet.cells)) {
        if (!cell || !cell.f) continue;
        const p = parseCoord(coord);
        if (!p) continue;
        const v = this.cellValue(sheetName, p.row, p.col);
        if (isErr(v)) { cell.v = v.code; cell.e = true; }
        else { cell.v = v === null ? 0 : v; if (cell.e) delete cell.e; }
      }
    }
  }

  // ---- evaluation
  implicitIntersect(range, row, col) {
    if (range.rows === 1 && range.cols === 1) return this.cellValue(range.sheet, range.r1, range.c1);
    if (range.cols === 1 && row >= range.r1 && row <= range.r2) return this.cellValue(range.sheet, row, range.c1);
    if (range.rows === 1 && col >= range.c1 && col <= range.c2) return this.cellValue(range.sheet, range.r1, col);
    return ERR.VALUE;
  }

  scalar(node, sheet, row, col) {
    const v = this.evalAst(node, sheet, row, col);
    if (v instanceof RangeRef) return this.implicitIntersect(v, row, col);
    return v;
  }

  resolveRange(node, sheet) {
    const s = node.sheet || sheet;
    if (!this.sheetOf(s)) throw ERR.REF;
    if (node.wholeCol) return new RangeRef(s, 1, node.c1, Math.max(1, this.usedMaxRow(s)), node.c2);
    if (node.wholeRow) return new RangeRef(s, node.r1, 1, node.r2, Math.max(1, this.usedMaxCol(s)));
    return new RangeRef(s, node.r1, node.c1, node.r2, node.c2);
  }

  resolveStructured(node, sheet, row) {
    const tables = this.workbook.tables || {};
    let table = node.table ? tables[node.table] : null;
    if (!table && node.table) {
      const key = Object.keys(tables).find(k => k.toLowerCase() === node.table.toLowerCase());
      table = key ? tables[key] : null;
    }
    if (!table && !node.table) table = Object.values(tables).find(t => t.sheet === sheet) || null;
    if (!table) throw ERR.REF;
    const [a, b] = table.ref.split(':').map(parseCoord);
    const headerRow = a.row;
    const lastRow = Math.max(b.row, this.usedMaxRow(table.sheet));
    const inner = node.spec.slice(1, -1).trim();
    const parts = inner.startsWith('[') ? inner.match(/\[[^\]]*\]/g).map(p => p.slice(1, -1).trim()) : [inner];
    let thisRow = false, all = false, headers = false, totals = false;
    const colNames = [];
    for (const p of parts) {
      const up = p.toUpperCase();
      if (up === '#THIS ROW' || p === '@') thisRow = true;
      else if (up === '#ALL') all = true;
      else if (up === '#HEADERS') headers = true;
      else if (up === '#DATA') { /* default */ } else if (up === '#TOTALS') totals = true;
      else if (p.startsWith('@')) { thisRow = true; colNames.push(p.slice(1).trim()); }
      else colNames.push(p.replace(/'(.)/g, '$1'));
    }
    const colIndex = name => {
      const i = table.columns.findIndex(c => String(c).trim().toLowerCase() === name.trim().toLowerCase());
      if (i < 0) throw ERR.REF;
      return a.col + i;
    };
    let c1 = a.col, c2 = b.col;
    if (colNames.length === 1) c1 = c2 = colIndex(colNames[0]);
    else if (colNames.length >= 2) { c1 = colIndex(colNames[0]); c2 = colIndex(colNames[colNames.length - 1]); }
    if (thisRow) return new RangeRef(table.sheet, row, c1, row, c2);
    if (headers) return new RangeRef(table.sheet, headerRow, c1, headerRow, c2);
    if (totals) throw ERR.REF;
    return new RangeRef(table.sheet, all ? headerRow : headerRow + 1, c1, lastRow, c2);
  }

  evalAst(node, sheet, row, col) {
    switch (node.k) {
      case 'num': case 'str': case 'bool': return node.v;
      case 'err': return new XlError(node.v);
      case 'missing': return null;
      case 'ref': {
        const s = node.sheet || sheet;
        if (!this.sheetOf(s)) return ERR.REF;
        return new RangeRef(s, node.r, node.c, node.r, node.c);
      }
      case 'range': return this.resolveRange(node, sheet);
      case 'sref': return this.resolveStructured(node, sheet, row);
      case 'name': return ERR.NAME;
      case 'un': {
        // unary plus is a no-op in Excel (=+A1 keeps text as text)
        if (node.op === '+') return this.evalAst(node.e, sheet, row, col);
        const v = toNum(this.scalar(node.e, sheet, row, col));
        if (isErr(v)) return v;
        return -v;
      }
      case 'pct': {
        const v = toNum(this.scalar(node.e, sheet, row, col));
        return isErr(v) ? v : v / 100;
      }
      case 'bin': return this.evalBinary(node, sheet, row, col);
      case 'fn': return this.callFunction(node.name, node.args, sheet, row, col);
      default: return ERR.VALUE;
    }
  }

  // operand of an operator: multi-cell ranges become arrays (element-wise math, like Excel arrays)
  operand(node, sheet, row, col) {
    const v = this.evalAst(node, sheet, row, col);
    if (v instanceof RangeRef) {
      if (v.rows === 1 && v.cols === 1) return this.cellValue(v.sheet, v.r1, v.c1);
      return this.rangeValues(v).map(x => (x === null ? null : x));
    }
    return v;
  }

  evalBinary(node, sheet, row, col) {
    const a = this.operand(node.l, sheet, row, col);
    const b = this.operand(node.r, sheet, row, col);
    if (Array.isArray(a) || Array.isArray(b)) {
      const len = Math.max(Array.isArray(a) ? a.length : 1, Array.isArray(b) ? b.length : 1);
      const out = new Array(len);
      for (let i = 0; i < len; i++) {
        const x = Array.isArray(a) ? (i < a.length ? a[i] : ERR.NA) : a;
        const y = Array.isArray(b) ? (i < b.length ? b[i] : ERR.NA) : b;
        out[i] = this.binScalar(node.op, x, y);
      }
      return out;
    }
    return this.binScalar(node.op, a, b);
  }

  binScalar(op, a, b) {
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (op === '&') return toText(a) + toText(b);
    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      const c = compareXl(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const x = toNum(a); if (isErr(x)) return x;
    const y = toNum(b); if (isErr(y)) return y;
    switch (op) {
      case '+': return x + y;
      case '-': return x - y;
      case '*': return x * y;
      case '/': return y === 0 ? ERR.DIV0 : x / y;
      case '^': { const r = Math.pow(x, y); return Number.isFinite(r) ? r : ERR.NUM; }
      default: return ERR.VALUE;
    }
  }

  // values of an argument, flattened; ranges skip text/bools for numeric aggregators
  *iterValues(node, sheet, row, col) {
    const v = this.evalAst(node, sheet, row, col);
    if (v instanceof RangeRef) {
      for (let r = v.r1; r <= v.r2; r++) {
        for (let c = v.c1; c <= v.c2; c++) yield { v: this.cellValue(v.sheet, r, c), fromRange: true };
      }
    } else if (Array.isArray(v)) {
      for (const x of v) yield { v: x, fromRange: true };
    } else {
      yield { v, fromRange: false };
    }
  }

  numbersOf(args, sheet, row, col, { ignoreErrors = false } = {}) {
    const out = [];
    for (const a of args) {
      for (const { v, fromRange } of this.iterValues(a, sheet, row, col)) {
        if (isErr(v)) { if (ignoreErrors) continue; throw v; }
        if (fromRange) { if (typeof v === 'number') out.push(v); continue; }
        if (isBlank(v)) { if (a.k !== 'missing') out.push(0); continue; }
        const n = toNum(v);
        if (isErr(n)) throw n;
        out.push(n);
      }
    }
    return out;
  }

  rangeArg(node, sheet, row, col) {
    const v = this.evalAst(node, sheet, row, col);
    if (v instanceof RangeRef) return v;
    throw ERR.VALUE;
  }

  rangeValues(range) {
    const key = range.key();
    let arr = this.colCache.get(key);
    if (arr) return arr;
    arr = [];
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) arr.push(this.cellValue(range.sheet, r, c));
    }
    this.colCache.set(key, arr);
    return arr;
  }

  rangeIndex(range) {
    const key = range.key();
    let idx = this.indexCache.get(key);
    if (idx) return idx;
    idx = new Map();
    const vals = this.rangeValues(range);
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (isErr(v)) continue;
      const k = criterionKey(v);
      let list = idx.get(k);
      if (!list) { list = []; idx.set(k, list); }
      list.push(i);
    }
    this.indexCache.set(key, idx);
    return idx;
  }

  // offsets (within the first criteria range shape) matching all criteria
  matchOffsets(pairs) {
    let candidates = null;
    for (const { range, crit } of pairs) {
      if (crit.error) throw crit.error;
      if (crit.eqKey) {
        const list = this.rangeIndex(range).get(crit.eqKey) || [];
        if (crit.eqKey.startsWith('s:') || crit.eqKey.startsWith('b:')) {
          candidates = intersect(candidates, list);
        } else {
          // numeric criteria also match numeric text -> index already merges by key
          candidates = intersect(candidates, list);
        }
      }
    }
    const size = pairs[0].range.rows * pairs[0].range.cols;
    const base = candidates || Array.from({ length: size }, (_, i) => i);
    return base.filter(i => pairs.every(({ range, crit }) => {
      if (crit.eqKey) return true;
      const v = this.rangeValues(range)[i];
      return !isErr(v) && crit.test(v);
    }));

    function intersect(a, b) {
      if (a === null) return b.slice();
      const set = new Set(b);
      return a.filter(x => set.has(x));
    }
  }

  sameShape(r, base) {
    return new RangeRef(r.sheet, r.r1, r.c1, r.r1 + base.rows - 1, r.c1 + base.cols - 1);
  }

  sumAt(range, offsets) {
    const vals = this.rangeValues(range);
    let total = 0;
    for (const i of offsets) {
      const v = vals[i];
      if (isErr(v)) throw v;
      if (typeof v === 'number') total += v;
    }
    return total;
  }

  // ---- functions
  callFunction(name, args, sheet, row, col) {
    const S = n => this.scalar(n, sheet, row, col);
    const N = n => { const v = toNum(S(n)); if (isErr(v)) throw v; return v; };
    const T = n => { const v = toText(S(n)); if (isErr(v)) throw v; return v; };
    const B = n => { const v = toBool(S(n)); if (isErr(v)) throw v; return v; };
    const arg = i => args[i] && args[i].k !== 'missing';

    try {
      switch (name) {
        // ---- logic
        case 'IF': {
          const c = B(args[0]);
          if (c) return arg(1) ? this.evalAst(args[1], sheet, row, col) : true;
          return args.length > 2 ? (arg(2) ? this.evalAst(args[2], sheet, row, col) : 0) : false;
        }
        case 'IFERROR': {
          const v = S(args[0]);
          return isErr(v) ? (arg(1) ? S(args[1]) : '') : v;
        }
        case 'IFNA': {
          const v = S(args[0]);
          return isErr(v) && v.code === '#N/A' ? (arg(1) ? S(args[1]) : '') : v;
        }
        case 'AND': case 'OR': {
          let res = name === 'AND';
          let seen = false;
          for (const a of args) for (const { v, fromRange } of this.iterValues(a, sheet, row, col)) {
            if (isErr(v)) return v;
            if (fromRange && typeof v !== 'boolean' && typeof v !== 'number') continue;
            const b = toBool(v); if (isErr(b)) return b;
            seen = true;
            res = name === 'AND' ? res && b : res || b;
          }
          return seen ? res : ERR.VALUE;
        }
        case 'NOT': return !B(args[0]);
        case 'TRUE': return true;
        case 'FALSE': return false;
        case 'ISERROR': return isErr(S(args[0]));
        case 'ISNA': { const v = S(args[0]); return isErr(v) && v.code === '#N/A'; }
        case 'ISBLANK': { const v = S(args[0]); return v === null || v === undefined; }
        case 'ISNUMBER': return typeof S(args[0]) === 'number';
        case 'ISTEXT': return typeof S(args[0]) === 'string';

        // ---- math / aggregates
        case 'SUM': return this.numbersOf(args, sheet, row, col).reduce((s, x) => s + x, 0);
        case 'PRODUCT': return this.numbersOf(args, sheet, row, col).reduce((s, x) => s * x, 1);
        case 'AVERAGE': case 'PROMEDIO': {
          const xs = this.numbersOf(args, sheet, row, col);
          return xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : ERR.DIV0;
        }
        case 'MIN': { const xs = this.numbersOf(args, sheet, row, col); return xs.length ? Math.min(...xs) : 0; }
        case 'MAX': { const xs = this.numbersOf(args, sheet, row, col); return xs.length ? Math.max(...xs) : 0; }
        case 'COUNT': {
          let n = 0;
          for (const a of args) for (const { v, fromRange } of this.iterValues(a, sheet, row, col)) {
            if (typeof v === 'number') n++;
            else if (!fromRange && !isErr(v) && !isBlank(v) && !isErr(toNum(v))) n++;
          }
          return n;
        }
        case 'COUNTA': {
          let n = 0;
          for (const a of args) for (const { v } of this.iterValues(a, sheet, row, col)) if (!(v === null || v === undefined)) n++;
          return n;
        }
        case 'COUNTBLANK': {
          let n = 0;
          for (const { v } of this.iterValues(args[0], sheet, row, col)) if (isBlank(v)) n++;
          return n;
        }
        case 'ABS': return Math.abs(N(args[0]));
        case 'INT': return Math.floor(N(args[0]));
        case 'MOD': { const d = N(args[1]); if (d === 0) return ERR.DIV0; const x = N(args[0]); return x - d * Math.floor(x / d); }
        case 'POWER': return Math.pow(N(args[0]), N(args[1]));
        case 'SQRT': { const x = N(args[0]); return x < 0 ? ERR.NUM : Math.sqrt(x); }
        case 'ROUND': case 'ROUNDUP': case 'ROUNDDOWN': {
          const x = N(args[0]);
          const d = arg(1) ? N(args[1]) : 0;
          const f = Math.pow(10, d);
          const y = Math.abs(x) * f;
          const r = name === 'ROUND' ? Math.round(Number((y).toPrecision(15))) : name === 'ROUNDUP' ? Math.ceil(Number(y.toPrecision(15))) : Math.floor(Number(y.toPrecision(15)));
          return Math.sign(x) * r / f;
        }
        case 'CEILING': { const x = N(args[0]); const s = arg(1) ? N(args[1]) : 1; return s === 0 ? 0 : Math.ceil(x / s) * s; }
        case 'FLOOR': { const x = N(args[0]); const s = arg(1) ? N(args[1]) : 1; return s === 0 ? ERR.DIV0 : Math.floor(x / s) * s; }

        case 'SUMPRODUCT': {
          const ranges = args.map(a => this.evalAst(a, sheet, row, col));
          const arrays = ranges.map(r => r instanceof RangeRef ? this.rangeValues(r) : (Array.isArray(r) ? r : [r]));
          const len = arrays[0].length;
          if (arrays.some(a => a.length !== len)) return ERR.VALUE;
          let total = 0;
          for (let i = 0; i < len; i++) {
            let p = 1;
            for (const a of arrays) { const v = a[i]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : 0; }
            total += p;
          }
          return total;
        }

        case 'SUMIF': case 'COUNTIF': case 'AVERAGEIF': {
          const range = this.rangeArg(args[0], sheet, row, col);
          const crit = buildCriterion(S(args[1]));
          const offsets = this.matchOffsets([{ range, crit }]);
          if (name === 'COUNTIF') return offsets.length;
          const sumRange = arg(2) ? this.sameShape(this.rangeArg(args[2], sheet, row, col), range) : range;
          const total = this.sumAt(sumRange, offsets);
          if (name === 'SUMIF') return total;
          const nums = offsets.filter(i => typeof this.rangeValues(sumRange)[i] === 'number').length;
          return nums ? total / nums : ERR.DIV0;
        }
        case 'SUMIFS': case 'AVERAGEIFS': case 'COUNTIFS': {
          const first = name === 'COUNTIFS' ? 0 : 1;
          const pairs = [];
          for (let i = first; i < args.length; i += 2) {
            const range = this.rangeArg(args[i], sheet, row, col);
            pairs.push({ range, crit: buildCriterion(S(args[i + 1])) });
          }
          const base = pairs[0].range;
          if (pairs.some(p => p.range.rows !== base.rows || p.range.cols !== base.cols)) return ERR.VALUE;
          const offsets = this.matchOffsets(pairs);
          if (name === 'COUNTIFS') return offsets.length;
          const sumRange = this.rangeArg(args[0], sheet, row, col);
          if (sumRange.rows !== base.rows || sumRange.cols !== base.cols) return ERR.VALUE;
          const total = this.sumAt(sumRange, offsets);
          if (name === 'SUMIFS') return total;
          const nums = offsets.filter(i => typeof this.rangeValues(sumRange)[i] === 'number').length;
          return nums ? total / nums : ERR.DIV0;
        }

        case 'SUBTOTAL': case 'AGGREGATE': {
          let fn = N(args[0]);
          let ignoreErrors = false;
          let refs = args.slice(1);
          if (name === 'AGGREGATE') {
            const opt = arg(1) ? N(args[1]) : 0;
            ignoreErrors = [2, 3, 6, 7].includes(opt);
            refs = args.slice(2);
          } else if (fn > 100) fn -= 100;
          const xs = this.numbersOf(refs, sheet, row, col, { ignoreErrors });
          switch (fn) {
            case 1: return xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : ERR.DIV0;
            case 2: return xs.length;
            case 3: {
              let n = 0;
              for (const a of refs) for (const { v } of this.iterValues(a, sheet, row, col)) {
                if (isErr(v) && ignoreErrors) continue;
                if (!(v === null || v === undefined)) n++;
              }
              return n;
            }
            case 4: return xs.length ? Math.max(...xs) : 0;
            case 5: return xs.length ? Math.min(...xs) : 0;
            case 6: return xs.reduce((s, x) => s * x, 1);
            case 9: return xs.reduce((s, x) => s + x, 0);
            default: return ERR.VALUE;
          }
        }

        // ---- lookup
        case 'VLOOKUP': case 'HLOOKUP': {
          const key = S(args[0]);
          if (isErr(key)) return key;
          const table = this.rangeArg(args[1], sheet, row, col);
          const idx = N(args[2]);
          const approx = arg(3) ? B(args[3]) : true;
          const vertical = name === 'VLOOKUP';
          const count = vertical ? table.rows : table.cols;
          if (idx < 1 || idx > (vertical ? table.cols : table.rows)) return ERR.REF;
          const at = i => vertical ? this.cellValue(table.sheet, table.r1 + i, table.c1) : this.cellValue(table.sheet, table.r1, table.c1 + i);
          const result = i => vertical ? this.cellValue(table.sheet, table.r1 + i, table.c1 + idx - 1) : this.cellValue(table.sheet, table.r1 + idx - 1, table.c1 + i);
          if (!approx) {
            const crit = typeof key === 'string' && /[*?]/.test(key) ? wildcardToRegex(key) : null;
            for (let i = 0; i < count; i++) {
              const v = at(i);
              if (crit ? (typeof v === 'string' && crit.test(v)) : (!isBlank(v) && compareXl(v, key) === 0 && typeRank(v) === typeRank(key))) {
                const r = result(i); return r === null ? 0 : r;
              }
            }
            return ERR.NA;
          }
          let found = -1;
          for (let i = 0; i < count; i++) {
            const v = at(i);
            if (isBlank(v) || typeRank(v) !== typeRank(key)) continue;
            if (compareXl(v, key) <= 0) found = i; else break;
          }
          if (found < 0) return ERR.NA;
          const r = result(found); return r === null ? 0 : r;
        }
        case 'MATCH': {
          const key = S(args[0]);
          if (isErr(key)) return key;
          const range = this.rangeArg(args[1], sheet, row, col);
          const type = arg(2) ? N(args[2]) : 1;
          const vals = this.rangeValues(range);
          if (type === 0) {
            const re = typeof key === 'string' && /[*?]/.test(key) ? wildcardToRegex(key) : null;
            for (let i = 0; i < vals.length; i++) {
              const v = vals[i];
              if (re ? (typeof v === 'string' && re.test(v)) : (!isBlank(v) && typeRank(v) === typeRank(key) && compareXl(v, key) === 0)) return i + 1;
            }
            return ERR.NA;
          }
          let found = -1;
          for (let i = 0; i < vals.length; i++) {
            const v = vals[i];
            if (isBlank(v) || typeRank(v) !== typeRank(key)) continue;
            const c = compareXl(v, key);
            if (type > 0 ? c <= 0 : c >= 0) found = i; else break;
          }
          return found < 0 ? ERR.NA : found + 1;
        }
        case 'INDEX': {
          const range = this.rangeArg(args[0], sheet, row, col);
          let r = arg(1) ? N(args[1]) : 0;
          let c = arg(2) ? N(args[2]) : 0;
          if (range.rows === 1 && args.length === 2) { c = r; r = 1; }
          if (r === 0 && c === 0) return range;
          if (r < 0 || c < 0 || r > range.rows || c > range.cols) return ERR.REF;
          if (r === 0) return new RangeRef(range.sheet, range.r1, range.c1 + c - 1, range.r2, range.c1 + c - 1);
          if (c === 0) c = 1;
          return new RangeRef(range.sheet, range.r1 + r - 1, range.c1 + c - 1, range.r1 + r - 1, range.c1 + c - 1);
        }

        // ---- text
        case 'CONCATENATE': case 'CONCAT': {
          let out = '';
          for (const a of args) for (const { v } of this.iterValues(a, sheet, row, col)) { if (isErr(v)) return v; out += toText(v); }
          return out;
        }
        case 'TEXTJOIN': {
          const delim = T(args[0]);
          const skip = B(args[1]);
          const parts = [];
          for (const a of args.slice(2)) for (const { v } of this.iterValues(a, sheet, row, col)) {
            if (isErr(v)) return v;
            const t = toText(v);
            if (skip && t === '') continue;
            parts.push(t);
          }
          return parts.join(delim);
        }
        case 'LEN': return T(args[0]).length;
        case 'UPPER': return T(args[0]).toUpperCase();
        case 'LOWER': return T(args[0]).toLowerCase();
        case 'TRIM': return T(args[0]).replace(/ +/g, ' ').trim();
        case 'LEFT': return T(args[0]).slice(0, arg(1) ? N(args[1]) : 1);
        case 'RIGHT': { const s = T(args[0]); const n = arg(1) ? N(args[1]) : 1; return n <= 0 ? '' : s.slice(-n); }
        case 'MID': { const s = T(args[0]); const st = N(args[1]); const n = N(args[2]); return st < 1 ? ERR.VALUE : s.substr(st - 1, n); }
        case 'FIND': case 'SEARCH': {
          const find = T(args[0]); const within = T(args[1]); const start = arg(2) ? N(args[2]) : 1;
          const i = name === 'FIND' ? within.indexOf(find, start - 1) : within.toLowerCase().indexOf(find.toLowerCase(), start - 1);
          return i < 0 ? ERR.VALUE : i + 1;
        }
        case 'SUBSTITUTE': return T(args[0]).split(T(args[1])).join(T(args[2]));
        case 'VALUE': { const n = toNum(T(args[0])); return n; }
        case 'TEXT': return toText(S(args[0]));

        // ---- dates (Excel serial numbers)
        case 'TODAY': {
          const d = new Date();
          return Math.floor((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
        }
        case 'DATE': return (Date.UTC(N(args[0]), N(args[1]) - 1, N(args[2])) - Date.UTC(1899, 11, 30)) / 86400000;
        case 'YEAR': case 'MONTH': case 'DAY': case 'WEEKDAY': {
          const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(N(args[0])) * 86400000);
          if (name === 'YEAR') return d.getUTCFullYear();
          if (name === 'MONTH') return d.getUTCMonth() + 1;
          if (name === 'DAY') return d.getUTCDate();
          return d.getUTCDay() + 1;
        }

        default: return ERR.NAME;
      }
    } catch (e) {
      if (isErr(e)) return e;
      throw e;
    }
  }
}

const ExcelUtils = { colToIndex, indexToCol, parseCoord, ERROR_CODES, isErr, compareXl, toNum };

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FormulaEngine, ...ExcelUtils };
} else {
  window.FormulaEngine = FormulaEngine;
  window.ExcelUtils = ExcelUtils;
}
