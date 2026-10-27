// ============================================================================
// REQUERIMIENTO VS PROYECCION - sheet "Append" (same columns as the Excel's Append1)
//   * Requerimiento: every order of Reunion's PegarData (Total de tallos 2 = Total Tallos)
//   * Proyeccion: Plan de cosecha's last tab (Total Tallos 1): one row per captured day,
//     or per week when that week has no days captured. Distribuidor / Cust Name / No. = "CORTE"
//   * Calculated like the Excel: Semana Prod / Semana Load = WEEKNUM(date, 16) (Sat..Fri weeks),
//     Status orden = ESTIMADO when No. contains "Est", Flor General = Flag!J:K lookup ("*FLOR*").
//     (dif = Total Tallos 1 - Total de tallos 2 is a calculated field of the pivot tables, as in the Excel)
// ============================================================================

const CosechaCor = require('./cosechaCor');

const HEADERS = ['STATUS', 'Distribuidor', 'Cust Name', 'No.', 'CATEGORIA', 'FLOR', 'Flor Color', 'Descripcion', 'Load Date',
  '#cajas', 'Prod', 'Term', '#Ramos', 'Ramos/Cajas', 'Total Tallos 1', 'Fecha Produccion', 'Caja', 'prioridad', 'Semana Prod',
  'Status orden (Formula)', 'Flor General (Formula)', 'Semana Load', 'Status orden', 'Flor General', 'Total de tallos 2',
  'Finca', 'Flor y Color', 'Flor 123'];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i]));
const EPOCH = Date.UTC(1899, 11, 30);

const indexToCol = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; };
const colToIndex = col => col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const isNum = v => typeof v === 'number' && Number.isFinite(v);
const num = v => (isNum(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) ? Number(v) : 0));
const text = v => (v === undefined || v === null ? '' : String(v));
const upper = v => text(v).toUpperCase();

// Excel WEEKNUM(date, 16): weeks start on Saturday, week 1 holds January 1
function weeknum16(serial) {
  if (!isNum(serial)) return null;
  const d = new Date(EPOCH + Math.floor(serial) * 86400000);
  const jan1 = Date.UTC(d.getUTCFullYear(), 0, 1);
  const offset = (new Date(jan1).getUTCDay() - 6 + 7) % 7; // days from the Saturday that opens week 1
  const doy = Math.round((d.getTime() - jan1) / 86400000);
  return Math.floor((doy + offset) / 7) + 1;
}

// Flag!J:K: the first "Palabra Clave" (J) that contains the flower name gives its "Resumen" (K).
// Requerimiento Vs Proyeccion has its own Flag (from its Excel); Reunion's is the fallback.
function flagLookup(reunionBook, rvpBook) {
  const flag = (rvpBook && rvpBook.sheets && rvpBook.sheets.Flag) || reunionBook.sheets.Flag;
  const rows = [];
  if (flag) {
    let last = 1;
    for (const coord of Object.keys(flag.cells)) {
      const m = /^J(\d+)$/.exec(coord);
      if (m && Number(m[1]) > last) last = Number(m[1]);
    }
    for (let r = 1; r <= last; r++) {
      const key = upper((flag.cells[`J${r}`] || {}).v);
      if (key) rows.push([key, text((flag.cells[`K${r}`] || {}).v)]);
    }
  }
  const cache = new Map();
  return flor => {
    const f = upper(flor);
    if (cache.has(f)) return cache.get(f);
    const hit = rows.find(([key]) => key.includes(f));
    const out = hit ? hit[1] : 'SIN CONSOLIDAR';
    cache.set(f, out);
    return out;
  };
}

function lastRowOf(cells, col) {
  let last = 0;
  const re = new RegExp(`^${col}(\\d+)$`);
  for (const coord of Object.keys(cells)) {
    const m = re.exec(coord);
    if (m && Number(m[1]) > last && text(cells[coord].v) !== '') last = Number(m[1]);
  }
  return last;
}

// "MATRICARIA YELLOWX 5" -> "MATRICARIA YELLOW": a Flor 123 without its pack
const florYColor = flor123 => text(flor123).replace(/\s*X\s*\d+\s*$/i, '').replace(/\s+/g, ' ').trim().toUpperCase() || null;

// Product names (Flor 123) known by Plan de cosecha and Reunion: normalized name -> name to show
const normName = v => text(v).replace(/\s+/g, ' ').trim().toUpperCase();
function productNames(reunionBook, cosechaBook) {
  const names = new Map();
  const reu = reunionBook.sheets.Reunion;
  if (reu) {
    const last = lastRowOf(reu.cells, 'A');
    for (let r = 3; r <= last; r++) {
      const n = text((reu.cells[`A${r}`] || {}).v).trim();
      if (n) names.set(normName(n), n);
    }
  }
  const found = cosechaBook && CosechaCor.lastCosechaSheet(cosechaBook);
  if (found) {
    const last = lastRowOf(found.sheet.cells, found.florCol);
    for (let r = 5; r <= last; r++) {
      const n = text((found.sheet.cells[`${found.florCol}${r}`] || {}).v).trim();
      if (n) names.set(normName(n), n);
    }
  }
  return names;
}

function requerimientoRows(reunionBook, flagOf, products) {
  const sheet = reunionBook.sheets.PegarData;
  if (!sheet) return [];
  const c = sheet.cells;
  const v = (col, r) => (c[`${col}${r}`] || {}).v;
  const out = [];
  const last = lastRowOf(c, 'E');
  for (let r = 2; r <= last; r++) {
    const flor = v('E', r);
    if (text(flor).trim() === '') continue;
    const row = new Array(HEADERS.length).fill(null);
    row[COL.STATUS] = 'Requerimiento';
    row[COL.Distribuidor] = v('A', r) ?? null;
    row[COL['Cust Name']] = v('B', r) ?? null;
    row[COL['No.']] = v('C', r) ?? null;
    row[COL.CATEGORIA] = v('D', r) ?? null;
    row[COL.FLOR] = flor;
    row[COL['Flor Color']] = v('F', r) ?? null;
    row[COL.Descripcion] = v('G', r) ?? null;
    row[COL['Load Date']] = v('H', r) ?? null;
    row[COL['#cajas']] = v('I', r) ?? null;
    row[COL.Prod] = v('J', r) ?? null;
    row[COL.Term] = v('K', r) ?? null;
    row[COL['#Ramos']] = v('L', r) ?? null;
    row[COL['Ramos/Cajas']] = v('M', r) ?? null;
    row[COL['Fecha Produccion']] = v('O', r) ?? null;
    row[COL.Caja] = v('P', r) ?? null;
    row[COL.prioridad] = v('Q', r) ?? null;
    row[COL['Semana Prod']] = weeknum16(v('O', r));
    row[COL['Status orden (Formula)']] = /est/i.test(text(v('C', r))) ? 'ESTIMADO' : 'FIRME';
    row[COL['Flor General (Formula)']] = flagOf(flor);
    row[COL['Semana Load']] = weeknum16(v('H', r));
    row[COL['Total de tallos 2']] = num(v('N', r));
    // Flor 123 of the order, as Reunion's Req finds it: pack name (FLOR Color X TxR2, TxR2 = txr_orden),
    // then FLOR Color, then FLOR
    const pack = normName(`${text(flor)} ${text(v('F', r))}X ${text(v('R', r))}`);
    const withColor = normName(`${text(flor)} ${text(v('F', r))}`);
    const plain = normName(flor);
    const key = [pack, withColor, plain].find(k => products.has(k));
    row[COL['Flor 123']] = key ? products.get(key) : null;
    // middle level (all the packs of a flower + color): from its Flor 123 like the CORTE rows, so both land
    // on the same line (e.g. MALELUCA orders come as "MALELUCA | GREEN" but the product is "MALELUCA")
    row[COL['Flor y Color']] = key ? florYColor(products.get(key)) : (withColor || null);
    out.push(row);
  }
  return out;
}

function proyeccionRows(cosechaBook, flagOf) {
  const found = CosechaCor.lastCosechaSheet(cosechaBook);
  if (!found) return [];
  const { name, sheet, florCol } = found;
  const c = sheet.cells;
  const v = (col, r) => (c[`${col}${r}`] || {}).v;
  const head = {};
  const weekCols = [];
  for (const [coord, cell] of Object.entries(c)) {
    const m = /^([A-Z]+)4$/.exec(coord);
    if (!m) continue;
    const h = upper(cell.v).trim();
    if (h === 'FLOR' && !head.flor) head.flor = m[1];
    else if (h === 'COLOR' && !head.color) head.color = m[1];
    else if (h === 'FINCA' && !head.finca) head.finca = m[1];
    if (!cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) weekCols.push({ col: m[1], week: cell.v });
  }
  weekCols.sort((a, b) => colToIndex(a.col) - colToIndex(b.col));
  const yw = /(\d{2})(\d{2})\D*$/.exec(name.trim());
  const year = yw ? 2000 + Number(yw[1]) : new Date().getFullYear();
  const sheetWeek = yw ? Number(yw[2]) : 1;
  const groups = (sheet.weekDays && sheet.weekDays.groups) || {};
  const out = [];
  const last = lastRowOf(c, florCol);
  for (let r = 5; r <= last; r++) {
    const flor123 = text(v(florCol, r)).trim();
    if (!flor123) continue;
    const flor = head.flor ? v(head.flor, r) : flor123;
    const base = () => {
      const row = new Array(HEADERS.length).fill(null);
      row[COL.STATUS] = 'Proyeccion';
      row[COL.Distribuidor] = 'CORTE';
      row[COL['Cust Name']] = 'CORTE';
      row[COL['No.']] = 'CORTE';
      row[COL.FLOR] = flor;
      row[COL['Flor Color']] = head.color ? v(head.color, r) ?? null : null;
      row[COL.Descripcion] = 'CORTE';
      row[COL['Status orden (Formula)']] = 'FIRME';
      row[COL['Flor General (Formula)']] = flagOf(flor);
      row[COL['Total de tallos 2']] = 0;
      row[COL.Finca] = head.finca ? v(head.finca, r) ?? null : null;
      row[COL['Flor 123']] = flor123;
      row[COL['Flor y Color']] = florYColor(flor123);
      return row;
    };
    for (const w of weekCols) {
      const y = w.week < sheetWeek - 26 ? year + 1 : year;
      const g = groups[w.col];
      const days = g && Array.isArray(g.cols) ? g.cols.map(col => num(v(col, r))) : [];
      if (days.some(x => x)) {
        // captured by day: one row per day with its production date
        const sat = saturdaySerial(y, w.week);
        days.forEach((qty, i) => {
          if (!qty) return;
          const row = base();
          row[COL['Total Tallos 1']] = qty;
          row[COL['Fecha Produccion']] = sat + i;
          row[COL['Semana Prod']] = w.week;
          out.push(row);
        });
      } else {
        const qty = num(v(w.col, r));
        if (!qty) continue;
        const row = base();
        row[COL['Total Tallos 1']] = qty;
        row[COL['Semana Prod']] = w.week;
        out.push(row);
      }
    }
  }
  return out;
}

// Saturday that opens week `week` of `year` (same calendar as Plan de cosecha)
function saturdaySerial(year, week) {
  const jan4 = Date.UTC(year, 0, 4);
  const dow = (new Date(jan4).getUTCDay() + 6) % 7;
  const monday = jan4 - dow * 86400000 + (week - 1) * 7 * 86400000;
  return Math.round((monday - 2 * 86400000 - EPOCH) / 86400000);
}

// Today's date in Tijuana as an Excel serial
function todaySerial() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Tijuana', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const [y, m, d] = parts.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}
function serialOfIso(iso) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Tijuana', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(iso));
  const [y, m, d] = parts.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}

// Inventario (Reunion): stock already cut (STATUS "Inventario", Total Tallos 1) counts in the week the inventory was taken
// (the week, Sat..Fri, of the day "Actualizar Inventario" ran; today if it was never run)
function inventarioRows(reunionBook, flagOf, products) {
  const sheet = reunionBook.sheets.Inventario;
  if (!sheet) return [];
  const c = sheet.cells;
  const v = (col, r) => (c[`${col}${r}`] || {}).v;
  const day = sheet.updatedAt ? serialOfIso(sheet.updatedAt) : todaySerial();
  const week = weeknum16(day);
  const out = [];
  const last = lastRowOf(c, 'A');
  for (let r = 2; r <= last; r++) {
    const flor = v('A', r);
    const qty = num(v('E', r));
    if (text(flor).trim() === '' || !qty) continue;
    const color = v('B', r);
    const pack = normName(`${text(flor)} ${text(color)}X ${text(v('D', r))}`);
    const withColor = normName(`${text(flor)} ${text(color)}`);
    const key = [pack, withColor, normName(flor)].find(k => products.has(k));
    const row = new Array(HEADERS.length).fill(null);
    row[COL.STATUS] = 'Inventario';
    row[COL.Distribuidor] = 'INVENTARIO';
    row[COL['Cust Name']] = 'INVENTARIO';
    row[COL['No.']] = 'INVENTARIO';
    row[COL.FLOR] = flor;
    row[COL['Flor Color']] = color ?? null;
    row[COL.Descripcion] = 'AINVENTARIO';
    row[COL['Total Tallos 1']] = qty;
    row[COL['Semana Prod']] = week;
    row[COL['Status orden (Formula)']] = 'FIRME';
    row[COL['Flor General (Formula)']] = flagOf(flor);
    row[COL['Total de tallos 2']] = 0;
    row[COL['Flor 123']] = key ? products.get(key) : null;
    row[COL['Flor y Color']] = key ? florYColor(products.get(key)) : (withColor || null);
    out.push(row);
  }
  return out;
}

function buildAppendRows(reunionBook, cosechaBook, rvpBook) {
  const flagOf = flagLookup(reunionBook, rvpBook);
  const products = productNames(reunionBook, cosechaBook);
  return [
    ...proyeccionRows(cosechaBook, flagOf),
    ...inventarioRows(reunionBook, flagOf, products),
    ...requerimientoRows(reunionBook, flagOf, products)
  ];
}

const HEAD_STYLE = { bold: true, bg: '#1f4e78', color: '#ffffff', align: 'center', border: { top: true, right: true, bottom: true, left: true } };
const DATE_COLS = new Set([COL['Load Date'], COL['Fecha Produccion']]);

// Formula columns of the order rows (like the Excel's table columns). {r} is the row.
// If someone edits one of these formulas in the sheet, the new formula is used for the whole column.
const DEFAULT_FORMULAS = {
  'Semana Prod': '=WEEKNUM(P{r},16)',
  'Status orden (Formula)': '=IF(ISNUMBER(SEARCH("Est***",D{r})),"ESTIMADO","FIRME")',
  'Semana Load': '=WEEKNUM(I{r},16)'
};
function formulaTemplates(prevSheet) {
  const out = { ...DEFAULT_FORMULAS, ...((prevSheet && prevSheet.formulaTemplates) || {}) };
  if (!prevSheet || !prevSheet.cells) return out;
  for (const header of Object.keys(DEFAULT_FORMULAS)) {
    const col = indexToCol(COL[header] + 1);
    const re = new RegExp(`^${col}(\\d+)$`);
    for (const [coord, cell] of Object.entries(prevSheet.cells)) {
      const m = re.exec(coord);
      if (!m || !cell.f) continue;
      const r = m[1];
      const tmpl = cell.f.replace(new RegExp(`(\\$?[A-Z]{1,3})${r}(?!\\d)`, 'g'), '$1{r}');
      if (tmpl !== out[header]) { out[header] = tmpl; break; } // edited by hand
    }
  }
  return out;
}

function buildAppendSheet(reunionBook, cosechaBook, rvpBook) {
  const rows = buildAppendRows(reunionBook, cosechaBook, rvpBook);
  const prev = rvpBook && rvpBook.sheets && rvpBook.sheets.Append;
  const templates = formulaTemplates(prev);
  const formulaCols = new Map(Object.keys(DEFAULT_FORMULAS).map(h => [COL[h], h]));
  const cells = {};
  HEADERS.forEach((h, i) => { cells[`${indexToCol(i + 1)}1`] = { v: h, s: HEAD_STYLE }; });
  rows.forEach((row, ri) => {
    const r = ri + 2;
    const isReq = row[COL.STATUS] === 'Requerimiento';
    row.forEach((val, ci) => {
      const header = formulaCols.get(ci);
      if (isReq && header) {
        // order rows carry the formula (its value is filled in for the default one; a changed one is recalculated)
        const f = templates[header].replace(/\{r\}/g, r);
        cells[`${indexToCol(ci + 1)}${r}`] = { f, v: templates[header] === DEFAULT_FORMULAS[header] ? (val ?? '') : '' };
        return;
      }
      if (val === null || val === undefined || val === '') return;
      const cell = { v: val };
      if (DATE_COLS.has(ci) && isNum(val)) cell.s = { nf: 'mm-dd-yy' };
      cells[`${indexToCol(ci + 1)}${r}`] = cell;
    });
  });
  const lastCol = indexToCol(HEADERS.length);
  const colWidths = {};
  HEADERS.forEach((h, i) => { colWidths[indexToCol(i + 1)] = Math.max(10, Math.min(28, h.length + 3)); });
  Object.assign(colWidths, { B: 24, C: 26, G: 30, H: 30, U: 22 });
  return {
    sheet: {
      maxRow: rows.length + 30,
      maxCol: HEADERS.length + 2,
      colWidths,
      rowHeights: {},
      merges: [],
      freeze: { rows: 1, cols: 0 },
      filters: [{ ref: `A1:${lastCol}${rows.length + 1}`, columns: {} }],
      cells,
      formulaTemplates: templates,
      generated: 'Se arma sola desde PegarData (Reunion) y la última pestaña de Plan de cosecha'
    },
    count: rows.length
  };
}

module.exports = { buildAppendSheet, buildAppendRows, HEADERS, weeknum16 };
