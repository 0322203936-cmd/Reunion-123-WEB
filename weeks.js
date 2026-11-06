// ============================================================================
// DÍAS POR SEMANA (Plan de cosecha) - expandable Sat..Fri capture columns
//   * Week columns = row 4 headers holding a week number (1-53) on "P Cosecha" sheets
//   * sheet.weekDays = { start, lastRow, groups: { <weekCol>: { week, year, cols:[7 cols], open } } }
//     The 7 day columns live after the used area (they never shift existing cells) and are
//     displayed right after their week while the group is open.
//   * The week cell keeps the total they typed (what they have to spread over the days); it turns red
//     while its days don't add up to it. Pivots / Append / Cor take the days when the row has any,
//     otherwise the week total (appendBuilder.js). Old rows may still hold =SUM(<its 7 days>).
// ============================================================================

(function () {
  const { indexToCol, colToIndex, parseCoord } = window.ExcelUtils;
  const App = () => window.MauriApp;
  const HEADER_ROW = 4;
  const FIRST_ROW = 5;
  const DAY_NAMES = ['Sáb', 'Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie'];
  const MONTHS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const DAY_COL_WIDTH = 12.5;
  // which weeks show their days is personal (this computer only), like the filters: it is never saved
  const OPEN_KEY = 'reunion123_week_open';
  let openState = {};
  try { openState = JSON.parse(localStorage.getItem(OPEN_KEY) || '{}') || {}; } catch (e) { openState = {}; }
  const openId = (sheetName, colStr) => `${String(sheetName).trim()}|${colStr}`;
  const isOpen = (sheetName, colStr) => !!openState[openId(sheetName, colStr)];
  function setOpen(sheetName, colStr, open) {
    if (open) openState[openId(sheetName, colStr)] = 1;
    else delete openState[openId(sheetName, colStr)];
    try { localStorage.setItem(OPEN_KEY, JSON.stringify(openState)); } catch (e) { /* storage unavailable */ }
  }

  const isWeekSheet = name => /^\s*p\s*cosecha\s*\d{4}/i.test(name || '');
  const sheetYearWeek = name => {
    const m = /(\d{2})(\d{2})\D*$/.exec(String(name).trim());
    return m ? { year: 2000 + Number(m[1]), week: Number(m[2]) } : null;
  };

  // Saturday that opens week `week` of `year`: the Monday of ISO week minus 2 days
  function weekSaturday(year, week) {
    const jan4 = new Date(Date.UTC(year, 0, 4));
    const isoMonday = new Date(jan4);
    isoMonday.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() + 6) % 7) + (week - 1) * 7);
    isoMonday.setUTCDate(isoMonday.getUTCDate() - 2);
    return isoMonday;
  }
  function dayLabels(year, week) {
    const sat = weekSaturday(year, week);
    return DAY_NAMES.map((n, i) => {
      const d = new Date(sat);
      d.setUTCDate(sat.getUTCDate() + i);
      return `${n} ${String(d.getUTCDate()).padStart(2, '0')}-${MONTHS[d.getUTCMonth()]}`;
    });
  }

  // a week column has a plain week number (1-53) in the header row
  function isWeekHeader(sheet, colStr) {
    const cell = sheet.cells[`${colStr}${HEADER_ROW}`];
    return !!cell && !cell.f && typeof cell.v === 'number' && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53;
  }
  function weekColumns(sheet) {
    const out = [];
    for (let c = 1; c <= (sheet.maxCol || 20); c++) {
      const colStr = indexToCol(c);
      if (isWeekHeader(sheet, colStr)) out.push({ col: c, colStr, week: sheet.cells[`${colStr}${HEADER_ROW}`].v });
    }
    return out;
  }

  function lastDataRow(sheet) {
    let last = FIRST_ROW;
    for (const coord of Object.keys(sheet.cells || {})) {
      const p = parseCoord(coord);
      if (p && p.col === 1 && p.row > last && sheet.cells[coord].v !== undefined && sheet.cells[coord].v !== '') last = p.row;
    }
    return last;
  }

  function usedMaxCol(sheet) {
    let max = sheet.maxCol || 1;
    for (const coord of Object.keys(sheet.cells || {})) {
      const p = parseCoord(coord);
      if (p && p.col > max) max = p.col;
    }
    return max; // column widths are ignored: Excel lists them up to XFD
  }

  // allocate the day columns of every week of the sheet (once), with headers and the week's formats
  function ensureGroups(sheetName) {
    const sheet = App().getWorkbook().sheets[sheetName];
    if (sheet.weekDays) return sheet.weekDays;
    const yw = sheetYearWeek(sheetName);
    const weeks = weekColumns(sheet);
    const start = usedMaxCol(sheet) + 2;
    const lastRow = lastDataRow(sheet);
    const wd = { start: indexToCol(start), lastRow, groups: {} };
    let next = start;
    for (const w of weeks) {
      const year = yw ? yw.year + (w.week < yw.week - 26 ? 1 : 0) : new Date().getFullYear();
      const cols = [];
      const labels = dayLabels(year, w.week);
      const headStyle = (sheet.cells[`${w.colStr}${HEADER_ROW}`] || {}).s;
      for (let i = 0; i < 7; i++, next++) {
        const colStr = indexToCol(next);
        cols.push(colStr);
        sheet.cells[`${colStr}${HEADER_ROW}`] = { v: labels[i], ...(headStyle ? { s: JSON.parse(JSON.stringify(headStyle)) } : {}) };
        for (let r = FIRST_ROW; r <= lastRow; r++) {
          const ws = (sheet.cells[`${w.colStr}${r}`] || {}).s || sheet.defaultStyle;
          if (ws && Object.keys(ws).length) sheet.cells[`${colStr}${r}`] = { s: JSON.parse(JSON.stringify(ws)) };
        }
        (sheet.colWidths = sheet.colWidths || {})[colStr] = DAY_COL_WIDTH;
      }
      wd.groups[w.colStr] = { week: w.week, year, cols, open: false };
    }
    sheet.maxCol = Math.max(sheet.maxCol || 20, next + 1);
    sheet.weekDays = wd;
    return wd;
  }

  // ---------------------------------------------------------------- display
  function orderColumns(sheet, cols) {
    const wd = sheet && sheet.weekDays;
    if (!wd) return cols;
    const name = App().getActiveSheet();
    const inGroup = new Set();
    for (const g of Object.values(wd.groups)) g.cols.forEach(c => inGroup.add(colToIndex(c)));
    const out = [];
    for (const c of cols) {
      if (inGroup.has(c)) continue;
      out.push(c);
      const g = wd.groups[indexToCol(c)];
      if (g && isOpen(name, indexToCol(c))) g.cols.forEach(dc => out.push(colToIndex(dc)));
    }
    return out;
  }

  function headerHtml(sheet, c) {
    const name = App().getActiveSheet();
    if (!isWeekSheet(name)) return '';
    const colStr = indexToCol(c);
    const g = sheet.weekDays && sheet.weekDays.groups[colStr];
    if (sheet.weekDays ? !g : !isWeekHeader(sheet, colStr)) return '';
    const open = !!g && isOpen(name, colStr);
    return `<button type="button" class="wk-toggle${open ? ' open' : ''}" data-wk="${colStr}" title="${open ? 'Ocultar días' : 'Capturar por día (sábado a viernes)'}">${open ? '−' : '+'}</button>`;
  }

  function toggle(colStr) {
    const app = App();
    const name = app.getActiveSheet();
    const sheet = app.getWorkbook().sheets[name];
    const first = !sheet.weekDays;
    if (first) app.pushUndo([name]);
    const wd = ensureGroups(name);
    const g = wd.groups[colStr];
    if (!g) return;
    const open = !isOpen(name, colStr);
    setOpen(name, colStr, open);
    // only the first time (the day columns are created) goes to the server; opening/closing is personal
    if (first) { app.markDirty(name); app.save(); }
    // keep the view where it was (re-rendering would otherwise jump)
    const vp = document.getElementById('spreadsheetViewport');
    const { scrollTop, scrollLeft } = vp;
    app.recalcAndRender();
    vp.scrollTop = scrollTop;
    vp.scrollLeft = scrollLeft;
    setTimeout(() => { vp.scrollTop = scrollTop; vp.scrollLeft = scrollLeft; }, 0);
    if (open) app.toast(`Semana ${g.week}: ${dayLabels(g.year, g.week)[0]} a ${dayLabels(g.year, g.week)[6]}`);
  }

  // ---------------------------------------------------------------- week total vs its days
  const numOf = v => (typeof v === 'number' ? v : Number(String(v == null ? '' : v).replace(/,/g, '')) || 0);
  // true when the row has days captured in that week and they don't add up to the week cell
  function weekMismatch(sheetName, r, c, sheet) {
    const wd = sheet && sheet.weekDays;
    if (!wd || r < FIRST_ROW || !isWeekSheet(sheetName)) return false;
    const colStr = indexToCol(c);
    const g = wd.groups[colStr];
    if (!g) return false;
    const week = sheet.cells[`${colStr}${r}`];
    if (week && week.f) return false; // old =SUM of its days: always adds up
    let any = false, sum = 0;
    for (const dc of g.cols) {
      const e = sheet.cells[`${dc}${r}`];
      if (!e || e.v === undefined || e.v === null || e.v === '') continue;
      any = true;
      sum += numOf(e.v);
    }
    return any && Math.abs(sum - numOf(week && week.v)) > 1e-9;
  }

  document.addEventListener('mousedown', (e) => {
    const btn = e.target.closest && e.target.closest('.wk-toggle');
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    toggle(btn.getAttribute('data-wk'));
  }, true);

  // ---------------------------------------------------------------- renaming a sheet to "P Cosecha YYWW"
  // updates its title (the YYWW cell at the top) and its week columns: WW, WW+1, WW+2...
  function weeksInYear(year) {
    const dec28 = new Date(Date.UTC(year, 11, 28));
    const thursday = new Date(dec28); thursday.setUTCDate(dec28.getUTCDate() + 3 - ((dec28.getUTCDay() + 6) % 7));
    const jan1 = new Date(Date.UTC(thursday.getUTCFullYear(), 0, 1));
    return Math.ceil(((thursday - jan1) / 86400000 + 1) / 7);
  }
  function onSheetRenamed(oldName, newName) {
    for (const id of Object.keys(openState)) {
      const cut = id.lastIndexOf('|');
      if (id.slice(0, cut) === String(oldName).trim()) { delete openState[id]; openState[openId(newName, id.slice(cut + 1))] = 1; }
    }
    try { localStorage.setItem(OPEN_KEY, JSON.stringify(openState)); } catch (e) { /* storage unavailable */ }
    const sheet = App().getWorkbook().sheets[newName];
    const yw = sheetYearWeek(newName);
    if (!sheet || !isWeekSheet(newName) || !yw) return false;
    const code = Number(String(yw.year).slice(2) + String(yw.week).padStart(2, '0'));
    const old = sheetYearWeek(oldName);
    const oldCode = old ? Number(String(old.year).slice(2) + String(old.week).padStart(2, '0')) : null;
    // a later week (e.g. a copy of 2646 renamed 2647): the numbers move left with their weeks
    const shifted = shiftForNewWeek(sheet, yw);
    // title: the YYWW number in the top rows (e.g. C2 = 2644)
    for (const [coord, cell] of Object.entries(sheet.cells)) {
      const m = /^([A-Z]+)([1-3])$/.exec(coord);
      if (!m || cell.f) continue;
      const v = Number(cell.v);
      if (Number.isInteger(v) && v >= 1000 && v <= 9999 && (v === oldCode || Math.floor(v / 100) === yw.year % 100)) cell.v = code;
    }
    // week columns: the first gets the sheet's week, the next ones follow (wrapping into next year)
    const cols = weekColumns(sheet);
    const max = weeksInYear(yw.year);
    let week = yw.week, year = yw.year;
    const plan = cols.map(w => {
      const out = { colStr: w.colStr, week, year };
      week++;
      if (week > max) { week = 1; year++; }
      return out;
    });
    for (const p of plan) {
      sheet.cells[`${p.colStr}${HEADER_ROW}`].v = p.week;
      const g = sheet.weekDays && sheet.weekDays.groups[p.colStr];
      if (g) {
        g.week = p.week;
        g.year = p.year;
        dayLabels(p.year, p.week).forEach((label, i) => {
          const cell = sheet.cells[`${g.cols[i]}${HEADER_ROW}`] || (sheet.cells[`${g.cols[i]}${HEADER_ROW}`] = {});
          cell.v = label;
        });
      }
    }
    App().toast(`Semanas de ${newName.trim()}: ${plan.map(p => p.week).join(', ')}`
      + (shifted ? ` · números recorridos ${shifted} semana(s)` : ''));
    return true;
  }

  // The sheet's first week column held week W (its header, before relabeling) and the new name starts at
  // a later week: every week's numbers (week cell + its 7 days) move k columns to the left, so each week
  // keeps its own numbers; the weeks that are new at the end stay empty. Returns k (0: nothing moved).
  function shiftForNewWeek(sheet, yw) {
    const cols = weekColumns(sheet);
    if (!cols.length) return 0;
    const groups = (sheet.weekDays && sheet.weekDays.groups) || {};
    const first = cols[0];
    let oldYear = groups[first.colStr] && groups[first.colStr].year;
    if (!oldYear) {
      // from the title (YYWW in the top rows), else the new name's year
      for (const [coord, cell] of Object.entries(sheet.cells)) {
        const v = Number(cell.v);
        if (/^[A-Z]+[1-3]$/.test(coord) && !cell.f && Number.isInteger(v) && v >= 1000 && v <= 9999) { oldYear = 2000 + Math.floor(v / 100); break; }
      }
    }
    if (!oldYear) oldYear = yw.year;
    const k = Math.round((weekSaturday(yw.year, yw.week) - weekSaturday(oldYear, first.week)) / (7 * 86400000));
    if (k < 1 || k >= cols.length) return 0;
    let last = FIRST_ROW;
    for (const coord of Object.keys(sheet.cells)) {
      const p = parseCoord(coord);
      if (p && p.row > last) last = p.row;
    }
    const take = coord => { const c = sheet.cells[coord]; return c && (c.v !== undefined || c.f) && !ownTotal(coord) ? { v: c.v, f: c.f } : null; };
    // a column total (=SUM(R5:R830) in column R) stays where it is
    const ownTotal = coord => {
      const c = sheet.cells[coord];
      if (!c || !c.f) return false;
      const col = coord.replace(/\d+$/, '');
      const refs = c.f.match(/\$?[A-Z]{1,3}\$?\d+/g) || [];
      return refs.length > 0 && refs.every(x => x.replace(/[$\d]/g, '') === col);
    };
    const put = (coord, val) => {
      if (ownTotal(coord)) return;
      const cell = sheet.cells[coord] || {};
      delete cell.v; delete cell.f; delete cell.e;
      if (val && val.f) cell.f = val.f;
      else if (val && val.v !== undefined) cell.v = val.v;
      if (Object.keys(cell).length) sheet.cells[coord] = cell; else delete sheet.cells[coord];
    };
    // what every week has now (rows of data only)
    const snap = cols.map(w => {
      const g = groups[w.colStr];
      const rows = {};
      for (let r = FIRST_ROW; r <= last; r++) rows[r] = { week: take(`${w.colStr}${r}`), days: g ? g.cols.map(dc => take(`${dc}${r}`)) : [] };
      return { g, rows };
    });
    cols.forEach((w, i) => {
      const src = snap[i + k];
      const g = groups[w.colStr];
      // formulas that moved (=SUM of the week's days, column totals) point at this week's column / days
      const map = new Map();
      if (src) map.set(cols[i + k].colStr, w.colStr);
      if (src && src.g && g) src.g.cols.forEach((dc, j) => map.set(dc, g.cols[j]));
      for (let r = FIRST_ROW; r <= last; r++) {
        const from = src ? src.rows[r] : null;
        let week = from && from.week;
        if (week && week.f) week = { f: week.f.replace(/(\$?)([A-Z]{1,3})(\$?\d+)/g, (m, a, col, rest) => (map.has(col) ? `${a}${map.get(col)}${rest}` : m)) };
        put(`${w.colStr}${r}`, week);
        if (g) g.cols.forEach((dc, j) => {
          let day = from && from.days[j];
          if (day && day.f) day = { f: day.f.replace(/(\$?)([A-Z]{1,3})(\$?\d+)/g, (m, a, col, rest) => (map.has(col) ? `${a}${map.get(col)}${rest}` : m)) };
          put(`${dc}${r}`, day);
        });
      }
    });
    return k;
  }

  window.MauriWeeks = { orderColumns, headerHtml, toggle, weekMismatch, dayLabels, weekSaturday, onSheetRenamed };
})();
