// ============================================================================
// DÍAS POR SEMANA (Plan de cosecha) - expandable Sat..Fri capture columns
//   * Week columns = row 4 headers holding a week number (1-53) on "P Cosecha" sheets
//   * sheet.weekDays = { start, lastRow, groups: { <weekCol>: { week, year, cols:[7 cols], open } } }
//     The 7 day columns live after the used area (they never shift existing cells) and are
//     displayed right after their week while the group is open.
//   * Once a row has a day captured, its week cell becomes =SUM(<its 7 days>)
// ============================================================================

(function () {
  const { indexToCol, colToIndex, parseCoord } = window.ExcelUtils;
  const App = () => window.MauriApp;
  const HEADER_ROW = 4;
  const FIRST_ROW = 5;
  const DAY_NAMES = ['Sáb', 'Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie'];
  const MONTHS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const DAY_COL_WIDTH = 12.5;

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
    const inGroup = new Set();
    for (const g of Object.values(wd.groups)) g.cols.forEach(c => inGroup.add(colToIndex(c)));
    const out = [];
    for (const c of cols) {
      if (inGroup.has(c)) continue;
      out.push(c);
      const g = wd.groups[indexToCol(c)];
      if (g && g.open) g.cols.forEach(dc => out.push(colToIndex(dc)));
    }
    return out;
  }

  function headerHtml(sheet, c) {
    const name = App().getActiveSheet();
    if (!isWeekSheet(name)) return '';
    const colStr = indexToCol(c);
    const g = sheet.weekDays && sheet.weekDays.groups[colStr];
    if (sheet.weekDays ? !g : !isWeekHeader(sheet, colStr)) return '';
    const open = !!(g && g.open);
    return `<button type="button" class="wk-toggle${open ? ' open' : ''}" data-wk="${colStr}" title="${open ? 'Ocultar días' : 'Capturar por día (sábado a viernes)'}">${open ? '−' : '+'}</button>`;
  }

  function toggle(colStr) {
    const app = App();
    const name = app.getActiveSheet();
    const sheet = app.getWorkbook().sheets[name];
    if (!sheet.weekDays) app.pushUndo([name]);
    const wd = ensureGroups(name);
    const g = wd.groups[colStr];
    if (!g) return;
    g.open = !g.open;
    app.markDirty(name);
    app.save();
    // keep the view where it was (re-rendering would otherwise jump)
    const vp = document.getElementById('spreadsheetViewport');
    const { scrollTop, scrollLeft } = vp;
    app.recalcAndRender();
    vp.scrollTop = scrollTop;
    vp.scrollLeft = scrollLeft;
    setTimeout(() => { vp.scrollTop = scrollTop; vp.scrollLeft = scrollLeft; }, 0);
    if (g.open) app.toast(`Semana ${g.week}: ${dayLabels(g.year, g.week)[0]} a ${dayLabels(g.year, g.week)[6]}`);
  }

  // ---------------------------------------------------------------- week = sum of its days
  function syncWeekTotals(workbook) {
    for (const [name, sheet] of Object.entries(workbook.sheets || {})) {
      const wd = sheet.weekDays;
      if (!wd) continue;
      let last = wd.lastRow;
      for (const coord of Object.keys(sheet.cells)) {
        const r = Number(coord.replace(/^[A-Z]+/, ''));
        if (r > last) last = r;
      }
      for (const [weekCol, g] of Object.entries(wd.groups)) {
        for (let r = FIRST_ROW; r <= last; r++) {
          const has = g.cols.some(c => {
            const e = sheet.cells[`${c}${r}`];
            return e && ((e.v !== undefined && e.v !== null && e.v !== '') || e.f);
          });
          if (!has) continue;
          const coord = `${weekCol}${r}`;
          const formula = `=SUM(${g.cols[0]}${r}:${g.cols[6]}${r})`;
          const cell = sheet.cells[coord] || (sheet.cells[coord] = {});
          if (cell.f === formula) continue;
          cell.f = formula;
          delete cell.v;
          if (App() && App().markDirty) App().markDirty(name);
        }
      }
    }
  }

  const proto = window.FormulaEngine && window.FormulaEngine.prototype;
  if (proto && !proto.__weeksPatched) {
    const recalc = proto.recalculateAll;
    proto.recalculateAll = function () {
      try { syncWeekTotals(this.workbook); } catch (e) { console.error(e); }
      return recalc.apply(this, arguments);
    };
    proto.__weeksPatched = true;
  }

  document.addEventListener('mousedown', (e) => {
    const btn = e.target.closest && e.target.closest('.wk-toggle');
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    toggle(btn.getAttribute('data-wk'));
  }, true);

  window.MauriWeeks = { orderColumns, headerHtml, toggle, dayLabels, weekSaturday };
})();
