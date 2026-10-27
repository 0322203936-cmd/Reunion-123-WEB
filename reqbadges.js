// ============================================================================
// REQ EN PLAN DE COSECHA - small box inside each day / week cell with the Req of Reunion
//   * Sheets with a "Flor 123" column (row 4): each row's Flor 123 is a Reunion product
//   * Req comes from Reunion itself (its own engine), so it is exactly what Reunion shows:
//     sheet "Reunion", row = product, columns titled "Req" (row 2), date = the "Fecha" column before it
//   * Day columns: Sabado..Viernes (the sheet's week) and the days opened with + (weeks.js);
//     week columns (44, 45...): the sum of that week's Sat..Fri
// ============================================================================

(function () {
  const params = new URLSearchParams(location.search);
  if (params.get('wb') !== 'cosecha') return;
  const { indexToCol, parseCoord } = window.ExcelUtils;
  const App = () => window.MauriApp;
  const HEADER_ROW = 4;
  const FIRST_ROW = 5;
  const DAY_HEADERS = ['SABADO', 'DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES'];
  const EPOCH = Date.UTC(1899, 11, 30);

  const clean = v => String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
  const norm = v => clean(v).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const serialOf = d => Math.round((d.getTime() - EPOCH) / 86400000);
  const fmt = n => (Math.round(n * 100) / 100).toLocaleString('en-US');

  let req = null; // Map product -> Map(serial -> req)
  let reqDates = new Set();

  async function loadReq() {
    const res = await fetch('/api/workbook?wb=reunion');
    if (!res.ok) throw new Error('No se pudo leer Reunion');
    const book = await res.json();
    const sheet = book.sheets && book.sheets.Reunion;
    if (!sheet) return;
    const engine = new window.FormulaEngine(book);
    engine.recalculateAll();
    const cells = sheet.cells;
    // Req columns: header "Req" in row 2; their date is the "Fecha" column just before
    const reqCols = [];
    for (let c = 2; c <= (sheet.maxCol || 80); c++) {
      if (norm((cells[`${indexToCol(c)}2`] || {}).v) === 'REQ') reqCols.push(c);
    }
    const map = new Map();
    let last = 3;
    for (const coord of Object.keys(cells)) {
      const m = /^A(\d+)$/.exec(coord);
      if (m && Number(m[1]) > last) last = Number(m[1]);
    }
    for (const c of reqCols) {
      const date = engine.getCellValue('Reunion', `${indexToCol(c - 1)}3`);
      if (typeof date !== 'number') continue;
      const serial = Math.floor(date);
      reqDates.add(serial);
      for (let r = 3; r <= last; r++) {
        const name = norm((cells[`A${r}`] || {}).v);
        if (!name) continue;
        const v = engine.getCellValue('Reunion', `${indexToCol(c)}${r}`);
        if (typeof v !== 'number') continue;
        if (!map.has(name)) map.set(name, new Map());
        const days = map.get(name);
        days.set(serial, (days.get(serial) || 0) + v);
      }
    }
    req = map;
  }

  // per sheet: Flor 123 column, day columns (col -> serial) and week columns (col -> [serials])
  const layoutCache = new WeakMap();
  function layoutOf(sheetName, sheet) {
    let info = layoutCache.get(sheet);
    const sig = sheetName + '|' + JSON.stringify(Object.values((sheet.weekDays && sheet.weekDays.groups) || {}).map(g => g.week));
    if (info && info.sig === sig && info.weekDaysRef === sheet.weekDays && info.maxCol === sheet.maxCol) return info;
    info = { flor: null, days: new Map(), weeks: new Map(), weekDaysRef: sheet.weekDays, maxCol: sheet.maxCol, sig };
    const head = c => sheet.cells[`${indexToCol(c)}${HEADER_ROW}`];
    const m = /(\d{2})(\d{2})\D*$/.exec(clean(sheetName));
    const year = m ? 2000 + Number(m[1]) : new Date().getFullYear();
    const sheetWeek = m ? Number(m[2]) : null;
    const yearOf = w => (sheetWeek && w < sheetWeek - 26 ? year + 1 : year);
    const saturday = w => window.MauriWeeks.weekSaturday(yearOf(w), w);
    const weekSerials = w => { const s = serialOf(saturday(w)); return [0, 1, 2, 3, 4, 5, 6].map(i => s + i); };
    let firstWeek = null;
    for (let c = 1; c <= (sheet.maxCol || 60); c++) {
      const cell = head(c);
      if (!cell) continue;
      const h = norm(cell.v);
      if (h === 'FLOR 123') info.flor = indexToCol(c);
      if (!cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) {
        if (firstWeek === null) firstWeek = cell.v;
        info.weeks.set(c, weekSerials(cell.v));
      }
    }
    // Sabado..Viernes columns belong to the sheet's own week (the first week column)
    if (firstWeek !== null) {
      const serials = weekSerials(firstWeek);
      for (let c = 1; c <= (sheet.maxCol || 60); c++) {
        const i = DAY_HEADERS.indexOf(norm((head(c) || {}).v));
        if (i >= 0) info.days.set(c, serials[i]);
      }
    }
    // days opened with + (weeks.js)
    for (const g of Object.values((sheet.weekDays && sheet.weekDays.groups) || {})) {
      const serials = weekSerials(g.week);
      g.cols.forEach((col, i) => info.days.set(parseCoord(`${col}1`).col, serials[i]));
    }
    layoutCache.set(sheet, info);
    return info;
  }

  function cellExtra(sheetName, r, c, sheet) {
    if (!req || r < FIRST_ROW || !/^\s*p\s*cosecha/i.test(sheetName)) return '';
    const info = layoutOf(sheetName, sheet);
    if (!info.flor) return '';
    let serials = null;
    if (info.days.has(c)) serials = [info.days.get(c)];
    else if (info.weeks.has(c)) serials = info.weeks.get(c);
    if (!serials) return '';
    const product = norm((sheet.cells[`${info.flor}${r}`] || {}).v);
    if (!product) return '';
    const days = req.get(product);
    const covered = serials.filter(s => reqDates.has(s));
    if (!covered.length) return ''; // Reunion does not reach these dates
    const total = days ? covered.reduce((s, x) => s + (days.get(x) || 0), 0) : 0;
    if (!total) return '';
    const isWeek = !info.days.has(c);
    const partial = covered.length < serials.length;
    const title = `Req de Reunion ${isWeek ? 'de la semana' : 'del día'}: ${fmt(total)}${partial ? ` (solo ${covered.length} de ${serials.length} días tienen datos en Reunion)` : ''}`;
    // yellow = week total, blue = a single day
    return `<span class="req-badge${isWeek ? ' week' : ''}${partial ? ' partial' : ''}" title="${title}">${fmt(total)}</span>`;
  }

  // the Req box takes room on the left: columns with boxes grow (on screen only) so the number still fits
  let measureCtx = null;
  function textWidth(text, font) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = font;
    return measureCtx.measureText(text).width;
  }
  function minColPx(sheetName, sheet, c) {
    if (!req || !/^\s*p\s*cosecha/i.test(sheetName)) return 0;
    const info = layoutOf(sheetName, sheet);
    if (!info.flor || (!info.days.has(c) && !info.weeks.has(c))) return 0;
    const col = indexToCol(c);
    let need = 0;
    for (let r = FIRST_ROW; r <= (sheet.maxRow || 300); r++) {
      const badge = cellExtra(sheetName, r, c, sheet);
      if (!badge) continue;
      const reqText = 'R ' + badge.replace(/<[^>]+>/g, '');
      const value = App().cellDisplayText(sheetName, `${col}${r}`) || '';
      const w = textWidth(reqText, '700 9.5px Calibri, sans-serif') + 12 + textWidth(value, '11pt Calibri, sans-serif') + 14;
      if (w > need) need = w;
    }
    return Math.ceil(need);
  }

  window.MauriReq = { cellExtra, minColPx, reload: () => loadReq().then(() => App().recalcAndRender()) };

  loadReq().then(() => {
    const wait = setInterval(() => {
      const app = App();
      if (!app || !app.getWorkbook() || !app.getEngine()) return;
      clearInterval(wait);
      app.recalcAndRender();
    }, 300);
  }).catch(err => console.error('Req de Reunion:', err));
})();
