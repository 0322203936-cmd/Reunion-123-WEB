// ============================================================================
// PLAN DE COSECHA -> COR DE REUNION (shared by server.js and the Reunion page)
//   * Source: the LAST "P Cosecha" tab that has a "Flor 123" column (row 4)
//   * Each row's Flor 123 is a Reunion product; its days are the Sat..Fri columns opened with +
//     (sheet.weekDays, weeks.js). Rows of the same product (fincas) add up.
//   * Every date of the weeks in that tab is "covered": Reunion's Cor for that product and date
//     becomes exactly the captured total (0 when nothing was captured). Only the number changes,
//     never the format. Titles (Cor that adds other cells) and products not in Flor 123 are left alone.
// ============================================================================

(function (root) {
  const HEADER_ROW = 4;
  const FIRST_ROW = 5;
  const EPOCH = Date.UTC(1899, 11, 30);

  const clean = v => String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
  const norm = v => clean(v).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const colToIndex = col => col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const indexToCol = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; };
  const num = v => (typeof v === 'number' ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v)) ? Number(v) : 0));

  // Saturday that opens week `week` of `year` (ISO Monday - 2), as an Excel serial
  function saturdaySerial(year, week) {
    const jan4 = Date.UTC(year, 0, 4);
    const dow = (new Date(jan4).getUTCDay() + 6) % 7;
    const monday = jan4 - dow * 86400000 + (week - 1) * 7 * 86400000;
    return Math.round((monday - 2 * 86400000 - EPOCH) / 86400000);
  }

  function lastCosechaSheet(book) {
    const order = (book && book.sheetOrder) || [];
    for (let i = order.length - 1; i >= 0; i--) {
      const name = order[i];
      const sheet = book.sheets[name];
      if (!sheet || sheet.hidden || !/^\s*p\s*cosecha/i.test(name)) continue;
      const cells = sheet.cells || {};
      for (const [coord, cell] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(coord);
        if (m && Number(m[2]) === HEADER_ROW && norm(cell.v) === 'FLOR 123') return { name, sheet, florCol: m[1] };
      }
    }
    return null;
  }

  // dates covered by the weeks of the tab + the day column of each date: { dates: Set, dayCols: [{ col, serial, weekCol }] }
  function dayColumns(found) {
    const { name, sheet } = found;
    const cells = sheet.cells || {};
    const m = /(\d{2})(\d{2})\D*$/.exec(clean(name));
    const year = m ? 2000 + Number(m[1]) : new Date().getFullYear();
    const sheetWeek = m ? Number(m[2]) : 1;
    // week columns in order (row 4: a plain week number)
    const weekCols = [];
    for (const [coord, cell] of Object.entries(cells)) {
      const mm = /^([A-Z]+)(\d+)$/.exec(coord);
      if (mm && Number(mm[2]) === HEADER_ROW && !cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) weekCols.push({ col: mm[1], week: cell.v });
    }
    weekCols.sort((a, b) => colToIndex(a.col) - colToIndex(b.col));
    const dates = new Set();
    const dayCols = [];
    const groups = (sheet.weekDays && sheet.weekDays.groups) || {};
    for (const w of weekCols) {
      const y = w.week < sheetWeek - 26 ? year + 1 : year;
      const sat = saturdaySerial(y, w.week);
      for (let i = 0; i < 7; i++) dates.add(sat + i);
      const g = groups[w.col];
      if (g && Array.isArray(g.cols)) g.cols.forEach((col, i) => dayCols.push({ col, serial: sat + i, weekCol: w.col }));
    }
    return { dates, dayCols };
  }
  function lastRowOf(cells, col) {
    let last = FIRST_ROW;
    for (const coord of Object.keys(cells)) {
      const mm = /^([A-Z]+)(\d+)$/.exec(coord);
      if (mm && mm[1] === col && Number(mm[2]) > last) last = Number(mm[2]);
    }
    return last;
  }

  // { sheet, dates: [serial], values: { PRODUCT: { serial: total } }, sig }
  function computeFeed(cosechaBook) {
    const found = lastCosechaSheet(cosechaBook);
    if (!found) return { sheet: null, dates: [], values: {}, sig: 'none' };
    const { name, sheet, florCol } = found;
    const cells = sheet.cells || {};
    const { dates, dayCols } = dayColumns(found);
    const last = lastRowOf(cells, florCol);
    const values = {};
    for (let r = FIRST_ROW; r <= last; r++) {
      const product = norm((cells[`${florCol}${r}`] || {}).v);
      if (!product) continue;
      const days = values[product] || (values[product] = {});
      for (const d of dayCols) {
        const v = num((cells[`${d.col}${r}`] || {}).v);
        if (v) days[d.serial] = (days[d.serial] || 0) + v;
      }
    }
    const sortedDates = [...dates].sort((a, b) => a - b);
    const sig = JSON.stringify([name, sortedDates[0], sortedDates[sortedDates.length - 1], values]);
    return { sheet: name, dates: sortedDates, values, sig };
  }

  // Cor typed in Reunion -> Plan de cosecha: the day cell of the only finca with a number that day,
  //   else of finca WRITE_FINCA
  //   (first row of the product when it has none) gets what makes the day's total equal `value`.
  //   Returns { cells: { coord: cell } } or { error }
  const WRITE_FINCA = 'ISABELA';
  function writeCor(cosechaBook, product, serial, value) {
    const found = lastCosechaSheet(cosechaBook);
    if (!found) return { error: 'No hay pestaña P Cosecha con Flor 123' };
    const { sheet, florCol } = found;
    const cells = sheet.cells || (sheet.cells = {});
    const day = dayColumns(found).dayCols.find(d => d.serial === serial);
    if (!day) return { error: `Esa fecha no está en ${clean(found.name)}` };
    const rows = [];
    for (let r = FIRST_ROW; r <= lastRowOf(cells, florCol); r++) if (norm((cells[`${florCol}${r}`] || {}).v) === product) rows.push(r);
    if (!rows.length) return { error: `${product} no está en Flor 123 de ${clean(found.name)}` };
    let fincaCol = null;
    for (const [coord, cell] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(coord);
      if (m && Number(m[2]) === HEADER_ROW && norm(cell.v) === 'FINCA') fincaCol = m[1];
    }
    // only one finca has a number that day: the change goes to it; otherwise to Isabela
    const withValue = rows.filter(r => num((cells[`${day.col}${r}`] || {}).v));
    const target = withValue.length === 1 ? withValue[0]
      : rows.find(r => fincaCol && norm((cells[`${fincaCol}${r}`] || {}).v) === WRITE_FINCA) || rows[0];
    const others = rows.filter(r => r !== target).reduce((t, r) => t + num((cells[`${day.col}${r}`] || {}).v), 0);
    const own = value - others;
    if (own < 0) return { error: `${product}: las otras fincas ya suman ${others} ese día` };
    const out = {};
    const coord = `${day.col}${target}`;
    const cell = { ...(cells[coord] || {}) };
    delete cell.f;
    if (own) cell.v = own; else delete cell.v;
    cells[coord] = out[coord] = cell;
    // the week is the sum of its days (same formula weeks.js writes); no days left: back to empty
    const g = sheet.weekDays.groups[day.weekCol];
    const weekCoord = `${day.weekCol}${target}`;
    const week = { ...(cells[weekCoord] || {}) };
    const total = g.cols.reduce((t, c) => t + num((cells[`${c}${target}`] || {}).v), 0);
    if (g.cols.some(c => num((cells[`${c}${target}`] || {}).v))) {
      week.f = `=SUM(${g.cols[0]}${target}:${g.cols[6]}${target})`;
      week.v = total;
    } else { delete week.f; delete week.v; }
    cells[weekCoord] = out[weekCoord] = week;
    return { sheet: found.name, cells: out };
  }

  // writes the feed into sheet "Reunion" of the Reunion book; returns the number of cells changed
  function applyFeed(reunionBook, feed) {
    const sheet = reunionBook && reunionBook.sheets && reunionBook.sheets.Reunion;
    if (!sheet || !feed || !feed.sheet) return 0;
    const cells = sheet.cells;
    const covered = new Set(feed.dates);
    // day blocks: "Cor" columns of row 2; block k is the date in D1 + k
    const corCols = [];
    for (let c = 1; c <= (sheet.maxCol || 80); c++) {
      if (norm((cells[`${indexToCol(c)}2`] || {}).v) === 'COR') corCols.push(indexToCol(c));
    }
    if (!corCols.length) return 0;
    const firstDate = cells[`${indexToCol(colToIndex(corCols[0]) - 1)}1`];
    const base = firstDate ? num(firstDate.v) : 0;
    if (!base) return 0;
    let last = 3;
    for (const coord of Object.keys(cells)) {
      const mm = /^A(\d+)$/.exec(coord);
      if (mm && Number(mm[1]) > last) last = Number(mm[1]);
    }
    let changed = 0;
    for (let r = 3; r <= last; r++) {
      const product = norm((cells[`A${r}`] || {}).v);
      if (!product || !(product in feed.values)) continue;
      const days = feed.values[product];
      corCols.forEach((col, k) => {
        const serial = Math.floor(base) + k;
        if (!covered.has(serial)) return;
        const coord = `${col}${r}`;
        const cell = cells[coord];
        if (cell && cell.f && /\$?[A-Z]{1,3}\$?\d+/.test(cell.f)) return; // title: adds other cells
        const target = days[serial] || 0;
        if (cell && !cell.f && cell.v === target) return;
        if (cell) { delete cell.f; cell.v = target; } else cells[coord] = { v: target };
        changed++;
      });
    }
    return changed;
  }

  // Reunion layout: Cor column -> date serial (block k is the date in D1 + k)
  function reunionCorColumns(sheet) {
    const cells = (sheet && sheet.cells) || {};
    const out = new Map();
    const corCols = [];
    for (let c = 1; c <= ((sheet && sheet.maxCol) || 80); c++) {
      if (norm((cells[`${indexToCol(c)}2`] || {}).v) === 'COR') corCols.push(indexToCol(c));
    }
    if (!corCols.length) return out;
    const firstDate = cells[`${indexToCol(colToIndex(corCols[0]) - 1)}1`];
    const base = firstDate ? Math.floor(num(firstDate.v)) : 0;
    if (base) corCols.forEach((col, k) => out.set(col, base + k));
    return out;
  }

  // ---- change history (last changes per product + date; a new source tab is a fresh start)
  const MAX_POINTS = 6; // 6 values = the last 5 changes
  const keyOf = (product, serial) => `${product}|${serial}`;
  function feedValues(feed) {
    const cur = {};
    if (!feed || !feed.sheet) return cur;
    for (const [product, days] of Object.entries(feed.values)) {
      for (const serial of feed.dates) cur[keyOf(product, serial)] = days[serial] || 0;
    }
    return cur;
  }
  // updates `hist` ({ source, baseline, setAt, entries }) in place; returns true when something changed.
  //   * a number has a 2 minute grace period after it is typed: changes inside it are the same capture
  //   * capturing where there was nothing (0) is not a modification; changing a settled number is
  const GRACE_MS = 2 * 60 * 1000;
  function trackChanges(hist, feed, now) {
    const cur = feedValues(feed);
    if (!feed || !feed.sheet) return false;
    if (hist.source !== feed.sheet || !hist.baseline) {
      hist.source = feed.sheet;
      hist.baseline = cur;
      hist.setAt = {};
      hist.since = now;
      hist.entries = hist.entries || {};
      return true;
    }
    hist.setAt = hist.setAt || {};
    const nowMs = Date.parse(now);
    let changed = false;
    for (const [key, v] of Object.entries(cur)) {
      const old = key in hist.baseline ? hist.baseline[key] : 0;
      if (old === v) continue;
      const setAt = hist.setAt[key];
      const fresh = setAt && nowMs - Date.parse(setAt) < GRACE_MS;
      const e = hist.entries[key];
      const last = e && e.points[e.points.length - 1];
      if (fresh) {
        // still capturing: if the number being corrected was a modification, it just gets the new value
        if (last && last.at === setAt) {
          const before = e.points[e.points.length - 2];
          if (before && before.v === v) {
            e.points.pop(); // back to the previous number: no modification after all
            if (e.points.length < 2) delete hist.entries[key];
          } else {
            last.v = v;
            last.at = now;
          }
        }
      } else if (old !== 0) {
        const entry = e || (hist.entries[key] = { points: [{ v: old, at: null }] });
        entry.points.push({ v, at: now, sheet: feed.sheet });
        if (entry.points.length > MAX_POINTS) entry.points = entry.points.slice(-MAX_POINTS);
        entry.unseen = true;
      }
      hist.setAt[key] = now;
      hist.baseline[key] = v;
      changed = true;
    }
    return changed;
  }

  const api = { computeFeed, writeCor, applyFeed, lastCosechaSheet, reunionCorColumns, trackChanges, keyOf, norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CosechaCor = api;
})(typeof window !== 'undefined' ? window : globalThis);
