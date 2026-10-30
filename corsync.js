// ============================================================================
// REUNION: Cor in real time from Plan de cosecha (cosechaCor.js does the math)
//   * The page keeps the latest feed from the server (/api/cor-feed + /api/events)
//   * Every recalculation applies it first, so the screen always shows Plan de cosecha's numbers:
//   * A number typed in a covered Cor is sent to Plan de cosecha (row of finca Isabela) and stays;
//     anything else typed there (text, formulas) goes back to the planned value
// ============================================================================

(function () {
  const App = () => window.MauriApp;
  // Keep the real-time feeds alive for the whole single-page session. The
  // active workbook decides which feed is applied, so switching apartados
  // does not require a full page reload.
  // (the page's single live connection, app.js: one per tab, browsers allow only a few per server)
  window.MauriEvents.on('cosecha', e => {
    const app = App();
    if (!app || app.getWorkbookKey?.() !== 'cosecha') return;
    const update = JSON.parse(e.data);
    const sheet = app.getWorkbook() && app.getWorkbook().sheets[update.sheet];
    if (!sheet) return;
    Object.assign(sheet.cells, update.cells);
    app.recalcCellsAndRender(Object.keys(update.cells).map(coord => [update.sheet, coord]));
  });
  if (!window.CosechaCor) return; // The real-time Cor bridge is optional.
  let feed = null;
  let quiet = false; // applying a server update (not a user edit)

  const Engine = window.FormulaEngine;
  if (Engine && !Engine.__corPatched) {
    Engine.beforeRecalc.push((workbook, changed) => {
      if (!feed || !workbook || !workbook.sheets || !workbook.sheets.Reunion) return;
      const coords = changed ? [] : undefined;
      const n = window.CosechaCor.applyFeed(workbook, feed, coords);
      if (coords) for (const c of coords) changed.push(`Reunion!${c}`);
      if (n && App()) {
        App().markDirty('Reunion');
        if (!quiet) App().toast(`Ese Cor solo acepta números (viene de Plan de cosecha ${feed.sheet.trim()})`, true);
      }
    });
    Engine.__corPatched = true;
  }

  function applyNow() {
    const app = App();
    if (!app || !app.getWorkbook() || !app.getEngine()) return false;
    quiet = true;
    try { app.recalcCellsAndRender(); } finally { quiet = false; } // the feed hook adds the Cor it changes
    return true;
  }
  function setFeed(next) {
    feed = next;
    history = (next && next.history && next.history.entries) || {};
    refreshPanel();
    if (!applyNow()) {
      const wait = setInterval(() => { if (applyNow()) clearInterval(wait); }, 300);
    }
  }

  // ---------------------------------------------------------------- change marks + small panel
  //   a Cor whose number changed in Plan de cosecha gets a corner mark; clicking it opens a
  //   panel with the current / previous quantity and the last 5 changes. "✓ Visto" removes the mark.
  let history = {};
  const C = window.CosechaCor;
  const layoutCache = new WeakMap();
  function corColumns(sheet) {
    const d1 = sheet.cells.D1 && sheet.cells.D1.v;
    let info = layoutCache.get(sheet);
    if (!info || info.d1 !== d1) { info = { d1, cols: C.reunionCorColumns(sheet) }; layoutCache.set(sheet, info); }
    return info.cols;
  }
  function keyAt(sheetName, sheet, row, colStr) {
    if (sheetName !== 'Reunion' || row < 3) return null;
    const serial = corColumns(sheet).get(colStr);
    if (!serial) return null;
    const product = C.norm((sheet.cells[`A${row}`] || {}).v);
    return product ? C.keyOf(product, serial) : null;
  }
  function cellExtra(sheetName, r, c, sheet) {
    const colStr = window.ExcelUtils.indexToCol(c);
    const key = keyAt(sheetName, sheet, r, colStr);
    const e = key && history[key];
    return e && e.unseen ? '<span class="cor-mark" title="Este Cor cambió en Plan de cosecha"></span>' : '';
  }

  const fmtNum = n => Number(n || 0).toLocaleString('en-US');
  const DAYS = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
  const MONTHS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  function fmtSerial(serial) {
    const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return `${DAYS[d.getUTCDay()]} ${d.getUTCDate()}-${MONTHS[d.getUTCMonth()]}`;
  }
  function fmtWhen(iso) {
    if (!iso) return 'antes';
    const d = new Date(iso);
    const today = new Date();
    const time = d.toLocaleTimeString('es-MX', { hour: 'numeric', minute: '2-digit' });
    if (d.toDateString() === today.toDateString()) return `hoy ${time}`;
    return `${DAYS[d.getDay()].toLowerCase()} ${d.getDate()}-${MONTHS[d.getMonth()]} ${time}`;
  }
  const esc = t => String(t).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  let panel = null;
  let panelKey = null;
  let panelLabel = '';
  function closePanel() { if (panel) panel.remove(); panel = null; panelKey = null; }
  function seen(body) {
    fetch('/api/cor-history/seen', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).catch(() => {});
  }
  function renderPanel() {
    const e = history[panelKey];
    if (!e || !e.unseen) { closePanel(); return; }
    const pts = e.points;
    const cur = pts[pts.length - 1].v;
    const prev = pts.length > 1 ? pts[pts.length - 2].v : cur;
    const diff = cur - prev;
    const arrow = diff > 0 ? `<b class="cm-up">▲ +${fmtNum(diff)}</b>` : diff < 0 ? `<b class="cm-down">▼ ${fmtNum(diff)}</b>` : '';
    const last = pts[pts.length - 1];
    const chain = pts.map((p, i) => `<span class="cm-pt${i === pts.length - 1 ? ' cur' : ''}">${fmtNum(p.v)} <small>(${fmtWhen(p.at)})</small></span>`).join('<span class="cm-arrow">→</span>');
    const unseenCount = Object.values(history).filter(x => x.unseen).length;
    if (!panel) {
      panel = document.createElement('div');
      panel.className = 'cor-panel';
      document.body.appendChild(panel);
      const bar = document.querySelector('.excel-formulabar');
      panel.style.top = `${Math.round(bar ? bar.getBoundingClientRect().bottom + 4 : 70)}px`;
      panel.addEventListener('click', ev => {
        const b = ev.target.closest('[data-cm]');
        if (!b) return;
        if (b.dataset.cm === 'close') closePanel();
        else if (b.dataset.cm === 'seen') { seen({ keys: [panelKey] }); history[panelKey].unseen = false; closePanel(); rerender(); }
        else if (b.dataset.cm === 'all') { seen({ all: true }); Object.values(history).forEach(x => { x.unseen = false; }); closePanel(); rerender(); }
      });
    }
    panel.innerHTML = `
      <div class="cm-row">
        <b class="cm-title">${esc(panelLabel)}</b>
        <span>Actual: <b>${fmtNum(cur)}</b></span>
        <span>Pasada: <b>${fmtNum(prev)}</b></span>
        ${arrow}
        <span class="cm-muted">${fmtWhen(last.at)}${last.sheet ? ' · ' + esc(String(last.sheet).trim()) : ''}</span>
        <span class="cm-spacer"></span>
        <button type="button" data-cm="seen" title="Quitar la marca de esta celda">✓ Visto</button>
        ${unseenCount > 1 ? `<button type="button" data-cm="all" title="Quitar todas las marcas">Marcar todo como visto (${unseenCount})</button>` : ''}
        <button type="button" data-cm="close" class="cm-x" title="Cerrar">✕</button>
      </div>
      <div class="cm-row cm-hist">Historial: ${chain}</div>`;
  }
  function refreshPanel() { if (panelKey) renderPanel(); }
  function rerender() { const app = App(); if (app && app.getEngine()) app.render(); } // marks only
  function onSelect(sheetName, coord) {
    const app = App();
    const sheet = app && app.getWorkbook() && app.getWorkbook().sheets[sheetName];
    const m = /^([A-Z]+)(\d+)$/.exec(coord || '');
    const key = sheet && m ? keyAt(sheetName, sheet, Number(m[2]), m[1]) : null;
    if (!key || !history[key] || !history[key].unseen) { closePanel(); return; }
    panelKey = key;
    const serial = corColumns(sheet).get(m[1]);
    panelLabel = `${String((sheet.cells[`A${m[2]}`] || {}).v || '').trim()} · Cor ${fmtSerial(serial)}`;
    renderPanel();
  }
  // ---------------------------------------------------------------- typing a Cor -> Plan de cosecha
  let pending = [];
  function onUserEdit(sheetName, coord) {
    if (quiet || !feed || !feed.sheet || sheetName !== 'Reunion') return;
    const sheet = App().getWorkbook().sheets.Reunion;
    const m = /^([A-Z]+)(\d+)$/.exec(coord);
    const key = m && keyAt(sheetName, sheet, Number(m[2]), m[1]);
    if (!key) return;
    const product = C.norm((sheet.cells[`A${m[2]}`] || {}).v);
    const serial = corColumns(sheet).get(m[1]);
    if (!(product in feed.values) || !feed.dates.includes(serial)) return;
    const cell = sheet.cells[coord] || {};
    const empty = cell.v === undefined || cell.v === null || cell.v === '';
    if (cell.f || (!empty && typeof cell.v !== 'number')) return; // not a number: the planned value comes back
    const value = empty ? 0 : cell.v;
    feed.values[product][serial] = value; // keep it on screen while the server writes it
    pending.push({ product, serial, value });
    if (pending.length === 1) setTimeout(flush, 0);
  }
  function flush() {
    const writes = pending;
    pending = [];
    fetch('/api/cor-write', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ writes }) })
      .then(r => r.json())
      .then(r => {
        if (r.errors && r.errors.length) {
          App().toast(r.errors[0] + (r.errors.length > 1 ? ` (+${r.errors.length - 1})` : ''), true);
          return fetch('/api/cor-feed').then(x => x.json()).then(setFeed);
        }
      })
      .catch(() => { App().toast('No se pudo guardar el Cor en Plan de cosecha', true); });
  }

  window.MauriCorMarks = { cellExtra, onSelect, onUserEdit };

  fetch('/api/cor-feed').then(r => (r.ok ? r.json() : null)).then(f => { if (f) setFeed(f); }).catch(() => {});
  window.MauriEvents.on('cor', e => { try { setFeed(JSON.parse(e.data)); } catch (err) { console.error(err); } });
})();
