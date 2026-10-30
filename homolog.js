// ============================================================================
// HOMOLOGACIONES (Plan de cosecha) - Reunion flower <-> Plan de cosecha combinations
//   * Left: Reunion flowers (rows of sheet "Reunion" whose Cor column has no formula = not a title)
//   * Right: every Flor - Variedad - Color - Grado found in the "P Cosecha" sheets
//   * Manual only: a flower can take many combinations, a combination can go to many flowers
//   * Stored once for all sheets: /api/homologations?wb=cosecha -> { version, items: { flor: [comboKey] } }
// ============================================================================

(function () {
  const App = () => window.MauriApp;
  const params = new URLSearchParams(location.search);
  const WB_KEY = params.get('wb') === 'cosecha' ? 'cosecha' : 'reunion';
  const HEADER_ROW = 4;
  const FIRST_ROW = 5;

  const btn = document.getElementById('btnHomologaciones');
  if (btn) btn.hidden = WB_KEY !== 'cosecha';
  if (WB_KEY !== 'cosecha' || !btn) return;

  const clean = v => String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
  const esc = s => clean(s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const norm = s => clean(s).toUpperCase();
  const comboKey = c => [c.flor, c.variedad, c.color, c.grado].map(norm).join('|');
  const comboLabel = key => {
    const [flor, variedad, color, grado] = key.split('|');
    return `${flor} - ${variedad || '(sin variedad)'} - ${color || '(sin color)'} - Grado ${grado || '(vacío)'}`;
  };

  let state = { version: 1, items: {} };
  let reunionFlowers = [];   // [name]
  let combos = [];           // [key]
  let selected = null;
  let saveTimer = null;
  let overlay = null;
  const ui = { flowerQuery: '', comboQuery: '', onlyPending: false, onlyUnassigned: false };

  // ---------------------------------------------------------------- data
  async function loadReunionFlowers() {
    const res = await fetch('/api/workbook?wb=reunion');
    if (!res.ok) throw new Error('No se pudo leer Reunion');
    const book = await res.json();
    const sheet = book.sheets && book.sheets.Reunion;
    if (!sheet) throw new Error('No se encontró la hoja Reunion');
    const cells = sheet.cells || {};
    // header row: the one whose column A says "Flor"; Cor column from the same row
    let headerRow = 2, corCol = 'E';
    for (let r = 1; r <= 10; r++) {
      if (norm((cells[`A${r}`] || {}).v) === 'FLOR') {
        headerRow = r;
        for (const [coord, cell] of Object.entries(cells)) {
          const m = /^([A-Z]+)(\d+)$/.exec(coord);
          if (m && Number(m[2]) === r && norm(cell.v) === 'COR') { corCol = m[1]; break; }
        }
        break;
      }
    }
    let last = headerRow;
    for (const coord of Object.keys(cells)) {
      const m = /^A(\d+)$/.exec(coord);
      if (m && Number(m[1]) > last && clean(cells[coord].v)) last = Number(m[1]);
    }
    const out = [];
    const seen = new Set();
    for (let r = headerRow + 1; r <= last; r++) {
      const name = clean((cells[`A${r}`] || {}).v);
      if (!name) continue;
      const cor = cells[`${corCol}${r}`];
      if (cor && cor.f) continue; // title row: its Cor adds up the flowers below
      if (seen.has(norm(name))) continue;
      seen.add(norm(name));
      out.push(name);
    }
    return out;
  }

  // combinations of the sheet you are on (all P Cosecha sheets if you are on another one),
  // only rows with a quantity in at least one week column
  let combosSource = '';
  function loadCombos() {
    const wb = App().getWorkbook();
    const engine = App().getEngine();
    const isCosecha = name => /^\s*p\s*cosecha/i.test(name);
    const active = App().getActiveSheet();
    const sheetNames = isCosecha(active) ? [active] : (wb.sheetOrder || []).filter(isCosecha);
    combosSource = isCosecha(active) ? clean(active) : 'todas las hojas P Cosecha';
    const keys = new Set();
    for (const name of sheetNames) {
      const sheet = wb.sheets[name];
      if (!sheet) continue;
      const cells = sheet.cells || {};
      const col = {};
      for (const [coord, cell] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(coord);
        if (!m || Number(m[2]) !== HEADER_ROW) continue;
        const h = norm(cell.v);
        if (h === 'FLOR' && !col.flor) col.flor = m[1];
        else if (h === 'VARIEDAD' && !col.variedad) col.variedad = m[1];
        else if (h === 'COLOR' && !col.color) col.color = m[1];
        else if (h === 'GRADO' && !col.grado) col.grado = m[1];
      }
      // week columns: a plain week number (1-53) in the header row
      const weekCols = [];
      for (const [coord, cell] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(coord);
        if (m && Number(m[2]) === HEADER_ROW && !cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) weekCols.push(m[1]);
      }
      if (!col.flor) continue;
      let last = FIRST_ROW;
      for (const coord of Object.keys(cells)) {
        const m = new RegExp(`^${col.flor}(\\d+)$`).exec(coord);
        if (m && Number(m[1]) > last) last = Number(m[1]);
      }
      const val = (c, r) => (c ? (cells[`${c}${r}`] || {}).v : '');
      for (let r = FIRST_ROW; r <= last; r++) {
        const flor = clean(val(col.flor, r));
        if (!flor) continue;
        const hasQty = weekCols.some(c => {
          const v = engine.getCellValue(name, `${c}${r}`);
          return typeof v === 'number' && v !== 0;
        });
        if (!hasQty) continue;
        keys.add(comboKey({ flor, variedad: val(col.variedad, r), color: val(col.color, r), grado: val(col.grado, r) }));
      }
    }
    return [...keys].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
  }

  async function loadState() {
    const res = await fetch('/api/homologations?wb=cosecha');
    const data = res.ok ? await res.json() : null;
    state = data && data.items && typeof data.items === 'object' ? data : { version: 1, items: {} };
    assignedKeys = null;
  }

  function saveSoon() {
    clearTimeout(saveTimer);
    assignedKeys = null;
    saveTimer = setTimeout(async () => {
      repaintGrid();
      try {
        const res = await fetch('/api/homologations?wb=cosecha', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(state)
        });
        if (!res.ok) throw new Error(String(res.status));
      } catch (e) {
        App().toast('No se pudieron guardar las homologaciones', true);
      }
    }, 300);
  }

  const assigned = flower => state.items[flower] || [];
  function usage() {
    const count = new Map();
    for (const list of Object.values(state.items)) for (const k of list) count.set(k, (count.get(k) || 0) + 1);
    return count;
  }
  function addCombo(flower, key) {
    const list = assigned(flower);
    if (list.includes(key)) return false;
    state.items[flower] = list.concat([key]);
    saveSoon();
    return true;
  }
  function toggleCombo(flower, key) {
    const list = assigned(flower).slice();
    const i = list.indexOf(key);
    if (i >= 0) list.splice(i, 1); else list.push(key);
    if (list.length) state.items[flower] = list; else delete state.items[flower];
    saveSoon();
    render();
  }

  // ---------------------------------------------------------------- ui
  function open() {
    if (overlay) return;
    overlay = document.createElement('div');
    overlay.className = 'hm-overlay';
    overlay.innerHTML = `
      <div class="hm-dialog" role="dialog" aria-label="Homologaciones">
        <div class="hm-bar" title="Arrastra para mover">
          <span class="hm-grip">⋮⋮</span>
          <button type="button" class="hm-close" title="Cerrar (Esc)">✕</button>
        </div>
        <div class="hm-body"><div class="hm-loading">Cargando…</div></div>
      </div>`;
    document.body.appendChild(overlay);
    overlay.querySelector('.hm-close').addEventListener('click', close);
    bindMove(overlay.querySelector('.hm-dialog'), overlay.querySelector('.hm-bar'));
    document.addEventListener('keydown', onKey, true);
    Promise.all([loadReunionFlowers(), loadState()])
      .then(([flowers]) => {
        if (!overlay) return; // closed while loading
        reunionFlowers = flowers;
        combos = loadCombos();
        if (!selected || !reunionFlowers.includes(selected)) selected = reunionFlowers[0] || null;
        buildBody();
      })
      .catch(err => {
        if (!overlay) return;
        overlay.querySelector('.hm-body').innerHTML = `<div class="hm-loading">${esc(err.message)}</div>`;
      });
  }
  function close() {
    if (!overlay) return;
    overlay.remove();
    overlay = null;
    document.removeEventListener('keydown', onKey, true);
  }
  function onKey(e) {
    if (!overlay || !overlay.contains(e.target)) return; // the sheet behind keeps working
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    else e.stopPropagation(); // typing in the panel must not reach the grid
  }

  // the panel floats over the sheet: drag its top bar to move it, its corner to resize it
  function bindMove(dialog, bar) {
    const r = dialog.getBoundingClientRect();
    dialog.style.left = `${Math.max(8, r.left)}px`;
    dialog.style.top = `${Math.max(8, r.top)}px`;
    overlay.classList.add('floating');
    bar.addEventListener('mousedown', e => {
      if (e.button !== 0 || e.target.closest('.hm-close')) return;
      e.preventDefault();
      const start = { x: e.clientX, y: e.clientY, left: dialog.offsetLeft, top: dialog.offsetTop };
      const move = ev => {
        const maxL = window.innerWidth - 80, maxT = window.innerHeight - 40;
        dialog.style.left = `${Math.min(maxL, Math.max(-dialog.offsetWidth + 80, start.left + ev.clientX - start.x))}px`;
        dialog.style.top = `${Math.min(maxT, Math.max(0, start.top + ev.clientY - start.y))}px`;
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        bar.classList.remove('moving');
      };
      bar.classList.add('moving');
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    });
  }

  function buildBody() {
    const body = overlay.querySelector('.hm-body');
    body.innerHTML = `
      <section class="hm-col hm-left">
        <div class="hm-col-head">
          <span>Flores de Reunion</span>
          <span class="hm-stat" data-stat="flowers"></span>
        </div>
        <input type="search" class="hm-search" data-q="flower" placeholder="Buscar flor…" value="${esc(ui.flowerQuery)}" />
        <label class="hm-check"><input type="checkbox" data-opt="onlyPending" ${ui.onlyPending ? 'checked' : ''}/> Solo sin homologar</label>
        <div class="hm-list" data-list="flowers"></div>
      </section>
      <section class="hm-col hm-right">
        <div class="hm-col-head hm-sub-head">
          <span>Plan de cosecha: Flor - Variedad - Color - Grado</span>
          <span class="hm-stat" data-stat="combos"></span>
        </div>
        <input type="search" class="hm-search" data-q="combo" placeholder="Buscar combinación…" value="${esc(ui.comboQuery)}" />
        <label class="hm-check"><input type="checkbox" data-opt="onlyUnassigned" ${ui.onlyUnassigned ? 'checked' : ''}/> Solo sin homologar</label>
        <div class="hm-list" data-list="combos"></div>
      </section>`;
    body.querySelector('[data-q="flower"]').addEventListener('input', e => { ui.flowerQuery = e.target.value; renderFlowers(); });
    body.querySelector('[data-q="combo"]').addEventListener('input', e => { ui.comboQuery = e.target.value; renderCombos(); });
    body.querySelectorAll('[data-opt]').forEach(el => el.addEventListener('change', e => { ui[e.target.dataset.opt] = e.target.checked; render(); }));
    body.querySelector('[data-list="flowers"]').addEventListener('click', e => {
      const item = e.target.closest('[data-flower]');
      if (!item) return;
      selected = item.dataset.flower;
      render();
    });
    bindDrag(body);
    // the ✕ on a highlighted combination removes it from the selected flower
    body.querySelector('[data-list="combos"]').addEventListener('click', e => {
      const x = e.target.closest('[data-remove]');
      if (x && selected) toggleCombo(selected, x.dataset.remove);
    });
    render();
  }

  // ---------------------------------------------------------------- drag & drop
  //   Reunion flower -> a combination (right): assigns that combination to the flower
  //   combination -> a flower (left) or the "Asignadas" box: assigns it
  //   assigned chip -> another flower: assigns it there too; -> the combinations list: removes it
  const DRAG_TYPE = 'text/plain';
  let dragging = null; // { key, fromChip } or { flower }
  function bindDrag(body) {
    body.addEventListener('dragstart', e => {
      const fl = e.target.closest('[data-flower]');
      if (fl) {
        dragging = { flower: fl.dataset.flower };
        e.dataTransfer.effectAllowed = 'copy';
        e.dataTransfer.setData(DRAG_TYPE, dragging.flower);
        body.classList.add('hm-dragging', 'hm-dragging-flower');
        fl.classList.add('hm-drag-src');
        return;
      }
      const combo = e.target.closest('[data-drag]');
      const chip = e.target.closest('[data-drag-chip]');
      if (!combo && !chip) return;
      dragging = { key: (combo || chip).dataset[combo ? 'drag' : 'dragChip'], fromChip: !!chip };
      e.dataTransfer.effectAllowed = 'copyMove';
      e.dataTransfer.setData(DRAG_TYPE, dragging.key);
      body.classList.add('hm-dragging');
      (combo || chip).classList.add('hm-drag-src');
    });
    body.addEventListener('dragend', () => {
      dragging = null;
      body.classList.remove('hm-dragging', 'hm-dragging-flower');
      body.querySelectorAll('.hm-drag-src, .hm-drop').forEach(el => el.classList.remove('hm-drag-src', 'hm-drop'));
    });
    const target = e => {
      if (!dragging) return null;
      if (dragging.flower) {
        const combo = e.target.closest('[data-drag]');
        return combo ? { el: combo, kind: 'combo' } : null;
      }
      const flower = e.target.closest('[data-flower]');
      if (flower) return { el: flower, kind: 'flower' };
      const chips = e.target.closest('[data-chips]');
      if (chips && selected && !dragging.fromChip) return { el: chips, kind: 'chips' };
      const list = e.target.closest('[data-list="combos"]');
      if (list && dragging.fromChip) return { el: list, kind: 'remove' };
      return null;
    };
    body.addEventListener('dragover', e => {
      const t = target(e);
      if (!t) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = t.kind === 'remove' ? 'move' : 'copy';
      if (!t.el.classList.contains('hm-drop')) {
        body.querySelectorAll('.hm-drop').forEach(el => el.classList.remove('hm-drop'));
        t.el.classList.add('hm-drop');
      }
    });
    body.addEventListener('dragleave', e => {
      const el = e.target.closest('.hm-drop');
      if (el && !el.contains(e.relatedTarget)) el.classList.remove('hm-drop');
    });
    body.addEventListener('drop', e => {
      const t = target(e);
      if (!t) return;
      e.preventDefault();
      const key = dragging.key;
      if (t.kind === 'combo') {
        const flower = dragging.flower;
        const k = t.el.dataset.drag;
        const added = addCombo(flower, k);
        selected = flower;
        App().toast(added ? `${flower} → ${comboLabel(k)}` : `${flower} ya la tenía`);
      } else if (t.kind === 'flower') {
        const flower = t.el.dataset.flower;
        const added = addCombo(flower, key);
        selected = flower;
        App().toast(added ? `${comboLabel(key)} → ${flower}` : `${flower} ya la tenía`);
      } else if (t.kind === 'chips') {
        addCombo(selected, key);
      } else if (t.kind === 'remove') {
        toggleCombo(selected, key);
        return;
      }
      render();
    });
  }

  function matches(text, query) {
    const words = norm(query).split(' ').filter(Boolean);
    const t = norm(text);
    return words.every(w => t.includes(w));
  }

  function renderFlowers() {
    const list = overlay.querySelector('[data-list="flowers"]');
    const rows = reunionFlowers.filter(f => matches(f, ui.flowerQuery) && (!ui.onlyPending || !assigned(f).length));
    list.innerHTML = rows.map(f => {
      const n = assigned(f).length;
      return `<div class="hm-item${f === selected ? ' sel' : ''}${n ? '' : ' pending'}" data-flower="${esc(f)}" draggable="true">
        <span class="hm-name">${esc(f)}</span><span class="hm-badge">${n}</span></div>`;
    }).join('') || '<div class="hm-empty">Sin resultados</div>';
    const done = reunionFlowers.filter(f => assigned(f).length).length;
    overlay.querySelector('[data-stat="flowers"]').textContent = `${done} de ${reunionFlowers.length} homologadas`;
  }

  function renderCombos() {
    const list = overlay.querySelector('[data-list="combos"]');
    const used = usage();
    const mine = new Set(selected ? assigned(selected) : []);
    const rows = combos.filter(k => matches(comboLabel(k), ui.comboQuery) && (!ui.onlyUnassigned || !used.get(k)));
    list.innerHTML = rows.map(k => {
      const n = used.get(k) || 0;
      return `<div class="hm-item hm-combo${n ? '' : ' pending'}${mine.has(k) ? ' mine' : ''}" draggable="true" data-drag="${esc(k)}">
        <span class="hm-name">${esc(comboLabel(k))}</span>
        <span class="hm-uses">${n ? `${n} flor${n > 1 ? 'es' : ''}` : 'sin homologar'}</span>${mine.has(k) ? `<button type="button" class="hm-x" data-remove="${esc(k)}" title="Quitar de ${esc(selected)}">✕</button>` : ''}</div>`;
    }).join('') || '<div class="hm-empty">Sin resultados</div>';
    const free = combos.filter(k => !used.get(k)).length;
    overlay.querySelector('[data-stat="combos"]').textContent = `${combosSource}: ${combos.length} con cantidad · ${free} sin homologar`;
  }

  function render() {
    if (!overlay || !overlay.querySelector('[data-list]')) return;
    renderFlowers();
    renderCombos();
  }

  // ---------------------------------------------------------------- grid: Flor in light red when it has
  // a quantity in some week but its Flor - Variedad - Color - Grado is not homologated yet
  const PENDING_STYLE = { bg: '#fde2e2' };
  let assignedKeys = null;
  const sheetCols = new WeakMap();
  function columnsOf(sheet) {
    let info = sheetCols.get(sheet);
    if (info) return info;
    info = { weeks: [] };
    for (const [coord, cell] of Object.entries(sheet.cells || {})) {
      const m = /^([A-Z]+)(\d+)$/.exec(coord);
      if (!m || Number(m[2]) !== HEADER_ROW) continue;
      const h = norm(cell.v);
      if (h === 'FLOR' && !info.flor) info.flor = m[1];
      else if (h === 'VARIEDAD' && !info.variedad) info.variedad = m[1];
      else if (h === 'COLOR' && !info.color) info.color = m[1];
      else if (h === 'GRADO' && !info.grado) info.grado = m[1];
      else if (!cell.f && Number.isInteger(cell.v) && cell.v >= 1 && cell.v <= 53) info.weeks.push(m[1]);
    }
    sheetCols.set(sheet, info);
    return info;
  }
  function cellStyle(sheetName, row, colStr, sheet) {
    if (!stateLoaded || row < FIRST_ROW || !/^\s*p\s*cosecha/i.test(sheetName)) return null;
    const info = columnsOf(sheet);
    if (colStr !== info.flor) return null;
    const val = c => (c ? (sheet.cells[`${c}${row}`] || {}).v : '');
    const flor = clean(val(info.flor));
    if (!flor) return null;
    const engine = App().getEngine();
    const hasQty = info.weeks.some(c => {
      const v = engine.getCellValue(sheetName, `${c}${row}`);
      return typeof v === 'number' && v !== 0;
    });
    if (!hasQty) return null;
    if (!assignedKeys) assignedKeys = new Set(Object.values(state.items).flat());
    const key = comboKey({ flor, variedad: val(info.variedad), color: val(info.color), grado: val(info.grado) });
    return assignedKeys.has(key) ? null : PENDING_STYLE;
  }
  function repaintGrid() {
    const app = App();
    if (app && app.getWorkbook() && app.getEngine()) app.render(); // colors only
  }

  let stateLoaded = false;
  loadState().then(() => {
    stateLoaded = true;
    assignedKeys = null;
    const wait = setInterval(() => {
      const app = App();
      if (!app || !app.getWorkbook() || !app.getEngine()) return;
      clearInterval(wait);
      repaintGrid();
    }, 300);
  }).catch(() => {});

  btn.addEventListener('click', open);
  window.MauriHomolog = { open, close, cellStyle, getState: () => state };
})();
