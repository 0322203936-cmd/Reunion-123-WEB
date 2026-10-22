// ============================================================================
// PIVOT TABLES (Tablas dinámicas) - Excel-like compact layout in Spanish
//   * workbook.pivots = [{ id, name, sheet, location, source:{sheet,ref},
//       rows:[field], cols:[field|'__values'], values:[{field,agg,name}],
//       filters:[{field,selected:null|[values]}], hidden:{field:[values]},
//       grandRows, grandCols, area }]
//   * Imported pivots keep Excel's output until "Actualizar"
// ============================================================================

(function () {
  const { indexToCol, parseCoord, colToIndex } = window.ExcelUtils;
  const App = () => window.MauriApp;
  const VALUES = '__values';
  const AGG_LABEL = { sum: 'Suma', count: 'Cuenta', average: 'Promedio', max: 'Máx', min: 'Mín', product: 'Producto', countNums: 'Contar números' };
  const T = {
    rowLabels: 'Etiquetas de fila', colLabels: 'Etiquetas de columna', grand: 'Total general', all: '(Todas)',
    many: '(Varios elementos)', blank: '(en blanco)', values: 'Valores', total: 'Total'
  };
  const STYLE = {
    header: { bold: true, bg: '#dce6f1', border: { bottom: true } },
    subtotal: { bold: true },
    grand: { bold: true, bg: '#dce6f1', border: { top: true } },
    filterLabel: { bold: true },
    item: {}
  };

  const wb = () => App().getWorkbook();
  const pivots = () => (wb().pivots = wb().pivots || []);

  function escapeHtml(s) {
    return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function parseRange(ref) {
    const [a, b] = String(ref).replace(/\$/g, '').split(':').map(parseCoord);
    if (!a) return null;
    const e = b || a;
    return { r1: Math.min(a.row, e.row), r2: Math.max(a.row, e.row), c1: Math.min(a.col, e.col), c2: Math.max(a.col, e.col) };
  }
  const rangeStr = g => `${indexToCol(g.c1)}${g.r1}:${indexToCol(g.c2)}${g.r2}`;

  // ---------------------------------------------------------------- source data
  function readSource(p) {
    const g = parseRange(p.source.ref);
    const engine = App().getEngine();
    if (!g || !wb().sheets[p.source.sheet]) throw new Error(`No se encontró el origen ${p.source.sheet}!${p.source.ref}`);
    const fields = [];
    for (let c = g.c1; c <= g.c2; c++) {
      const v = engine.getCellValue(p.source.sheet, `${indexToCol(c)}${g.r1}`);
      fields.push(v === '' || v === null || v === undefined ? `Columna${c - g.c1 + 1}` : String(v));
    }
    const records = [];
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const rec = {};
      let any = false;
      fields.forEach((f, i) => {
        let v = engine.getCellValue(p.source.sheet, `${indexToCol(g.c1 + i)}${r}`);
        if (v === '' || v === undefined) v = null;
        if (v !== null) any = true;
        rec[f] = v;
      });
      if (any) records.push(rec);
    }
    return { fields, records };
  }

  const itemKey = v => (v === null || v === undefined || v === '' ? '\u0000blank' : typeof v === 'number' ? `n:${v}` : `s:${String(v).toLowerCase()}`);
  const itemLabel = v => (v === null || v === undefined || v === '' ? T.blank : v);
  function compareItems(a, b) {
    const ka = a === null || a === undefined || a === '', kb = b === null || b === undefined || b === '';
    if (ka || kb) return ka === kb ? 0 : ka ? 1 : -1;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return String(a).localeCompare(String(b), 'es', { sensitivity: 'base', numeric: true });
  }

  function aggregate(agg, values) {
    const nums = values.filter(v => typeof v === 'number');
    switch (agg) {
      case 'count': { const n = values.filter(v => v !== null && v !== undefined && v !== '').length; return n || null; }
      case 'countNums': return nums.length || null;
      case 'average': return nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : null;
      case 'max': return nums.length ? Math.max(...nums) : null;
      case 'min': return nums.length ? Math.min(...nums) : null;
      case 'product': return nums.length ? nums.reduce((s, x) => s * x, 1) : null;
      default: return nums.length ? Number(nums.reduce((s, x) => s + x, 0).toPrecision(15)) : null;
    }
  }

  function passesFilters(p, rec) {
    for (const [field, list] of Object.entries(p.hidden || {})) {
      if (!list || !list.length || !(field in rec)) continue;
      const k = itemKey(rec[field]);
      if (list.some(h => itemKey(h) === k)) return false;
    }
    for (const f of p.filters || []) {
      if (!f.selected || !(f.field in rec)) continue;
      const k = itemKey(rec[f.field]);
      if (!f.selected.some(s => itemKey(s) === k)) return false;
    }
    return true;
  }

  // order of a field's items: explicit sort, else Excel's stored item order (new items at the end), else ascending
  function itemSorter(p, field) {
    const dir = (p.sort || {})[field];
    if (dir === 'descending') return (a, b) => compareItems(b.value, a.value);
    const order = !dir && (p.order || {})[field];
    if (!order) return (a, b) => compareItems(a.value, b.value);
    const pos = new Map(order.map((v, i) => [itemKey(v), i]));
    const at = v => (pos.has(itemKey(v)) ? pos.get(itemKey(v)) : Infinity);
    return (a, b) => {
      const pa = at(a.value), pb = at(b.value);
      return pa !== pb ? pa - pb : compareItems(a.value, b.value);
    };
  }

  // nested tree of distinct items for a list of fields
  function buildTree(p, records, fields) {
    const root = { children: new Map(), records };
    for (const rec of records) {
      let node = root;
      for (const f of fields) {
        const k = itemKey(rec[f]);
        if (!node.children.has(k)) node.children.set(k, { value: rec[f], children: new Map(), records: [] });
        node = node.children.get(k);
        node.records.push(rec);
      }
    }
    const sortNode = (n, depth) => {
      n.list = [...n.children.values()];
      if (depth < fields.length) n.list.sort(itemSorter(p, fields[depth]));
      n.list.forEach(ch => sortNode(ch, depth + 1));
    };
    sortNode(root, 0);
    return root;
  }

  const isCollapsed = (p, field, value) => ((p.collapsed || {})[field] || []).some(v => itemKey(v) === itemKey(value));

  // label of a page filter: (Todas), the single visible item, or (Varios elementos)
  function filterLabel(p, f, records) {
    if (f.selected && f.selected.length === 1) return itemLabel(f.selected[0]);
    if (f.selected && f.selected.length > 1) return T.many;
    const hidden = (p.hidden || {})[f.field];
    if (!hidden || !hidden.length) return T.all;
    const hiddenKeys = new Set(hidden.map(itemKey));
    const visible = new Map();
    for (const rec of records) {
      const k = itemKey(rec[f.field]);
      if (!hiddenKeys.has(k) && !visible.has(k)) visible.set(k, rec[f.field]);
      if (visible.size > 1) return T.many;
    }
    return visible.size === 1 ? itemLabel([...visible.values()][0]) : T.many;
  }

  // ---------------------------------------------------------------- layout
  // compact: all row fields in one indented column; outline/tabular: one column per row field
  function computeGrid(p) {
    const { fields, records } = readSource(p);
    const data = records.filter(r => passesFilters(p, r));
    const rowFields = (p.rows || []).filter(f => f !== VALUES && fields.includes(f));
    const colFields = (p.cols || []).filter(f => f !== VALUES && fields.includes(f));
    const values = (p.values || []).filter(v => fields.includes(v.field));
    const layout = p.layout || 'compact';
    const compact = layout === 'compact';
    const tabular = layout === 'tabular';
    const valuesInRows = (p.rows || []).includes(VALUES) && values.length > 1 && compact;
    const lab = compact ? 1 : Math.max(1, rowFields.length); // columns used by row labels
    const grid = []; // rows of {v, s}
    const put = (r, c, v, s) => { (grid[r] = grid[r] || [])[c] = { v, s }; };
    const indentStyle = (base, depth) => (depth ? { ...base, indent: depth } : { ...base });

    // ---- filters block (above the table)
    let r0 = 0;
    for (const f of p.filters || []) {
      put(r0, 0, f.field, STYLE.filterLabel);
      put(r0, 1, filterLabel(p, f, records), STYLE.item);
      r0++;
    }
    if ((p.filters || []).length) r0++; // blank row between filters and table
    const tableTop = r0;

    // ---- column axis: leaves = combinations of col items (+ values level)
    const colTree = buildTree(p, data, colFields);
    const colLeaves = [];
    const walkCols = (node, path) => {
      if (!node.list.length || path.length === colFields.length) { colLeaves.push({ path, records: node.records }); return; }
      node.list.forEach(ch => walkCols(ch, path.concat([ch.value])));
    };
    if (colFields.length) walkCols(colTree, []); else colLeaves.push({ path: [], records: data });
    const valueList = values.length ? values : [];
    const multiValuesCols = !valuesInRows && valueList.length > 1;
    const showGrandCol = colFields.length > 0 && p.grandCols !== false;

    const headerRows = colFields.length + (multiValuesCols && colFields.length ? 1 : 0) + 1;
    const columns = []; // [{ leaf, value }]
    for (const leaf of colLeaves) {
      if (valuesInRows || !valueList.length) columns.push({ leaf, value: valueList[0] || null });
      else valueList.forEach(v => columns.push({ leaf, value: v }));
    }
    if (showGrandCol) {
      if (valuesInRows || valueList.length <= 1) columns.push({ leaf: { path: null, records: data }, value: valueList[0] || null, grand: true });
      else valueList.forEach(v => columns.push({ leaf: { path: null, records: data }, value: v, grand: true }));
    }
    const rowHeader = rr => {
      if (compact || !rowFields.length) put(rr, 0, T.rowLabels, STYLE.header);
      else rowFields.forEach((f, i) => put(rr, i, f, STYLE.header));
    };

    if (colFields.length) {
      put(tableTop, 0, valueList.length === 1 ? valueList[0].name : '', STYLE.header);
      for (let c = 1; c < lab; c++) put(tableTop, c, '', STYLE.header);
      put(tableTop, lab, compact ? T.colLabels : colFields[0], STYLE.header);
      for (let c = lab + 1; c < lab + columns.length; c++) put(tableTop, c, '', STYLE.header);
      colFields.forEach((f, level) => {
        const rr = tableTop + 1 + level;
        if (level === colFields.length - 1 && !multiValuesCols) rowHeader(rr);
        else for (let c = 0; c < lab; c++) put(rr, c, '', STYLE.header);
        let prev = null;
        columns.forEach((col, i) => {
          if (col.grand) { put(rr, lab + i, level === 0 ? (multiValuesCols ? `${T.total} ${col.value.name}` : T.grand) : '', STYLE.header); return; }
          const label = itemLabel(col.leaf.path[level]);
          const key = col.leaf.path.slice(0, level + 1).map(itemKey).join('|');
          put(rr, lab + i, key !== prev || !multiValuesCols ? label : '', STYLE.header);
          prev = key;
        });
      });
      if (multiValuesCols) {
        const rr = tableTop + headerRows - 1;
        rowHeader(rr);
        columns.forEach((col, i) => put(rr, lab + i, col.grand ? '' : col.value.name, STYLE.header));
      }
    } else {
      if (rowFields.length || !valueList.length) rowHeader(tableTop);
      else put(tableTop, 0, '', STYLE.header);
      columns.forEach((col, i) => put(tableTop, lab + i, col.value ? col.value.name : '', STYLE.header));
    }

    // ---- body rows
    let rr = tableTop + headerRows;
    const cellValue = (records, col, v = col.value) => {
      if (!v) return null;
      const subset = col.grand || !colFields.length ? records : records.filter(rec => col.leaf.path.every((x, i) => itemKey(rec[colFields[i]]) === itemKey(x)));
      return aggregate(v.agg, subset.map(rec => rec[v.field] ?? null));
    };
    const emitValues = (records, style) => {
      if (!valueList.length) return;
      columns.forEach((col, i) => put(rr, lab + i, cellValue(records, col), style));
    };
    const rowTree = buildTree(p, data, rowFields);

    // compact / outline: the parent row carries its subtotal, children follow
    const walkTree = (node, depth) => {
      for (const ch of node.list) {
        const isLeaf = depth === rowFields.length - 1;
        const open = !isLeaf && !isCollapsed(p, rowFields[depth], ch.value);
        const style = isLeaf ? { ...STYLE.item } : { ...STYLE.subtotal };
        const label = itemLabel(ch.value);
        if (compact) put(rr, 0, label, indentStyle(style, depth));
        else put(rr, depth, label, style);
        if (valuesInRows) {
          rr++;
          valueList.forEach(v => {
            put(rr, 0, v.name, indentStyle(STYLE.item, depth + 1));
            columns.forEach((col, i) => put(rr, lab + i, cellValue(ch.records, col, v), STYLE.item));
            rr++;
          });
        } else {
          emitValues(ch.records, style);
          rr++;
        }
        if (open) walkTree(ch, depth + 1);
      }
    };
    // tabular: the first child shares its parent's row, a "Total X" row closes each group
    const walkTabular = (node, depth) => {
      for (const ch of node.list) {
        const isLeaf = depth === rowFields.length - 1;
        const open = !isLeaf && !isCollapsed(p, rowFields[depth], ch.value);
        put(rr, depth, itemLabel(ch.value), STYLE.item);
        if (!open) { emitValues(ch.records, isLeaf ? STYLE.item : STYLE.subtotal); rr++; continue; }
        walkTabular(ch, depth + 1);
        put(rr, depth, `${T.total} ${itemLabel(ch.value)}`, STYLE.subtotal);
        for (let c = depth + 1; c < lab; c++) put(rr, c, '', STYLE.subtotal);
        emitValues(ch.records, STYLE.subtotal);
        rr++;
      }
    };
    if (rowFields.length) (tabular ? walkTabular : walkTree)(rowTree, 0);
    else if (valueList.length) { put(rr, 0, '', STYLE.item); emitValues(data, STYLE.item); rr++; }

    // grand total row
    if (p.grandRows !== false && (rowFields.length || !valueList.length)) {
      put(rr, 0, T.grand, STYLE.grand);
      for (let c = 1; c < lab; c++) put(rr, c, '', STYLE.grand);
      if (!valuesInRows && valueList.length) columns.forEach((col, i) => put(rr, lab + i, cellValue(data, col), STYLE.grand));
      else columns.forEach((_, i) => put(rr, lab + i, '', STYLE.grand));
      rr++;
    }
    if (!rowFields.length && !colFields.length && !valueList.length) {
      grid.length = 0;
      put(0, 0, 'Para generar un informe, elija los campos', { italic: true, color: '#64748b' });
      put(1, 0, 'de la lista de campos de la tabla dinámica.', { italic: true, color: '#64748b' });
      return { grid, tableTop: 0, fields };
    }
    return { grid, tableTop, fields };
  }

  // write a pivot into its sheet; returns the area written
  function renderPivot(p) {
    const sheet = wb().sheets[p.sheet];
    if (!sheet) return;
    const { grid, tableTop } = computeGrid(p);
    const anchor = p.anchor ? parseCoord(p.anchor) : null;
    const loc = parseRange(p.location);
    const top = anchor ? anchor.row : loc.r1 - ((p.filters || []).length ? (p.filters.length + 1) : 0);
    const left = anchor ? anchor.col : loc.c1;
    // clear the previous output
    const old = p.area ? parseRange(p.area) : null;
    const clearArea = old || { r1: Math.max(1, top), r2: loc ? loc.r2 : top, c1: left, c2: loc ? loc.c2 : left };
    const gapRow = (p.filters || []).length ? top + p.filters.length : null; // not part of the pivot: user text there stays
    for (let r = clearArea.r1; r <= clearArea.r2; r++) {
      if (r === gapRow) continue;
      for (let c = clearArea.c1; c <= clearArea.c2; c++) delete sheet.cells[`${indexToCol(c)}${r}`];
    }
    let maxC = 0;
    grid.forEach((row, i) => {
      if (!row) return;
      row.forEach((cell, j) => {
        if (!cell) return;
        maxC = Math.max(maxC, j);
        const coord = `${indexToCol(left + j)}${top + i}`;
        const entry = {};
        if (cell.v !== null && cell.v !== undefined && cell.v !== '') entry.v = cell.v;
        if (cell.s && Object.keys(cell.s).length) entry.s = JSON.parse(JSON.stringify(cell.s));
        if (Object.keys(entry).length) sheet.cells[coord] = entry;
      });
    });
    const bottom = top + Math.max(grid.length, 1) - 1;
    p.anchor = `${indexToCol(left)}${Math.max(1, top)}`;
    p.location = `${indexToCol(left)}${top + tableTop}:${indexToCol(left + maxC)}${bottom}`;
    p.area = `${indexToCol(left)}${Math.max(1, top)}:${indexToCol(left + maxC)}${bottom}`;
    sheet.maxRow = Math.max(sheet.maxRow || 100, bottom + 20);
    sheet.maxCol = Math.max(sheet.maxCol || 20, left + maxC + 3);
    const widths = sheet.colWidths = sheet.colWidths || {};
    const firstCol = indexToCol(left);
    if (!widths[firstCol]) widths[firstCol] = 22;
  }

  function areaOf(p) {
    if (p.area) return parseRange(p.area);
    const loc = parseRange(p.location);
    if (!loc) return null;
    const extra = (p.filters || []).length ? p.filters.length + 1 : 0;
    return { ...loc, r1: Math.max(1, loc.r1 - extra) };
  }

  function pivotAt(sheetName, coord) {
    const pt = parseCoord(coord);
    if (!pt) return null;
    return pivots().find(p => {
      if (p.sheet !== sheetName) return false;
      const a = areaOf(p);
      return a && pt.row >= a.r1 && pt.row <= a.r2 && pt.col >= a.c1 && pt.col <= a.c2;
    }) || null;
  }

  function commit(p, message) {
    try {
      renderPivot(p);
    } catch (e) {
      App().toast(e.message, true);
      return;
    }
    App().markDirty(p.sheet);
    App().recalcAndRender();
    App().save();
    if (message) App().toast(message);
  }

  // ---------------------------------------------------------------- create dialog
  function currentRegion(sheetName, coord) {
    const sheet = wb().sheets[sheetName];
    const filled = (r, c) => {
      const cell = sheet.cells[`${indexToCol(c)}${r}`];
      return !!(cell && ((cell.v !== undefined && cell.v !== null && cell.v !== '') || cell.f));
    };
    const p = parseCoord(coord);
    let g = { r1: p.row, r2: p.row, c1: p.col, c2: p.col };
    let grew = true;
    while (grew) {
      grew = false;
      const rowHas = r => { for (let c = Math.max(1, g.c1 - 1); c <= g.c2 + 1; c++) if (filled(r, c)) return true; return false; };
      const colHas = c => { for (let r = Math.max(1, g.r1 - 1); r <= g.r2 + 1; r++) if (filled(r, c)) return true; return false; };
      if (g.r1 > 1 && rowHas(g.r1 - 1)) { g.r1--; grew = true; }
      if (rowHas(g.r2 + 1)) { g.r2++; grew = true; }
      if (g.c1 > 1 && colHas(g.c1 - 1)) { g.c1--; grew = true; }
      if (colHas(g.c2 + 1)) { g.c2++; grew = true; }
      if (g.r2 - g.r1 > 200000) break;
    }
    return g;
  }

  function uniqueSheetName() {
    let i = 1;
    while (wb().sheets[`Hoja${i}`]) i++;
    return `Hoja${i}`;
  }

  function uniquePivotName() {
    let i = 1;
    const names = new Set(pivots().map(p => p.name));
    while (names.has(`TablaDinámica${i}`)) i++;
    return `TablaDinámica${i}`;
  }

  function openCreateDialog() {
    const sheetName = App().getActiveSheet();
    const sel = App().getSelectionBounds();
    let g = sel.isMulti ? sel : currentRegion(sheetName, `${indexToCol(sel.minCol)}${sel.minRow}`);
    if (sel.isMulti) g = { r1: sel.minRow, r2: sel.maxRow, c1: sel.minCol, c2: sel.maxCol };
    const needsQuote = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName);
    const srcText = `${needsQuote ? `'${sheetName.replace(/'/g, "''")}'` : sheetName}!${rangeStr(g)}`;

    const wrap = document.createElement('div');
    wrap.className = 'pv-overlay no-print';
    wrap.innerHTML = `
      <div class="pv-dialog" role="dialog" aria-modal="true">
        <div class="pv-dialog-title">Crear tabla dinámica</div>
        <div class="pv-dialog-body">
          <label class="pv-label">Seleccione una tabla o rango</label>
          <input class="pv-input" id="pvSource" value="${escapeHtml(srcText)}" />
          <label class="pv-label">Elija dónde desea colocar el informe de tabla dinámica</label>
          <label class="pv-radio"><input type="radio" name="pvDest" value="new" checked /> Nueva hoja de cálculo</label>
          <label class="pv-radio"><input type="radio" name="pvDest" value="existing" /> Hoja de cálculo existente</label>
          <input class="pv-input" id="pvDestRef" placeholder="Ubicación, p. ej. Hoja1!A3" disabled />
          <div class="pv-error" id="pvError"></div>
        </div>
        <div class="pv-dialog-actions">
          <button type="button" class="pv-btn primary" data-act="ok">Aceptar</button>
          <button type="button" class="pv-btn" data-act="cancel">Cancelar</button>
        </div>
      </div>`;
    document.body.appendChild(wrap);
    document.body.classList.add('modal-open');
    const close = () => { wrap.remove(); document.body.classList.remove('modal-open'); };
    const destRef = wrap.querySelector('#pvDestRef');
    wrap.querySelectorAll('input[name="pvDest"]').forEach(r => r.addEventListener('change', () => {
      destRef.disabled = wrap.querySelector('input[name="pvDest"]:checked').value !== 'existing';
      if (!destRef.disabled) destRef.focus();
    }));
    wrap.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
    wrap.addEventListener('click', e => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (act === 'cancel') return close();
      if (act !== 'ok') return;
      const err = wrap.querySelector('#pvError');
      const src = parseSheetRef(wrap.querySelector('#pvSource').value, sheetName);
      if (!src || !parseRange(src.ref) || !wb().sheets[src.sheet]) { err.textContent = 'La referencia del origen no es válida.'; return; }
      const sg = parseRange(src.ref);
      if (sg.r2 - sg.r1 < 1) { err.textContent = 'El rango debe tener encabezados y al menos una fila de datos.'; return; }
      let target;
      if (wrap.querySelector('input[name="pvDest"]:checked').value === 'new') {
        const name = uniqueSheetName();
        App().pushUndo([name]);
        const w = wb();
        w.sheets[name] = { maxRow: 100, maxCol: 20, colWidths: {}, merges: [], cells: {} };
        const idx = w.sheetOrder.indexOf(sheetName);
        w.sheetOrder.splice(idx >= 0 ? idx : w.sheetOrder.length, 0, name);
        target = { sheet: name, anchor: 'A3' };
      } else {
        const t = parseSheetRef(destRef.value, sheetName);
        if (!t || !parseCoord(t.ref.split(':')[0]) || !wb().sheets[t.sheet]) { err.textContent = 'La ubicación de destino no es válida.'; return; }
        App().pushUndo([t.sheet]);
        target = { sheet: t.sheet, anchor: parseCoord(t.ref.split(':')[0]).coord };
      }
      const p = {
        id: `pv${Date.now()}`, name: uniquePivotName(), sheet: target.sheet, anchor: target.anchor,
        location: `${target.anchor}:${target.anchor}`, source: src,
        rows: [], cols: [], values: [], filters: [], hidden: {}, grandRows: true, grandCols: true
      };
      pivots().push(p);
      close();
      renderPivot(p);
      App().markDirty(p.sheet);
      App().recalcAndRender();
      if (App().getActiveSheet() !== p.sheet) App().switchSheet(p.sheet);
      App().selectCoord(p.anchor);
      App().save();
      openPanel(p);
    });
    setTimeout(() => wrap.querySelector('#pvSource').select(), 0);
  }

  function parseSheetRef(text, defaultSheet) {
    const s = String(text || '').trim().replace(/^=/, '');
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(s);
    if (m) return { sheet: (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]).trim(), ref: m[3].replace(/\$/g, '').trim().toUpperCase() };
    return { sheet: defaultSheet, ref: s.replace(/\$/g, '').toUpperCase() };
  }

  // ---------------------------------------------------------------- field list panel
  let panel = null;
  let panelPivotId = null;

  function closePanel() {
    if (panel) panel.remove();
    panel = null;
    panelPivotId = null;
    document.body.classList.remove('pv-panel-open');
  }

  function openPanel(p) {
    let fields;
    try { fields = readSource(p).fields; } catch (e) { App().toast(e.message, true); return; }
    if (!panel) {
      panel = document.createElement('aside');
      panel.className = 'pv-panel no-print';
      document.body.appendChild(panel);
      document.body.classList.add('pv-panel-open');
    }
    panelPivotId = p.id;
    const used = new Set([...p.rows, ...p.cols, ...p.values.map(v => v.field), ...p.filters.map(f => f.field)]);
    const chip = (zone, field, label, extra = '') => `
      <div class="pv-chip" draggable="true" data-zone="${zone}" data-field="${escapeHtml(field)}">
        <span>${escapeHtml(label)}</span>
        <span class="pv-chip-actions">${extra}<button type="button" class="pv-x" data-remove title="Quitar">✕</button></span>
      </div>`;
    const zoneHtml = (zone, title, items) => `
      <div class="pv-zone" data-zone="${zone}">
        <div class="pv-zone-title">${title}</div>
        <div class="pv-zone-body">${items || '<div class="pv-zone-empty">Arrastrá campos aquí</div>'}</div>
      </div>`;
    const aggSelect = v => `<select class="pv-agg" data-field="${escapeHtml(v.field)}" title="Resumir valores por">${Object.entries(AGG_LABEL).map(([k, l]) => `<option value="${k}"${k === v.agg ? ' selected' : ''}>${l}</option>`).join('')}</select>`;
    const filterBtn = (zone, f) => `<button type="button" class="pv-filter-btn${((p.hidden || {})[f] || []).length || (zone === 'filters' && (p.filters.find(x => x.field === f) || {}).selected) ? ' active' : ''}" data-filter="${escapeHtml(f)}" title="Filtrar">⏷</button>`;

    panel.innerHTML = `
      <header class="pv-panel-head">
        <div>
          <div class="pv-panel-title">Campos de tabla dinámica</div>
          <div class="pv-panel-sub">${escapeHtml(p.name)} · origen ${escapeHtml(p.source.sheet)}!${escapeHtml(p.source.ref)}</div>
        </div>
        <button type="button" class="pv-close" data-close title="Cerrar">✕</button>
      </header>
      <input type="search" class="pv-search" placeholder="Buscar campo" />
      <div class="pv-fields">
        ${fields.map(f => `<label class="pv-field" draggable="true" data-field="${escapeHtml(f)}"><input type="checkbox" ${used.has(f) ? 'checked' : ''} data-toggle="${escapeHtml(f)}" /> ${escapeHtml(f)}</label>`).join('')}
      </div>
      <div class="pv-zones">
        ${zoneHtml('filters', 'Filtros', p.filters.map(f => chip('filters', f.field, f.field, filterBtn('filters', f.field))).join(''))}
        ${zoneHtml('cols', 'Columnas', p.cols.map(f => chip('cols', f, f === VALUES ? `Σ ${T.values}` : f, f === VALUES ? '' : filterBtn('cols', f))).join(''))}
        ${zoneHtml('rows', 'Filas', p.rows.map(f => chip('rows', f, f === VALUES ? `Σ ${T.values}` : f, f === VALUES ? '' : filterBtn('rows', f))).join(''))}
        ${zoneHtml('values', 'Valores', p.values.map(v => chip('values', v.field, v.name, aggSelect(v))).join(''))}
      </div>
      <footer class="pv-panel-foot">
        <button type="button" class="pv-btn" data-delete>Eliminar tabla</button>
        <button type="button" class="pv-btn primary" data-refresh>Actualizar</button>
      </footer>`;
    bindPanel(p, fields);
  }

  function currentPivot() { return pivots().find(x => x.id === panelPivotId); }

  function ensureValuesPseudoField(p) {
    const inCols = p.cols.includes(VALUES), inRows = p.rows.includes(VALUES);
    if (p.values.length > 1 && !inCols && !inRows) p.cols.push(VALUES);
    if (p.values.length <= 1) { p.cols = p.cols.filter(f => f !== VALUES); p.rows = p.rows.filter(f => f !== VALUES); }
  }

  function removeField(p, zone, field) {
    if (zone === 'rows') p.rows = p.rows.filter(f => f !== field);
    if (zone === 'cols') p.cols = p.cols.filter(f => f !== field);
    if (zone === 'filters') p.filters = p.filters.filter(f => f.field !== field);
    if (zone === 'values') {
      const i = p.values.findIndex(v => v.field === field);
      if (i >= 0) p.values.splice(i, 1);
    }
    ensureValuesPseudoField(p);
  }

  function addField(p, zone, field, isNumeric) {
    if (field === VALUES) {
      p.rows = p.rows.filter(f => f !== VALUES); p.cols = p.cols.filter(f => f !== VALUES);
      (zone === 'rows' ? p.rows : p.cols).push(VALUES);
      return;
    }
    // a field lives in only one of rows/cols/filters (values may repeat it)
    if (zone !== 'values') {
      p.rows = p.rows.filter(f => f !== field);
      p.cols = p.cols.filter(f => f !== field);
      p.filters = p.filters.filter(f => f.field !== field);
    }
    if (zone === 'rows') p.rows.push(field);
    else if (zone === 'cols') p.cols.splice(p.cols.includes(VALUES) ? p.cols.indexOf(VALUES) : p.cols.length, 0, field);
    else if (zone === 'filters') p.filters.push({ field, selected: null });
    else if (zone === 'values') {
      const agg = isNumeric ? 'sum' : 'count';
      p.values.push({ field, agg, name: `${AGG_LABEL[agg]} de ${field}` });
    }
    ensureValuesPseudoField(p);
  }

  function isNumericField(p, field) {
    try {
      const { records } = readSource(p);
      let nums = 0, vals = 0;
      for (const r of records.slice(0, 500)) { const v = r[field]; if (v === null) continue; vals++; if (typeof v === 'number') nums++; }
      return vals > 0 && nums / vals >= 0.6;
    } catch (e) { return false; }
  }

  function change(fn) {
    const p = currentPivot();
    if (!p) return;
    App().pushUndo([p.sheet]);
    fn(p);
    commit(p);
    openPanel(p);
  }

  function bindPanel(p, fields) {
    panel.querySelector('[data-close]').addEventListener('click', closePanel);
    panel.querySelector('.pv-search').addEventListener('input', e => {
      const q = e.target.value.trim().toLowerCase();
      panel.querySelectorAll('.pv-field').forEach(el => { el.style.display = !q || el.dataset.field.toLowerCase().includes(q) ? '' : 'none'; });
    });
    panel.querySelectorAll('[data-toggle]').forEach(cb => cb.addEventListener('change', () => {
      const field = cb.dataset.toggle;
      change(pv => {
        if (cb.checked) {
          const numeric = isNumericField(pv, field);
          addField(pv, numeric ? 'values' : 'rows', field, numeric);
        } else {
          ['rows', 'cols', 'filters', 'values'].forEach(z => removeField(pv, z, field));
          while (pv.values.some(v => v.field === field)) removeField(pv, 'values', field);
        }
      });
    }));
    panel.querySelectorAll('[data-remove]').forEach(btn => btn.addEventListener('click', () => {
      const chipEl = btn.closest('.pv-chip');
      change(pv => removeField(pv, chipEl.dataset.zone, chipEl.dataset.field));
    }));
    panel.querySelectorAll('.pv-agg').forEach(sel => sel.addEventListener('change', () => {
      const field = sel.dataset.field;
      change(pv => {
        const v = pv.values.find(x => x.field === field);
        if (v) { v.agg = sel.value; v.name = `${AGG_LABEL[sel.value]} de ${field}`; }
      });
    }));
    panel.querySelectorAll('[data-filter]').forEach(btn => btn.addEventListener('click', e => {
      e.stopPropagation();
      openItemFilter(btn, btn.dataset.filter, btn.closest('.pv-chip').dataset.zone);
    }));
    panel.querySelector('[data-refresh]').addEventListener('click', () => {
      const pv = currentPivot();
      if (!pv) return;
      App().pushUndo([pv.sheet]);
      commit(pv, 'Tabla dinámica actualizada');
      openPanel(pv);
    });
    panel.querySelector('[data-delete]').addEventListener('click', () => {
      const pv = currentPivot();
      if (!pv || !confirm(`¿Eliminar la tabla dinámica "${pv.name}"?`)) return;
      App().pushUndo([pv.sheet]);
      const a = areaOf(pv);
      const sheet = wb().sheets[pv.sheet];
      if (a && sheet) for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) delete sheet.cells[`${indexToCol(c)}${r}`];
      wb().pivots = pivots().filter(x => x.id !== pv.id);
      App().markDirty(pv.sheet);
      App().recalcAndRender();
      App().save();
      closePanel();
    });

    // drag & drop between the field list and the four zones
    let drag = null;
    panel.addEventListener('dragstart', e => {
      const el = e.target.closest('[draggable="true"]');
      if (!el) return;
      drag = { field: el.dataset.field, from: el.dataset.zone || null };
      e.dataTransfer.setData('text/plain', drag.field);
      e.dataTransfer.effectAllowed = 'move';
    });
    panel.querySelectorAll('.pv-zone').forEach(zoneEl => {
      zoneEl.addEventListener('dragover', e => { e.preventDefault(); zoneEl.classList.add('over'); });
      zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('over'));
      zoneEl.addEventListener('drop', e => {
        e.preventDefault();
        zoneEl.classList.remove('over');
        if (!drag) return;
        const to = zoneEl.dataset.zone;
        const { field, from } = drag;
        drag = null;
        if (from === to) return;
        change(pv => {
          if (from) removeField(pv, from, field);
          if (field === VALUES && to !== 'rows' && to !== 'cols') return;
          addField(pv, to, field, to === 'values' ? isNumericField(pv, field) : false);
        });
      });
    });
    panel.addEventListener('dragend', () => { drag = null; });
  }

  // checklist of items for a row/column/filter field
  let itemMenu = null;
  function closeItemFilter() { if (itemMenu) { itemMenu.remove(); itemMenu = null; } }
  function openItemFilter(anchor, field, zone) {
    closeItemFilter();
    const p = currentPivot();
    if (!p) return;
    let records;
    try { records = readSource(p).records; } catch (e) { return; }
    const map = new Map();
    for (const r of records) { const k = itemKey(r[field]); if (!map.has(k)) map.set(k, r[field]); }
    const items = [...map.values()].sort(compareItems);
    const hidden = new Set(((p.hidden || {})[field] || []).map(itemKey));
    const filterDef = zone === 'filters' ? p.filters.find(f => f.field === field) : null;
    const selected = filterDef && filterDef.selected ? new Set(filterDef.selected.map(itemKey)) : null;
    const isOn = v => !hidden.has(itemKey(v)) && (!selected || selected.has(itemKey(v)));
    itemMenu = document.createElement('div');
    itemMenu.className = 'af-menu pv-item-menu no-print';
    itemMenu.innerHTML = `
      <div class="pv-item-title">${escapeHtml(field)}</div>
      <input type="search" class="af-search" placeholder="Buscar" />
      <div class="af-list">
        <label class="af-item af-all"><input type="checkbox" data-all /> (Seleccionar todo)</label>
        ${items.map((v, i) => `<label class="af-item" data-text="${escapeHtml(String(itemLabel(v)).toLowerCase())}"><input type="checkbox" data-i="${i}" ${isOn(v) ? 'checked' : ''} /> ${escapeHtml(itemLabel(v))}</label>`).join('')}
      </div>
      <div class="af-footer">
        <button type="button" class="af-ok" data-ok>Aceptar</button>
        <button type="button" class="af-cancel" data-cancel>Cancelar</button>
      </div>`;
    document.body.appendChild(itemMenu);
    const rect = anchor.getBoundingClientRect();
    itemMenu.style.left = `${Math.max(8, Math.min(window.innerWidth - 270, rect.left - 230))}px`;
    itemMenu.style.top = `${Math.max(8, Math.min(window.innerHeight - 420, rect.bottom + 4))}px`;
    const boxes = () => [...itemMenu.querySelectorAll('[data-i]')];
    const all = itemMenu.querySelector('[data-all]');
    const sync = () => { all.checked = boxes().every(b => b.checked); all.indeterminate = !all.checked && boxes().some(b => b.checked); };
    sync();
    all.addEventListener('change', () => { boxes().forEach(b => { if (b.closest('.af-item').style.display !== 'none') b.checked = all.checked; }); sync(); });
    itemMenu.querySelector('.af-list').addEventListener('change', e => { if (!e.target.matches('[data-all]')) sync(); });
    itemMenu.querySelector('.af-search').addEventListener('input', e => {
      const q = e.target.value.trim().toLowerCase();
      itemMenu.querySelectorAll('.af-item:not(.af-all)').forEach(el => { el.style.display = !q || el.dataset.text.includes(q) ? '' : 'none'; });
    });
    itemMenu.querySelector('[data-cancel]').addEventListener('click', closeItemFilter);
    itemMenu.querySelector('[data-ok]').addEventListener('click', () => {
      const on = boxes().filter(b => b.checked).map(b => items[Number(b.dataset.i)]);
      const off = boxes().filter(b => !b.checked).map(b => items[Number(b.dataset.i)]);
      closeItemFilter();
      change(pv => {
        pv.hidden = pv.hidden || {};
        if (zone === 'filters') {
          const f = pv.filters.find(x => x.field === field);
          delete pv.hidden[field];
          if (f) f.selected = off.length ? on : null;
        } else if (off.length) pv.hidden[field] = off;
        else delete pv.hidden[field];
      });
    });
  }

  document.addEventListener('mousedown', e => {
    if (itemMenu && !itemMenu.contains(e.target) && !e.target.closest('[data-filter]')) closeItemFilter();
  });

  // ---------------------------------------------------------------- hooks
  function onSelectionChange(sheetName, coord) {
    const p = pivotAt(sheetName, coord);
    if (p) { if (panelPivotId !== p.id) openPanel(p); }
    else if (panel) closePanel();
  }

  function onSheetRenamed(oldName, newName) {
    for (const p of pivots()) {
      if (p.sheet === oldName) p.sheet = newName;
      if (p.source && p.source.sheet === oldName) p.source.sheet = newName;
    }
  }

  function onSheetDeleted(name) {
    wb().pivots = pivots().filter(p => p.sheet !== name);
    if (panel && !currentPivot()) closePanel();
  }

  document.getElementById('btnInsertPivot')?.addEventListener('click', openCreateDialog);
  document.getElementById('btnRefreshPivots')?.addEventListener('click', () => {
    const sheetName = App().getActiveSheet();
    const list = pivots().filter(p => p.sheet === sheetName);
    if (!list.length) return App().toast('Esta hoja no tiene tablas dinámicas', true);
    App().pushUndo([sheetName]);
    let ok = 0;
    for (const p of list) { try { renderPivot(p); ok++; } catch (e) { App().toast(e.message, true); } }
    App().markDirty(sheetName);
    App().recalcAndRender();
    App().save();
    App().toast(`${ok} tabla(s) dinámica(s) actualizada(s)`);
    const pv = currentPivot();
    if (pv) openPanel(pv);
  });

  window.MauriPivot = { onSelectionChange, onSheetRenamed, onSheetDeleted, renderPivot, computeGrid, pivotAt, openCreateDialog };
})();
