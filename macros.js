// ============================================================================
// MACROS - the workbook's VBA macros rewritten as app functions
//   1) Actualizar_ColumnaT_PegarData      (Module1)
//   2) ConsolidarDatos_Rapido_Optimizado  (ThisWorkbook)
// ============================================================================

(function () {
  const { indexToCol, parseCoord } = window.ExcelUtils;
  const App = () => window.MauriApp;
  const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

  // ---------------------------------------------------------------- helpers
  const sheetOf = name => App().getWorkbook().sheets[name];

  function cell(sheet, coord) { return sheet.cells[coord]; }
  function val(sheet, coord) {
    const c = sheet.cells[coord];
    return c && c.v !== undefined && c.v !== null ? c.v : '';
  }
  function isEmpty(v) { return v === '' || v === null || v === undefined; }

  function setValue(sheet, coord, value, styleFrom) {
    const prev = sheet.cells[coord];
    const next = { v: value };
    if (prev && prev.s) next.s = prev.s;
    else if (styleFrom && styleFrom.s) next.s = JSON.parse(JSON.stringify(styleFrom.s));
    if (typeof value === 'string' && /^#(DIV\/0!|VALUE!|REF!|N\/A|NAME\?|NUM!)$/.test(value)) next.e = true;
    sheet.cells[coord] = next;
  }

  function clearContents(sheet, coord) {
    const c = sheet.cells[coord];
    if (!c) return;
    if (c.s) sheet.cells[coord] = { s: c.s };
    else delete sheet.cells[coord];
  }

  function lastRowWithValue(sheet, colStr) {
    let last = 0;
    for (const [coord, c] of Object.entries(sheet.cells)) {
      if (!c || isEmpty(c.v)) continue;
      const p = parseCoord(coord);
      if (p && p.colStr === colStr && p.row > last) last = p.row;
    }
    return last;
  }

  // Cells(row, Columns.Count).End(xlToLeft).Column — a merged block counts up to its right edge
  function lastColWithValue(sheet, row) {
    let last = 0;
    for (const [coord, c] of Object.entries(sheet.cells)) {
      if (!c || (isEmpty(c.v) && !c.f)) continue;
      const p = parseCoord(coord);
      if (p && p.row === row && p.col > last) last = p.col;
    }
    for (const m of sheet.merges || []) {
      const [a, b] = m.split(':').map(parseCoord);
      if (a && b && a.row <= row && b.row >= row && a.col <= last && b.col >= last) last = Math.max(last, b.col);
    }
    return last;
  }

  // AutoFilter criterion "Marigold": the cell text equals it (case-insensitive)
  const textEquals = (sheetName, coord, word) =>
    App().cellDisplayText(sheetName, coord).toLowerCase() === String(word).toLowerCase();

  // Excel-style MsgBox
  function msgBox(text, title = 'Reunion 123', icon = 'info') {
    return new Promise(resolve => {
      const wrap = document.createElement('div');
      wrap.className = 'msgbox-overlay no-print';
      wrap.innerHTML = `
        <div class="msgbox" role="alertdialog" aria-modal="true">
          <div class="msgbox-title">${escapeHtml(title)}</div>
          <div class="msgbox-body"><span class="msgbox-icon ${icon}">${icon === 'warn' ? '⚠' : 'ℹ'}</span><span>${escapeHtml(text)}</span></div>
          <div class="msgbox-actions"><button type="button" class="msgbox-ok">Aceptar</button></div>
        </div>`;
      document.body.appendChild(wrap);
      document.body.classList.add('modal-open');
      const done = () => { wrap.remove(); document.body.classList.remove('modal-open'); resolve(); };
      wrap.querySelector('.msgbox-ok').addEventListener('click', done);
      wrap.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === 'Escape') done(); });
      wrap.querySelector('.msgbox-ok').focus();
    });
  }

  function escapeHtml(s) {
    return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function busy(on) {
    document.body.classList.toggle('macro-running', on);
  }

  // ============================================================================
  // 1) Actualizar_ColumnaT_PegarData
  // ============================================================================
  const MATRICARIA_5 = new Set([
    'STATER BROS', 'GALLERIA #3758', 'NOR-CAL PRODUCE INC.', 'MAYESH',
    'GALLERIA #710269', 'SAFEWAY NORCAL', "GELSON'S MARKETS /  11703", "CHARLIE'S PRODUCE SEATTLE"
  ]);

  async function actualizarColumnaT() {
    const NAME = 'PegarData';
    const ws = sheetOf(NAME);
    if (!ws) return msgBox('No existe la hoja PegarData.', 'Error', 'warn');

    // === VALIDACIÓN DE DATOS === (COUNTA(E:E) = 0)
    const lastRow = lastRowWithValue(ws, 'E');
    if (lastRow === 0) return msgBox('No hay datos para procesar en la hoja PegarData.', 'Reunion 123', 'warn');

    busy(true);
    try {
      App().pushUndo([NAME]);
      const T = r => `T${r}`;

      // === PASO 1: T = N / M (as values) ===
      if (lastRow >= 2) {
        const engine = App().getEngine();
        for (let r = 2; r <= lastRow; r++) {
          const n = engine.evaluateFormula(`=N${r}/M${r}`, NAME, T(r));
          setValue(ws, T(r), n);
        }
      }

      // Filtered assignment on column T (20) — like AsignarValorFiltrado
      const lastT = () => lastRowWithValue(ws, 'T');
      const tPasses = (r, criteria) => criteria.every(c => {
        const m = /^<>(.+)$/.exec(c);
        if (!m) return true;
        const v = val(ws, T(r));
        const n = Number(m[1]);
        return typeof v === 'number' ? v !== n : String(v).toLowerCase() !== m[1].toLowerCase();
      });
      const asignar = (flor, nuevoValor, ...criteria) => {
        const end = lastT();
        for (let r = 2; r <= end; r++) {
          if (!textEquals(NAME, `E${r}`, flor)) continue;
          if (!tPasses(r, criteria)) continue;
          setValue(ws, T(r), nuevoValor);
        }
      };

      // === PASO 2: MARIGOLD (T = 10 si <>5 y <>10) ===
      asignar('Marigold', 10, '<>5', '<>10');
      // === PASO 3: VERONICA (T = 10 si <>8) ===
      asignar('Veronica', 10, '<>8');
      // === PASO 4: SNAPDRAGON (T = 5 si <>5 y <>10) ===
      asignar('Snapdragon', 5, '<>5', '<>10');

      // === PASO 5: MATRICARIA (valor según cliente) ===
      const endE = lastRowWithValue(ws, 'E');
      for (let r = 2; r <= endE; r++) {
        if (!textEquals(NAME, `E${r}`, 'Matricaria')) continue;
        const cliente = String(val(ws, `B${r}`)).trim().toUpperCase();
        let v = 6;
        if (cliente === 'GALLERIA #710332') v = 4;
        else if (MATRICARIA_5.has(cliente)) v = 5;
        setValue(ws, T(r), v);
      }

      App().markDirty(NAME);
      App().recalcAndRender();
      App().save();
      busy(false);
      await msgBox('Proceso completado correctamente.', 'Listo');
    } catch (err) {
      busy(false);
      await msgBox(`Ocurrió un error: ${err.message}`, 'Error', 'warn');
    }
  }

  // ============================================================================
  // 2) ConsolidarDatos_Rapido_Optimizado
  // ============================================================================
  async function whoAmI() {
    try {
      const res = await fetch('/api/whoami');
      const data = await res.json();
      return data.user || '';
    } catch (e) {
      return '';
    }
  }

  async function consolidarDatos() {
    const t0 = performance.now();
    const wb = App().getWorkbook();
    const ws = sheetOf('Cortes');
    if (!ws) return msgBox('No existe la hoja Cortes.', 'Error', 'warn');

    busy(true);
    try {
      const userName = await whoAmI();
      // === Hoja origen === lastRow by column A, lastCol from row 1 (End(xlToLeft))
      const lastRow = lastRowWithValue(ws, 'A');
      const lastCol = lastColWithValue(ws, 1);
      const value = (r, c) => {
        const v = App().getEngine().getCellValue('Cortes', `${indexToCol(c)}${r}`);
        return v === null || v === undefined ? '' : v;
      };

      // === Procesar todo en memoria ===
      const out = [];
      for (let i = 2; i <= lastRow; i++) {
        for (let j = 2; j <= lastCol; j += 3) {
          if (j + 2 > lastCol) continue;
          const a = ws.cells[`${indexToCol(j)}${i}`], b = ws.cells[`${indexToCol(j + 1)}${i}`], c = ws.cells[`${indexToCol(j + 2)}${i}`];
          const empty = x => !x || (x.f === undefined && isEmpty(x.v));
          if (empty(a) && empty(b) && empty(c)) continue;
          out.push([value(i, 1), value(i, j), value(i, j + 1), value(i, j + 2)]);
        }
      }

      // === Hoja de salida ===
      let wsOut = sheetOf('Flattened');
      const touched = ['Flattened', 'Bitacora'];
      App().pushUndo(touched);
      if (!wsOut) {
        wb.sheets.Flattened = { maxRow: 100, maxCol: 20, colWidths: {}, merges: [], cells: {} };
        wb.sheetOrder.push('Flattened');
        wsOut = wb.sheets.Flattened;
        ['Flor', 'Fecha', 'Corte', 'Botes'].forEach((h, k) => setValue(wsOut, `${indexToCol(k + 1)}1`, h));
      }
      const dateStyle = wsOut.cells.B12 || wsOut.cells.B3 || null;
      out.forEach((rowVals, k) => {
        const r = k + 2;
        rowVals.forEach((v, c) => {
          const coord = `${indexToCol(c + 1)}${r}`;
          if (v === '') clearContents(wsOut, coord); // VBA writes Empty
          else setValue(wsOut, coord, v, c === 1 ? dateStyle : null);
        });
      });
      // Limpiar desde la fila siguiente hasta el final
      const lastOut = Math.max(lastRowWithValue(wsOut, 'A'), lastRowWithValue(wsOut, 'B'), lastRowWithValue(wsOut, 'C'), lastRowWithValue(wsOut, 'D'));
      for (let r = out.length + 2; r <= lastOut; r++) {
        for (let c = 1; c <= 4; c++) clearContents(wsOut, `${indexToCol(c)}${r}`);
      }
      wsOut.maxRow = Math.max(wsOut.maxRow || 100, out.length + 21);

      // === Bitácora ===
      let wsLog = sheetOf('Bitacora');
      if (!wsLog) {
        wb.sheets.Bitacora = { maxRow: 100, maxCol: 20, colWidths: {}, merges: [], cells: {} };
        wb.sheetOrder.push('Bitacora');
        wsLog = wb.sheets.Bitacora;
        ['Fecha/Hora', 'Usuario', 'Archivo', 'Acción'].forEach((h, k) => setValue(wsLog, `${indexToCol(k + 1)}1`, h));
      }
      const now = new Date();
      const serial = (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds()) - EXCEL_EPOCH_MS) / 86400000;
      const logRow = lastRowWithValue(wsLog, 'A') + 1;
      const dateFrom = wsLog.cells.A2 && wsLog.cells.A2.s ? wsLog.cells.A2 : { s: { nf: 'm/d/yy h:mm' } };
      setValue(wsLog, `A${logRow}`, Math.round(serial * 86400) / 86400, dateFrom);
      setValue(wsLog, `B${logRow}`, userName);
      setValue(wsLog, `C${logRow}`, wb.workbookName || 'Reunion 123');
      setValue(wsLog, `D${logRow}`, 'Ejecución de macro ConsolidarDatos_Rapido');
      wsLog.maxRow = Math.max(wsLog.maxRow || 100, logRow + 20);

      touched.forEach(n => App().markDirty(n));
      // === Activar hoja "Reunion" si existe ===
      App().recalcAndRender();
      if (sheetOf('Reunion') && !sheetOf('Reunion').hidden) App().switchSheet('Reunion');
      App().save();
      busy(false);
      const secs = ((performance.now() - t0) / 1000).toFixed(2);
      await msgBox(`Datos consolidados y bitácora actualizada. (${secs} seg)`, 'Proceso completado');
    } catch (err) {
      busy(false);
      await msgBox(`Ocurrió un error: ${err.message}`, 'Error', 'warn');
    }
  }

  window.MauriMacros = { actualizarColumnaT, consolidarDatos };
})();
