// ============================================================================
// APP CONTROLLER - SISTEMA EXCEL REUNION 123
// ============================================================================

(function () {
  const { colToIndex, indexToCol, parseCoord } = window.ExcelUtils;

  let workbook = null;
  let engine = null;
  let activeSheet = 'Sheet1';
  let selectedCoord = 'A1';
  let selectionEndCoord = 'A1';
  let isDraggingSelection = false;
  let isDraggingFillHandle = false;
  let fillPreviewTargetCoord = null;
  let clipboardState = null; // { sheetName, minRow, maxRow, minCol, maxCol, isCut, grid, tsv }
  let inlineEditor = null;
  let zoomLevel = 100;
  let formatPainterActive = false;
  let formatPainterSource = null;

  // Undo / Redo history
  const undoStack = [];
  const redoStack = [];
  const MAX_HISTORY = 40;

  // DOM References
  const gridHead = document.getElementById('excelGridHead');
  const gridBody = document.getElementById('excelGridBody');
  const sheetCanvasWrapper = document.getElementById('sheetCanvasWrapper');
  const cellNameBox = document.getElementById('cellNameBox');
  const formulaInput = document.getElementById('formulaInput');
  const cellMetaBadge = document.getElementById('cellMetaBadge');
  const sheetTabsList = document.getElementById('sheetTabsList');

  // Format toolbar DOM references
  const btnPaste = document.getElementById('btnPaste');
  const btnCut = document.getElementById('btnCut');
  const btnCopy = document.getElementById('btnCopy');
  const fontFamilySelect = document.getElementById('fontFamilySelect');
  const fontSizeSelect = document.getElementById('fontSizeSelect');
  const btnGrowFont = document.getElementById('btnGrowFont');
  const btnShrinkFont = document.getElementById('btnShrinkFont');
  const btnBold = document.getElementById('btnBold');
  const btnItalic = document.getElementById('btnItalic');
  const btnUnderline = document.getElementById('btnUnderline');
  const btnStrike = document.getElementById('btnStrike');
  const btnFontColor = document.getElementById('btnFontColor');
  const btnFontColorMenu = document.getElementById('btnFontColorMenu');
  const fontColorMenu = document.getElementById('fontColorMenu');
  const fontColorSwatch = document.getElementById('fontColorSwatch');
  const btnFillColor = document.getElementById('btnFillColor');
  const btnFillColorMenu = document.getElementById('btnFillColorMenu');
  const fillColorMenu = document.getElementById('fillColorMenu');
  const fillColorSwatch = document.getElementById('fillColorSwatch');
  let lastFillColor = '#ffff00';
  let lastFontColor = '#c00000';
  const btnValignTop = document.getElementById('btnValignTop');
  const btnValignMiddle = document.getElementById('btnValignMiddle');
  const btnValignBottom = document.getElementById('btnValignBottom');
  const btnAlignLeft = document.getElementById('btnAlignLeft');
  const btnAlignCenter = document.getElementById('btnAlignCenter');
  const btnAlignRight = document.getElementById('btnAlignRight');
  const btnWrapText = document.getElementById('btnWrapText');
  const btnMergeCenter = document.getElementById('btnMergeCenter');
  const btnBorders = document.getElementById('btnBorders');
  const borderMenu = document.getElementById('borderMenu');
  const btnNumFmt = document.getElementById('btnNumFmt');
  const numFmtLabel = document.getElementById('numFmtLabel');
  const numFmtMenu = document.getElementById('numFmtMenu');
  const btnCurrencyFmt = document.getElementById('btnCurrencyFmt');
  const btnPercentFmt = document.getElementById('btnPercentFmt');
  const btnCommaFmt = document.getElementById('btnCommaFmt');
  const btnDecimalInc = document.getElementById('btnDecimalInc');
  const btnDecimalDec = document.getElementById('btnDecimalDec');
  const btnFormatPainter = document.getElementById('btnFormatPainter');
  const btnClearFormat = document.getElementById('btnClearFormat');

  // Which "apartado" (workbook) this page shows. The three workbooks now stay
  // in the same browser session so changing the top tabs does not reload the
  // complete Excel engine every time.
  const VALID_WB_KEYS = ['reunion', 'cosecha', 'rvp'];
  const wbFromUrl = () => {
    const value = new URLSearchParams(location.search).get('wb');
    return VALID_WB_KEYS.includes(value) ? value : 'reunion';
  };
  let WB_KEY = wbFromUrl();
  let API_Q = `?wb=${WB_KEY}`;
  let BACKUP_KEY = WB_KEY === 'reunion' ? 'reunion123_workbook_backup' : `${WB_KEY}_workbook_backup`;
  const workbookCache = new Map();
  const workbookEngineCache = new Map();
  const workbookLoading = new Map();
  // Last workbook state acknowledged by the server. This lets autosave send
  // cell-level patches instead of serializing the entire sheet on every edit.
  const syncedSheetSnapshots = new Map();
  const CACHE_DB_NAME = 'reunion123-workbook-cache';
  const CACHE_STORE = 'workbooks';

  function openWorkbookCache() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(CACHE_DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(CACHE_STORE)) db.createObjectStore(CACHE_STORE);
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function readCachedWorkbook(key) {
    try {
      const db = await openWorkbookCache();
      return await new Promise((resolve, reject) => {
        const request = db.transaction(CACHE_STORE, 'readonly').objectStore(CACHE_STORE).get(key);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
      });
    } catch (error) { return null; }
  }

  async function writeCachedWorkbook(key, value) {
    try {
      const db = await openWorkbookCache();
      await new Promise((resolve, reject) => {
        const request = db.transaction(CACHE_STORE, 'readwrite').objectStore(CACHE_STORE).put(value, key);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    } catch (error) { /* Cache is an acceleration only; Neon remains authoritative. */ }
  }

  function cloneJson(value) {
    return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
  }

  function rememberSyncedWorkbook(key, nextWorkbook) {
    const snapshots = new Map();
    for (const [name, sheet] of Object.entries(nextWorkbook?.sheets || {})) {
      snapshots.set(name, cloneJson(sheet));
    }
    syncedSheetSnapshots.set(key, snapshots);
  }

  function rememberSyncedPatch(key, payload) {
    const snapshots = syncedSheetSnapshots.get(key);
    if (!snapshots) return;
    for (const [name, patch] of Object.entries(payload.cellPatches || {})) {
      const sheet = snapshots.get(name);
      if (!sheet) continue;
      if (!sheet.cells) sheet.cells = {};
      for (const [coord, cell] of Object.entries((patch && patch.set) || {})) sheet.cells[coord] = cloneJson(cell);
      for (const coord of ((patch && patch.delete) || [])) delete sheet.cells[coord];
    }
  }

  function setWorkbookKey(key) {
    WB_KEY = VALID_WB_KEYS.includes(key) ? key : 'reunion';
    API_Q = `?wb=${WB_KEY}`;
    BACKUP_KEY = WB_KEY === 'reunion' ? 'reunion123_workbook_backup' : `${WB_KEY}_workbook_backup`;
  }

  function initialWorkbookPath(key) {
    if (key === 'cosecha') return 'data/cosecha_initial_workbook.json';
    if (key === 'rvp') return 'data/rvp_initial_workbook.json';
    return 'data/initial_workbook.json';
  }

  async function fetchWithTimeout(url, options = {}, timeoutMs = 15000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await fetch(url, { ...options, signal: controller.signal });
    } finally {
      clearTimeout(timer);
    }
  }

  async function savePatchWithRetry(url, options) {
    let lastError = null;
    for (let attempt = 0; attempt < 5; attempt++) {
      try {
        const response = await fetchWithTimeout(url, options, 8000);
        if (response.ok) return response;
        lastError = new Error(String(response.status));
      } catch (error) {
        lastError = error;
      }
      if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 300 * (attempt + 1)));
    }
    throw lastError || new Error('No se pudo guardar');
  }

  async function fetchWorkbookForKey(key) {
    // Neon is the only source of truth in the deployed app. Do not silently
    // replace a failed database response with a bundled JSON snapshot: that
    // made the interface look healthy while showing stale data.
    let lastError = null;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const response = await fetchWithTimeout(`/api/workbook?wb=${key}&index=1`, { cache: 'no-store' });
        if (response.ok) {
          const index = await response.json();
          const visible = (index.sheetOrder || Object.keys(index.sheets || {})).filter(name => !index.sheets?.[name]?.hidden);
          const initialName = index.activeSheet && index.sheets?.[index.activeSheet] && !index.sheets[index.activeSheet].hidden
            ? index.activeSheet : (visible[0] || Object.keys(index.sheets || {})[0]);
          if (initialName) {
            const sheetResponse = await fetchWithTimeout(`/api/sheet?wb=${key}&name=${encodeURIComponent(initialName)}`, { cache: 'no-store' });
            if (!sheetResponse.ok) throw new Error(`No se pudo cargar la hoja ${initialName} desde Neon (${sheetResponse.status})`);
            index.sheets[initialName] = await sheetResponse.json();
          }
          return index;
        }
        const detail = await response.text().catch(() => '');
        lastError = new Error(`No se pudo cargar ${key} desde Neon (${response.status}${detail ? `: ${detail.slice(0, 160)}` : ''})`);
        // Cloudflare can return a transient 502/503/504 while the large
        // workbook response is being assembled. Retry Neon, never a snapshot.
        if (![502, 503, 504].includes(response.status)) break;
      } catch (error) {
        lastError = error;
        if (attempt === 2) throw error;
      }
      await new Promise(resolve => setTimeout(resolve, 700 * (attempt + 1)));
    }
    throw lastError || new Error(`No se pudo cargar ${key} desde Neon`);
  }

  async function loadSheetFromNeon(key, sheetName) {
    const response = await fetchWithTimeout(`/api/sheet?wb=${key}&name=${encodeURIComponent(sheetName)}`, { cache: 'no-store' });
    if (!response.ok) throw new Error(`No se pudo cargar la hoja ${sheetName} desde Neon (${response.status})`);
    return response.json();
  }

  function workbookReady(nextWorkbook, key) {
    const ready = nextWorkbook || { workbookName: key, sheetOrder: [], sheets: {} };
    if (!ready.sheetOrder || ready.sheetOrder.length === 0) {
      ready.sheetOrder = Object.keys(ready.sheets || {});
    }
    return ready;
  }

  function prefetchWorkbook(key) {
    if (key === WB_KEY || workbookCache.has(key)) return Promise.resolve(workbookCache.get(key));
    if (workbookLoading.has(key)) return workbookLoading.get(key);
    const request = fetchWorkbookForKey(key)
      .then(next => {
        const ready = workbookReady(next, key);
        // Recalculate in the background so the click only has to swap the
        // already-loaded workbook and paint its current sheet.
        const warmEngine = new window.FormulaEngine(ready);
        warmEngine.recalculateAll();
        workbookCache.set(key, ready);
        workbookEngineCache.set(key, warmEngine);
        rememberSyncedWorkbook(key, ready);
        return ready;
      })
      .catch(() => null)
      .finally(() => workbookLoading.delete(key));
    workbookLoading.set(key, request);
    return request;
  }

  function activateWorkbook(nextWorkbook, key, options = {}) {
    setWorkbookKey(key);
    workbook = workbookReady(nextWorkbook, key);
    workbookCache.set(WB_KEY, workbook);
    if (!syncedSheetSnapshots.has(WB_KEY)) rememberSyncedWorkbook(WB_KEY, workbook);
    if (workbook.workbookName) document.title = workbook.workbookName;

    const statusInfo = document.getElementById('statusFormulaInfo');
    if (statusInfo) statusInfo.textContent = `${workbook.workbookName || 'Reunion 123'} • Excel Engine Active`;

    const firstVisible = visibleSheets()[0] || workbook.sheetOrder[0] || 'Sheet1';
    activeSheet = workbook.activeSheet && workbook.sheets[workbook.activeSheet] && !workbook.sheets[workbook.activeSheet].hidden
      ? workbook.activeSheet : firstVisible;
    const cachedEngine = workbookEngineCache.get(WB_KEY);
    engine = cachedEngine || new window.FormulaEngine(workbook);
    if (!cachedEngine) {
      const recalc = () => {
        // The first paint must not wait for every formula in every sheet.
        // Visible formulas are evaluated lazily by the grid; the complete
        // pass runs only after the browser is idle.
        if (engine !== engineRef || workbook !== engineRef.workbook) return;
        engineRef.recalculateAll();
        if (WB_KEY === key) renderSheet();
      };
      const engineRef = engine;
      if (options.deferRecalc) {
        if (typeof window.requestIdleCallback === 'function') window.requestIdleCallback(recalc, { timeout: 1500 });
        else setTimeout(recalc, 0);
      } else {
        recalc();
      }
    }
    renderSheetTabs();
    renderSheet();
    selectCell('A1', false);
    sheetTabsList.querySelector('.sheet-tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });

    document.querySelectorAll('.wb-switch [data-wb]').forEach(a => a.classList.toggle('active', a.dataset.wb === WB_KEY));
    const btnActualizarData = document.getElementById('btnActualizarData');
    if (btnActualizarData) btnActualizarData.hidden = WB_KEY !== 'reunion';
    const btnActualizarInventario = document.getElementById('btnActualizarInventario');
    if (btnActualizarInventario) btnActualizarInventario.hidden = WB_KEY !== 'reunion';
    updateStatusMetrics();
    window.dispatchEvent(new CustomEvent('mauri:workbook-switch', { detail: { key: WB_KEY } }));
  }

  let workbookEventSource = null;
  function applyIncomingCellPatch(target, payload) {
    if (!target || !payload) return false;
    for (const [name, patch] of Object.entries(payload.cellPatches || {})) {
      const sheet = target.sheets && target.sheets[name];
      if (!sheet) continue;
      if (!sheet.cells) sheet.cells = {};
      for (const [coord, cell] of Object.entries((patch && patch.set) || {})) sheet.cells[coord] = cell;
      for (const coord of ((patch && patch.delete) || [])) delete sheet.cells[coord];
    }
    for (const [key, value] of Object.entries(payload.meta || {})) target[key] = value;
    return true;
  }

  function listenWorkbookChanges() {
    if (workbookEventSource) workbookEventSource.close();
    workbookEventSource = new EventSource('/api/events');
    workbookEventSource.addEventListener('workbook-patch', event => {
      try {
        const payload = JSON.parse(event.data);
        const key = payload.key;
        const target = key === WB_KEY ? workbook : workbookCache.get(key);
        if (!target || !applyIncomingCellPatch(target, payload)) return;
        workbookCache.set(key, target);
        void writeCachedWorkbook(key, target);
        if (key === WB_KEY) {
          rememberSyncedPatch(key, payload);
          engine.recalculateAll();
          renderSheet();
          selectCell(selectedCoord, false);
        }
      } catch (error) { console.error('No se pudo aplicar cambio en vivo:', error); }
    });
    workbookEventSource.addEventListener('workbook-refresh', event => {
      try {
        const payload = JSON.parse(event.data);
        const key = payload.key;
        void fetchWorkbookForKey(key).then(fresh => {
          const ready = workbookReady(fresh, key);
          workbookCache.set(key, ready);
          rememberSyncedWorkbook(key, ready);
          if (key === WB_KEY && !dirtySheets.size) activateWorkbook(ready, key, { deferRecalc: true });
        }).catch(() => {});
      } catch (error) { /* reconnect/background refresh will recover */ }
    });
  }
  let actualizarDataPollTimer = null;
  let actualizarDataPollInFlight = false;
  const actualizarDataNoticeKey = 'reunion123_actualizar_data_notice';
  let actualizarInventarioPollTimer = null;
  let actualizarInventarioPollInFlight = false;
  const actualizarInventarioNoticeKey = 'reunion123_actualizar_inventario_notice';

  async function init() {
    try {
      const loaded = await fetchWorkbookForKey(WB_KEY);
      activateWorkbook(loaded, WB_KEY, { deferRecalc: true });
    } catch (err) {
      // Keep the UI explicit about a database outage instead of opening a
      // stale local workbook that could be mistaken for current Neon data.
      activateWorkbook({ workbookName: WB_KEY, sheetOrder: [], sheets: {} }, WB_KEY);
      showToast(err.message, true);
      // A transient 503 can happen while Cloudflare/Neon establishes a
      // connection. Keep retrying Neon in the background and recover the
      // workbook without requiring another manual refresh.
      const failedKey = WB_KEY;
      void (async () => {
        for (let attempt = 0; attempt < 6; attempt++) {
          await new Promise(resolve => setTimeout(resolve, 1500 * (attempt + 1)));
          if (WB_KEY !== failedKey || workbook?.sheetOrder?.length) return;
          try {
            const recovered = await fetchWorkbookForKey(failedKey);
            if (WB_KEY !== failedKey) return;
            activateWorkbook(recovered, failedKey, { deferRecalc: true });
            showToast('Libro recuperado desde Neon.', false);
            return;
          } catch { /* keep retrying Neon; never use a local snapshot */ }
        }
      })();
    }

    sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
    const zLbl = document.getElementById('zoomLabel');
    if (zLbl) zLbl.textContent = `${zoomLevel}%`;

    bindEvents();
    bindWorkspaceNavigation();
    listenWorkbookChanges();
    pollActualizarData();
    pollActualizarInventario();

    // Other apartados are warmed only when the user points to or focuses a
    // tab. Never download/recalculate the other large workbooks during the
    // first paint of the current one.
  }

  // Undo history stores only the sheets an action touches (big workbooks would not fit otherwise)
  const UNDO_BYTES_BUDGET = 120 * 1024 * 1024;
  function captureState(names) {
    const saved = {};
    let bytes = 0;
    for (const n of new Set(names)) {
      const json = workbook.sheets[n] ? JSON.stringify(workbook.sheets[n]) : null;
      saved[n] = json;
      bytes += json ? json.length : 0;
    }
    const pivotsJson = JSON.stringify(workbook.pivots || []);
    return { order: [...(workbook.sheetOrder || [])], keys: Object.keys(workbook.sheets), saved, bytes: bytes + pivotsJson.length, pivots: pivotsJson };
  }

  function restoreState(snap) {
    const keys = new Set(snap.keys);
    for (const n of Object.keys(workbook.sheets)) {
      if (!keys.has(n)) { delete workbook.sheets[n]; markSheetDirty(n); }
    }
    for (const [n, json] of Object.entries(snap.saved)) {
      if (json === null) delete workbook.sheets[n];
      else workbook.sheets[n] = JSON.parse(json);
      markSheetDirty(n);
    }
    workbook.sheetOrder = [...snap.order];
    if (snap.pivots !== undefined) workbook.pivots = JSON.parse(snap.pivots);
  }

  function trimHistory(stack) {
    let total = stack.reduce((sum, s) => sum + s.bytes, 0);
    while (stack.length > MAX_HISTORY || (stack.length > 1 && total > UNDO_BYTES_BUDGET)) total -= stack.shift().bytes;
  }

  function pushUndoState(sheetNames = [activeSheet]) {
    const snap = captureState(sheetNames);
    const last = undoStack[undoStack.length - 1];
    if (last && JSON.stringify(last.saved) === JSON.stringify(snap.saved) && last.order.join('|') === snap.order.join('|') && snap.bytes < 2e6) return;
    undoStack.push(snap);
    trimHistory(undoStack);
    redoStack.length = 0;
    updateUndoRedoButtons();
  }

  function mirrorOf(snap) {
    const current = new Set(Object.keys(workbook.sheets));
    const names = new Set(Object.keys(snap.saved));
    for (const n of current) if (!snap.keys.includes(n)) names.add(n);
    return captureState([...names]);
  }

  function updateUndoRedoButtons() {
    document.getElementById('btnUndo').disabled = undoStack.length === 0;
    document.getElementById('btnRedo').disabled = redoStack.length === 0;
  }

  function undo() {
    if (undoStack.length === 0) return;
    const prev = undoStack.pop();
    redoStack.push(mirrorOf(prev));
    trimHistory(redoStack);
    restoreState(prev);
    if (!workbook.sheets[activeSheet] || workbook.sheets[activeSheet].hidden) activeSheet = visibleSheets()[0];
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    updateUndoRedoButtons();
    renderSheetTabs();
    renderSheet();
    selectCell(selectedCoord, false);
    saveWorkbookState();
    showToast('Action undone (Ctrl+Z)');
  }

  function redo() {
    if (redoStack.length === 0) return;
    const next = redoStack.pop();
    undoStack.push(mirrorOf(next));
    trimHistory(undoStack);
    restoreState(next);
    if (!workbook.sheets[activeSheet] || workbook.sheets[activeSheet].hidden) activeSheet = visibleSheets()[0];
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    updateUndoRedoButtons();
    renderSheetTabs();
    renderSheet();
    selectCell(selectedCoord, false);
    saveWorkbookState();
    showToast('Action redone (Ctrl+Y)');
  }

  // Saving sends only changed cells (PATCH); structural sheet changes fall back
  // to sending that complete sheet for compatibility.
  const dirtySheets = new Set();
  const dirtyCellCoords = new Map();
  function markSheetDirty(name) { if (name) dirtySheets.add(name); }
  function markCellDirty(name, coord) {
    if (!name || !coord) return;
    markSheetDirty(name);
    if (!dirtyCellCoords.has(name)) dirtyCellCoords.set(name, new Set());
    dirtyCellCoords.get(name).add(coord);
  }

  function buildSheetChange(current, previous) {
    if (!current || !previous) return { full: cloneJson(current) };

    const currentKeys = Object.keys(current);
    const previousKeys = Object.keys(previous);
    const structuralKeys = new Set([...currentKeys, ...previousKeys]);
    structuralKeys.delete('cells');
    for (const key of structuralKeys) {
      if (JSON.stringify(current[key]) !== JSON.stringify(previous[key])) {
        return { full: cloneJson(current) };
      }
    }

    const set = {};
    const deleted = [];
    const currentCells = current.cells || {};
    const previousCells = previous.cells || {};
    for (const coord of new Set([...Object.keys(currentCells), ...Object.keys(previousCells)])) {
      const now = currentCells[coord];
      const before = previousCells[coord];
      if (now === undefined) deleted.push(coord);
      else if (before === undefined || JSON.stringify(now) !== JSON.stringify(before)) set[coord] = cloneJson(now);
    }
    return { cells: { set, delete: deleted } };
  }

  let saveTimeout = null;
  let saving = Promise.resolve();
  function queueWorkbookSave() {
    saveTimeout = null;
    const names = [...dirtySheets];
    if (!names.length || !workbook) return;
    dirtySheets.clear();
    const capturedCoords = new Map();
    for (const name of names) {
      const coords = dirtyCellCoords.get(name);
      if (coords && coords.size) capturedCoords.set(name, new Set(coords));
      dirtyCellCoords.delete(name);
    }
    const savedWorkbook = workbook;
    const savedWbKey = WB_KEY;
    const baselines = syncedSheetSnapshots.get(savedWbKey) || new Map();
    const payload = {
      // Cell edits do not need to rewrite workbook metadata. Metadata is
      // attached only when a structural sheet change is sent.
      meta: {},
      sheets: {},
      cellPatches: {},
      deleted: []
    };
    const snapshots = new Map();
    for (const n of names) {
      if (!savedWorkbook.sheets[n]) {
        payload.deleted.push(n);
        continue;
      }
      const current = savedWorkbook.sheets[n];
      const tracked = capturedCoords.get(n);
      let change;
      if (tracked && tracked.size && baselines.get(n)) {
        const set = {};
        const deleted = [];
        for (const coord of tracked) {
          const cell = current.cells && current.cells[coord];
          if (cell === undefined) deleted.push(coord);
          else set[coord] = cloneJson(cell);
        }
        change = { cells: { set, delete: deleted } };
        snapshots.set(n, { cells: cloneJson(change.cells) });
      } else {
        change = buildSheetChange(current, baselines.get(n));
        snapshots.set(n, { full: cloneJson(current) });
      }
      if (change.full) payload.sheets[n] = change.full;
      else if (change.cells.set && (Object.keys(change.cells.set).length || change.cells.delete.length)) payload.cellPatches[n] = change.cells;
    }
    if (Object.keys(payload.sheets).length || payload.deleted.length) {
      payload.meta = {
        workbookName: savedWorkbook.workbookName,
        sheetOrder: savedWorkbook.sheetOrder,
        tables: savedWorkbook.tables || {},
        pivots: savedWorkbook.pivots || [],
        activeSheet
      };
    }
    if (!Object.keys(payload.sheets).length && !Object.keys(payload.cellPatches).length && !payload.deleted.length) return;
    saving = saving.then(async () => {
      try {
        await savePatchWithRetry(`/api/workbook?wb=${savedWbKey}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        const nextBaselines = syncedSheetSnapshots.get(savedWbKey) || new Map();
        for (const [name, snapshot] of snapshots) {
          if (snapshot.full) {
            nextBaselines.set(name, snapshot.full);
            continue;
          }
          const baseline = nextBaselines.get(name);
          if (!baseline) continue;
          if (!baseline.cells) baseline.cells = {};
          for (const [coord, cell] of Object.entries(snapshot.cells.set || {})) baseline.cells[coord] = cell;
          for (const coord of (snapshot.cells.delete || [])) delete baseline.cells[coord];
        }
        for (const name of payload.deleted) nextBaselines.delete(name);
        syncedSheetSnapshots.set(savedWbKey, nextBaselines);
        void writeCachedWorkbook(savedWbKey, savedWorkbook);
      } catch (e) {
        names.forEach(markSheetDirty); // retry with the next save
        for (const [name, coords] of capturedCoords) for (const coord of coords) markCellDirty(name, coord);
        showToast('No se pudo confirmar el guardado; se reintentará automáticamente', true);
      }
    });
  }

  function saveWorkbookState() {
    markSheetDirty(activeSheet);
    clearTimeout(saveTimeout);
    saveTimeout = setTimeout(queueWorkbookSave, 300);
  }

  async function flushPendingSave() {
    clearTimeout(saveTimeout);
    if (saveTimeout !== null) queueWorkbookSave();
    await saving;
  }

  function setActualizarDataButtonState(state) {
    const button = document.getElementById('btnActualizarData');
    if (!button) return;
    const running = state && state.status === 'running';
    button.disabled = running;
    button.classList.toggle('is-running', running);
    button.setAttribute('aria-busy', running ? 'true' : 'false');
    button.textContent = running ? 'Actualizando Data…' : 'Actualizar Data';
  }

  async function reloadWorkbookAfterActualizarData() {
    workbook = await fetchWorkbookForKey(WB_KEY);
    workbookCache.set(WB_KEY, workbook);
    rememberSyncedWorkbook(WB_KEY, workbook);
    workbookEngineCache.delete(WB_KEY);
    activeSheet = workbook.sheets.PegarData && !workbook.sheets.PegarData.hidden ? 'PegarData' : activeSheet;
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    renderSheetTabs();
    renderSheet();
    selectCell('A1', false);
  }

  async function pollActualizarData() {
    if (WB_KEY !== 'reunion' || actualizarDataPollInFlight) return;
    clearTimeout(actualizarDataPollTimer);
    actualizarDataPollTimer = null;
    actualizarDataPollInFlight = true;
    try {
      const response = await fetch('/api/actualizar-data?wb=reunion', { cache: 'no-store' });
      const state = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(state.error || `Error ${response.status}`);
      setActualizarDataButtonState(state);

      if (state.status === 'running') {
        actualizarDataPollTimer = setTimeout(pollActualizarData, 2000);
        return;
      }

      const noticeId = state.finishedAt || '';
      const lastNotice = localStorage.getItem(actualizarDataNoticeKey) || '';
      if (noticeId && noticeId !== lastNotice && state.status === 'succeeded') {
        await reloadWorkbookAfterActualizarData();
        localStorage.setItem(actualizarDataNoticeKey, noticeId);
        showToast(`PegarData actualizada: ${state.rows} filas.`, false);
      } else if (noticeId && noticeId !== lastNotice && state.status === 'failed') {
        localStorage.setItem(actualizarDataNoticeKey, noticeId);
        showToast(`Actualizar Data falló: ${state.error || 'error desconocido'}`, true);
      }
    } catch (error) {
      setActualizarDataButtonState({ status: 'idle' });
      showToast(`No se pudo consultar Actualizar Data: ${error.message}`, true);
    } finally {
      actualizarDataPollInFlight = false;
    }
  }

  async function actualizarDataDesdePosco() {
    const button = document.getElementById('btnActualizarData');
    if (!button || button.disabled) return;
    setActualizarDataButtonState({ status: 'running' });
    showToast('Descargando POSCO y actualizando PegarData…');
    try {
      // Conserva antes cualquier edición local pendiente de guardar.
      saveWorkbookState();
      await new Promise(resolve => setTimeout(resolve, 500));
      await saving;

      const response = await fetch('/api/actualizar-data?wb=reunion', { method: 'POST' });
      const state = await response.json().catch(() => ({}));
      if (!response.ok && response.status !== 202) throw new Error(state.error || `Error ${response.status}`);
      await pollActualizarData();
    } catch (error) {
      setActualizarDataButtonState({ status: 'idle' });
      showToast(`Actualizar Data falló: ${error.message}`, true);
    }
  }

  function setActualizarInventarioButtonState(state) {
    const button = document.getElementById('btnActualizarInventario');
    if (!button) return;
    const running = state && state.status === 'running';
    button.disabled = running;
    button.classList.toggle('is-running', running);
    button.setAttribute('aria-busy', running ? 'true' : 'false');
    button.textContent = running ? 'Actualizando Inventario…' : 'Actualizar Inventario';
  }

  async function reloadWorkbookAfterActualizarInventario() {
    workbook = await fetchWorkbookForKey(WB_KEY);
    workbookCache.set(WB_KEY, workbook);
    rememberSyncedWorkbook(WB_KEY, workbook);
    workbookEngineCache.delete(WB_KEY);
    activeSheet = workbook.sheets.Inventario && !workbook.sheets.Inventario.hidden ? 'Inventario' : activeSheet;
    engine = new window.FormulaEngine(workbook);
    engine.recalculateAll();
    renderSheetTabs();
    renderSheet();
    selectCell('A1', false);
  }

  async function pollActualizarInventario() {
    if (WB_KEY !== 'reunion' || actualizarInventarioPollInFlight) return;
    clearTimeout(actualizarInventarioPollTimer);
    actualizarInventarioPollTimer = null;
    actualizarInventarioPollInFlight = true;
    try {
      const response = await fetch('/api/actualizar-inventario?wb=reunion', { cache: 'no-store' });
      const state = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(state.error || `Error ${response.status}`);
      setActualizarInventarioButtonState(state);

      if (state.status === 'running') {
        actualizarInventarioPollTimer = setTimeout(pollActualizarInventario, 2000);
        return;
      }

      const noticeId = state.finishedAt || '';
      const lastNotice = localStorage.getItem(actualizarInventarioNoticeKey) || '';
      if (noticeId && noticeId !== lastNotice && state.status === 'succeeded') {
        await reloadWorkbookAfterActualizarInventario();
        localStorage.setItem(actualizarInventarioNoticeKey, noticeId);
        showToast(`Inventario actualizado: ${state.rows} filas.`, false);
      } else if (noticeId && noticeId !== lastNotice && state.status === 'failed') {
        localStorage.setItem(actualizarInventarioNoticeKey, noticeId);
        showToast(`Actualizar Inventario falló: ${state.error || 'error desconocido'}`, true);
      }
    } catch (error) {
      setActualizarInventarioButtonState({ status: 'idle' });
      showToast(`No se pudo consultar Actualizar Inventario: ${error.message}`, true);
    } finally {
      actualizarInventarioPollInFlight = false;
    }
  }

  async function actualizarInventarioDesdePosco() {
    const button = document.getElementById('btnActualizarInventario');
    if (!button || button.disabled) return;
    setActualizarInventarioButtonState({ status: 'running' });
    showToast('Descargando POSCO y actualizando Inventario…');
    try {
      saveWorkbookState();
      await new Promise(resolve => setTimeout(resolve, 500));
      await saving;

      const response = await fetch('/api/actualizar-inventario?wb=reunion', { method: 'POST' });
      const state = await response.json().catch(() => ({}));
      if (!response.ok && response.status !== 202) throw new Error(state.error || `Error ${response.status}`);
      await pollActualizarInventario();
    } catch (error) {
      setActualizarInventarioButtonState({ status: 'idle' });
      showToast(`Actualizar Inventario falló: ${error.message}`, true);
    }
  }

  const FONT_FAMILY_WHITELIST = new Set(['Aptos Narrow', 'Aptos', 'Calibri', 'Arial', 'Georgia', 'Times New Roman', 'Courier New', 'Verdana']);

  function formatCellDisplay(rawVal, s) {
    const numFmt = s && s.numFmt;
    if (!numFmt || numFmt === 'general' || typeof rawVal !== 'number' || Number.isNaN(rawVal)) return rawVal;
    const decimals = s && typeof s.decimals === 'number' ? s.decimals : 2;
    const fixed = rawVal.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    if (numFmt === 'currency') return `$${fixed}`;
    if (numFmt === 'percent') return `${fixed}%`;
    return fixed; // 'number'
  }

  // With border-collapse, the left/top neighbor's border wins ties, so a
  // border is drawn on both cells sharing the edge.
  // Style of a cell: its own, else the sheet default (only inside the area it covered in Excel)
  function styleAt(sheet, row, col) {
    const cell = sheet.cells[`${indexToCol(col)}${row}`];
    if (cell && cell.s) return cell.s;
    const d = sheet.defaultStyle;
    if (!d) return null;
    const b = sheet.styleBounds;
    if (b && (row > b.rows || col > b.cols)) return null;
    return d;
  }

  function getEffectiveBorder(sheet, col, row, rowSpan = 1, colSpan = 1) {
    const b = (c, r) => {
      const st = styleAt(sheet, r, c);
      return (st && st.border) || {};
    };
    const r2 = row + rowSpan - 1;
    const c2 = col + colSpan - 1;
    const out = { top: false, bottom: false, left: false, right: false };
    // a merged block takes each edge from the cells along that edge (and their neighbours)
    for (let c = col; c <= c2; c++) {
      if (b(c, row).top || (row > 1 && b(c, row - 1).bottom)) out.top = true;
      if (b(c, r2).bottom || b(c, r2 + 1).top) out.bottom = true;
    }
    for (let r = row; r <= r2; r++) {
      if (b(col, r).left || (col > 1 && b(col - 1, r).right)) out.left = true;
      if (b(c2, r).right || b(c2 + 1, r).left) out.right = true;
    }
    return out;
  }

  // a cell without its own style uses the sheet default; editing its format starts from that
  function baseStyle(sheet) {
    return sheet && sheet.defaultStyle ? JSON.parse(JSON.stringify(sheet.defaultStyle)) : {};
  }

  function buildCellStyleAttr(s, border) {
    if (!s) s = {};
    const parts = [];
    if (s.bg) parts.push(`background-color:${s.bg}`);
    if (s.color) parts.push(`color:${s.color}`);
    if (s.fontFamily && FONT_FAMILY_WHITELIST.has(s.fontFamily)) parts.push(`font-family:'${s.fontFamily}',${s.fontFamily === 'Aptos Narrow' ? "'Arial Narrow'," : ''}sans-serif`);
    if (s.fontSize) parts.push(`font-size:${Math.max(6, Math.min(72, Number(s.fontSize) || 11))}pt`);
    if (s.bold) parts.push('font-weight:700');
    if (s.italic) parts.push('font-style:italic');
    const decorations = [];
    if (s.underline) decorations.push('underline');
    if (s.strike) decorations.push('line-through');
    if (decorations.length) parts.push(`text-decoration:${decorations.join(' ')}`);
    if (s.align) parts.push(`text-align:${s.align}`);
    if (s.valign) parts.push(`vertical-align:${s.valign}`);
    if (s.wrap) parts.push('white-space:normal', 'word-break:break-word', 'height:auto');
    if (s.indent) parts.push(`padding-left:${4 + Math.min(15, Number(s.indent) || 0) * 12}px`);
    if (border) {
      if (border.top) parts.push('border-top:1px solid #1e293b');
      if (border.right) parts.push('border-right:1px solid #1e293b');
      if (border.bottom) parts.push('border-bottom:1px solid #1e293b');
      if (border.left) parts.push('border-left:1px solid #1e293b');
    }
    return parts.length ? ` style="${parts.join(';')}"` : '';
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getLastRealRow(sheetName) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet || !sheet.cells) return 1;
    let lastRow = 1;
    for (const [coord, cell] of Object.entries(sheet.cells)) {
      if (cell && cell.v !== '' && cell.v !== null && cell.v !== undefined) {
        const p = parseCoord(coord);
        if (p && p.row > lastRow) lastRow = p.row;
      }
    }
    return lastRow;
  }

  function buildMergeMap(sheet) {
    const skipCells = new Set();
    const spanMap = {};
    const merges = sheet.merges || [];

    for (const rng of merges) {
      const parts = rng.split(':');
      if (parts.length !== 2) continue;
      const start = parseCoord(parts[0]);
      const end = parseCoord(parts[1]);
      if (!start || !end) continue;
      const colspan = end.col - start.col + 1;
      const rowspan = end.row - start.row + 1;
      spanMap[start.coord] = { colspan, rowspan };
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          const coord = `${indexToCol(c)}${r}`;
          if (coord !== start.coord) skipCells.add(coord);
        }
      }
    }
    return { skipCells, spanMap };
  }

  // Render Sheet Tabs
  function visibleSheets() {
    return (workbook.sheetOrder || Object.keys(workbook.sheets)).filter(n => workbook.sheets[n] && !workbook.sheets[n].hidden);
  }

  function renderSheetTabs() {
    const order = visibleSheets();
    const canDelete = order.length > 1;
    sheetTabsList.innerHTML = order
      .map(
        name => `<div class="sheet-tab${name === activeSheet ? ' active' : ''}" data-sheet="${escapeHtml(name)}" draggable="true" title="Click to open • Double-click to rename • Right-click for options • Drag to move">
          <span class="tab-color-dot"${workbook.sheets[name].tabColor ? ` style="background:${escapeHtml(workbook.sheets[name].tabColor)}"` : ''}></span>
          <span class="sheet-tab-label">${escapeHtml(name)}</span>
          ${canDelete ? `<span class="sheet-tab-close" data-delete-sheet="${escapeHtml(name)}" title="Delete sheet ${escapeHtml(name)}">×</span>` : ''}
        </div>`
      )
      .join('');
  }

  // ============================================================================
  // GRID RENDERING (virtualized rows, frozen panes, row heights, hidden rows/cols,
  // Excel number formats and conditional formatting)
  // ============================================================================
  const VIRTUAL_THRESHOLD = 250;   // visible rows above which only a window is rendered
  const OVERSCAN_ROWS = 30;
  const COL_VIRTUAL_THRESHOLD = 40; // visible columns above which only the columns on screen are rendered
  const OVERSCAN_COLS = 10;
  const DEFAULT_ROW_PX = 20;
  const HEADER_ROW_PX = 24;
  const ROW_HEADER_PX = 44;
  const view = {
    sheetName: null, visRows: [], offsets: [], visCols: [], colLeft: {}, colPx: {},
    frozenRows: 0, frozenCols: 0, virtual: false, winStart: 0, winEnd: -1,
    mergeInfo: null, cfIndex: null
  };

  const colWidthPx = (sheet, colStr) => Math.max(10, Math.round((sheet.colWidths?.[colStr] || 8.43) * 7 + 5));
  const rowHeightPx = (sheet, r) => {
    const h = sheet.rowHeights && sheet.rowHeights[r];
    return h ? Math.max(1, Math.round(h * 96 / 72)) : DEFAULT_ROW_PX;
  };

  function zoomFactor() { return zoomLevel / 100; }

  function buildConditionalIndex(sheet) {
    const byCol = new Map();
    const list = sheet.condFormats || [];
    list.forEach((cf, cfIdx) => {
      const rects = cf.ranges.map(r => {
        const [a, b] = r.split(':').map(parseCoord);
        if (!a) return null;
        const e = b || a;
        return { r1: Math.min(a.row, e.row), r2: Math.max(a.row, e.row), c1: Math.min(a.col, e.col), c2: Math.max(a.col, e.col) };
      }).filter(Boolean);
      if (!rects.length) return;
      const anchor = { r: rects[0].r1, c: rects[0].c1 };
      for (const rect of rects) {
        for (let c = rect.c1; c <= rect.c2; c++) {
          if (!byCol.has(c)) byCol.set(c, []);
          byCol.get(c).push({ rect, cf, cfIdx, anchor });
        }
      }
    });
    return byCol;
  }

  function conditionalStyleFor(sheet, r, c, value) {
    if (!view.cfIndex) return null;
    const entries = view.cfIndex.get(c);
    if (!entries) return null;
    const matches = [];
    for (const en of entries) {
      if (r < en.rect.r1 || r > en.rect.r2) continue;
      for (const rule of en.cf.rules) matches.push({ rule, anchor: en.anchor });
    }
    if (!matches.length) return null;
    matches.sort((a, b) => (a.rule.priority || 0) - (b.rule.priority || 0));
    const { compareXl, isErr } = window.ExcelUtils;
    const evalF = (f, anchor) => engine.evaluateRelative(f, activeSheet, anchor.r, anchor.c, r, c);
    const out = {};
    for (const { rule, anchor } of matches) {
      let hit = false;
      try {
        if (rule.type === 'expression') {
          const v = evalF(rule.formulas[0], anchor);
          hit = !isErr(v) && (v === true || (typeof v === 'number' && v !== 0));
        } else {
          const cellV = (value === '' || value === null || value === undefined) ? 0 : value;
          if (isErr(cellV) || (typeof cellV === 'string' && cellV.startsWith('#') && sheet.cells[`${indexToCol(c)}${r}`]?.e)) continue;
          const a = evalF(rule.formulas[0], anchor);
          const b = rule.formulas[1] !== undefined ? evalF(rule.formulas[1], anchor) : null;
          if (isErr(a) || isErr(b)) continue;
          const cmp = x => compareXl(cellV, x === null || x === undefined ? 0 : x);
          switch (rule.operator) {
            case 'equal': hit = cmp(a) === 0; break;
            case 'notEqual': hit = cmp(a) !== 0; break;
            case 'greaterThan': hit = cmp(a) > 0; break;
            case 'lessThan': hit = cmp(a) < 0; break;
            case 'greaterThanOrEqual': hit = cmp(a) >= 0; break;
            case 'lessThanOrEqual': hit = cmp(a) <= 0; break;
            case 'between': { const lo = compareXl(a, b) <= 0 ? a : b; const hi = lo === a ? b : a; hit = cmp(lo) >= 0 && cmp(hi) <= 0; break; }
            case 'notBetween': { const lo = compareXl(a, b) <= 0 ? a : b; const hi = lo === a ? b : a; hit = cmp(lo) < 0 || cmp(hi) > 0; break; }
            default: hit = false;
          }
        }
      } catch (e) {
        hit = false;
      }
      if (hit) {
        for (const [k, v] of Object.entries(rule.style || {})) if (!(k in out)) out[k] = v;
        if (rule.stopIfTrue) break;
      }
    }
    return Object.keys(out).length ? out : null;
  }

  function cellDisplay(cellObj, style) {
    const raw = cellObj && cellObj.v !== undefined && cellObj.v !== null ? cellObj.v : '';
    if (cellObj && cellObj.e) return { text: String(raw), kind: 'error', color: null };
    if (raw === '') return { text: '', kind: 'blank', color: null };
    if (typeof raw === 'boolean') return { text: raw ? 'TRUE' : 'FALSE', kind: 'bool', color: null };
    if (typeof raw === 'number' && style && style.numFmt) return { text: String(formatCellDisplay(raw, style)), kind: 'num', color: null };
    const fmt = style && style.nf;
    const r = window.ExcelNumberFormat.formatValue(raw, fmt);
    return { text: r.text, kind: typeof raw === 'number' ? 'num' : 'text', color: r.color };
  }

  // ============================================================================
  // AUTOFILTER (Excel-style: header dropdowns, value checklist, sort, clear)
  // ============================================================================
  function parseRef(ref) {
    const [a, b] = String(ref).split(':').map(parseCoord);
    if (!a) return null;
    const e = b || a;
    return { r1: Math.min(a.row, e.row), r2: Math.max(a.row, e.row), c1: Math.min(a.col, e.col), c2: Math.max(a.col, e.col) };
  }

  function cellText(sheet, r, c) {
    const cellObj = sheet.cells[`${indexToCol(c)}${r}`];
    return cellDisplay(cellObj, styleAt(sheet, r, c)).text;
  }

  function wildcardRe(pattern) {
    let re = '';
    for (let i = 0; i < pattern.length; i++) {
      const ch = pattern[i];
      if (ch === '~' && i + 1 < pattern.length) { re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
      re += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp(`^${re}$`, 'i');
  }

  function filterColumnPasses(spec, sheet, r, c) {
    const text = cellText(sheet, r, c);
    if (spec.values || spec.blanks) {
      if (text === '') return !!spec.blanks;
      const low = text.toLowerCase();
      return (spec.values || []).some(v => String(v).toLowerCase() === low);
    }
    if (spec.custom && spec.custom.length) {
      const raw = sheet.cells[`${indexToCol(c)}${r}`]?.v;
      const test = ({ op, val }) => {
        const n = Number(val);
        const numeric = typeof raw === 'number' && val !== '' && Number.isFinite(n);
        if (op === 'equal' || op === 'notEqual') {
          const hit = /[*?]/.test(String(val)) ? wildcardRe(String(val)).test(text) : (numeric ? raw === n : text.toLowerCase() === String(val).toLowerCase());
          return op === 'equal' ? hit : !hit;
        }
        const a = numeric ? raw : text.toLowerCase();
        const b = numeric ? n : String(val).toLowerCase();
        return op === 'greaterThan' ? a > b : op === 'greaterThanOrEqual' ? a >= b : op === 'lessThan' ? a < b : op === 'lessThanOrEqual' ? a <= b : true;
      };
      return spec.and ? spec.custom.every(test) : spec.custom.some(test);
    }
    return true;
  }

  function recomputeFilterHidden(sheet) {
    const hidden = new Set();
    for (const f of sheet.filters || []) {
      const rg = parseRef(f.ref);
      if (!rg) continue;
      const cols = Object.entries(f.columns || {});
      if (!cols.length) continue;
      for (let r = rg.r1 + 1; r <= rg.r2; r++) {
        for (const [colId, spec] of cols) {
          if (!filterColumnPasses(spec, sheet, r, rg.c1 + Number(colId))) { hidden.add(r); break; }
        }
      }
    }
    sheet.filterHiddenRows = [...hidden].sort((a, b) => a - b);
  }

  function filterHeaderMap(sheet) {
    const map = new Map();
    (sheet.filters || []).forEach((f, fi) => {
      const rg = parseRef(f.ref);
      if (!rg) return;
      for (let c = rg.c1; c <= rg.c2; c++) {
        map.set(`${indexToCol(c)}${rg.r1}`, { fi, colId: c - rg.c1, active: !!(f.columns && f.columns[String(c - rg.c1)]) });
      }
    });
    return map;
  }

  let afMenu = null;
  function closeFilterMenu() {
    if (afMenu) { afMenu.remove(); afMenu = null; }
  }

  function openFilterMenu(btn) {
    closeFilterMenu();
    const sheet = workbook.sheets[activeSheet];
    const [fi, colId] = btn.dataset.af.split(':').map(Number);
    const f = sheet.filters[fi];
    const rg = parseRef(f.ref);
    const col = rg.c1 + colId;
    const header = cellText(sheet, rg.r1, col) || indexToCol(col);
    const current = (f.columns || {})[String(colId)];

    // values visible under the other columns' filters
    const others = Object.entries(f.columns || {}).filter(([k]) => Number(k) !== colId);
    const counts = new Map();
    let hasBlank = false;
    for (let r = rg.r1 + 1; r <= rg.r2; r++) {
      if (others.some(([k, spec]) => !filterColumnPasses(spec, sheet, r, rg.c1 + Number(k)))) continue;
      const t = cellText(sheet, r, col);
      if (t === '') { hasBlank = true; continue; }
      const key = t.toLowerCase();
      if (!counts.has(key)) {
        const raw = sheet.cells[`${indexToCol(col)}${r}`]?.v;
        counts.set(key, { text: t, num: typeof raw === 'number' ? raw : null });
      }
    }
    const items = [...counts.values()].sort((a, b) => {
      if (a.num !== null && b.num !== null) return a.num - b.num;
      if (a.num !== null) return -1;
      if (b.num !== null) return 1;
      return a.text.localeCompare(b.text, 'es', { sensitivity: 'base' });
    });
    const isChecked = t => {
      if (!current) return true;
      if (current.values || current.blanks) return (current.values || []).some(v => String(v).toLowerCase() === t.toLowerCase());
      return filterValueByCustom(current, t);
    };
    const filterValueByCustom = () => true;

    afMenu = document.createElement('div');
    afMenu.className = 'af-menu no-print';
    afMenu.innerHTML = `
      <button type="button" class="af-action" data-af-act="asc">↑ Ordenar de A a Z</button>
      <button type="button" class="af-action" data-af-act="desc">↓ Ordenar de Z a A</button>
      <div class="af-sep"></div>
      <button type="button" class="af-action" data-af-act="clear" ${current ? '' : 'disabled'}>✕ Borrar filtro de "${escapeHtml(header)}"</button>
      <div class="af-sep"></div>
      <input type="search" class="af-search" placeholder="Buscar" autocomplete="off" />
      <div class="af-list">
        <label class="af-item af-all"><input type="checkbox" data-af-all /> (Seleccionar todo)</label>
        ${items.map((it, i) => `<label class="af-item" data-text="${escapeHtml(it.text.toLowerCase())}"><input type="checkbox" data-af-val="${i}" ${isChecked(it.text) ? 'checked' : ''} /> ${escapeHtml(it.text)}</label>`).join('')}
        ${hasBlank ? `<label class="af-item" data-text="(vacías)"><input type="checkbox" data-af-blank ${!current || current.blanks ? 'checked' : ''} /> (Vacías)</label>` : ''}
      </div>
      <div class="af-footer">
        <button type="button" class="af-ok" data-af-act="ok">Aceptar</button>
        <button type="button" class="af-cancel" data-af-act="cancel">Cancelar</button>
      </div>`;
    document.body.appendChild(afMenu);
    const rect = btn.getBoundingClientRect();
    const menuW = 260;
    afMenu.style.left = `${Math.max(8, Math.min(window.innerWidth - menuW - 8, rect.left))}px`;
    afMenu.style.top = `${Math.min(window.innerHeight - 420, rect.bottom + 2)}px`;

    const boxes = () => [...afMenu.querySelectorAll('[data-af-val], [data-af-blank]')];
    const allBox = afMenu.querySelector('[data-af-all]');
    const syncAll = () => {
      const vis = boxes().filter(b => b.closest('.af-item').style.display !== 'none');
      allBox.checked = vis.length > 0 && vis.every(b => b.checked);
      allBox.indeterminate = !allBox.checked && vis.some(b => b.checked);
    };
    syncAll();
    allBox.addEventListener('change', () => {
      boxes().forEach(b => { if (b.closest('.af-item').style.display !== 'none') b.checked = allBox.checked; });
      syncAll();
    });
    afMenu.querySelector('.af-list').addEventListener('change', e => { if (!e.target.matches('[data-af-all]')) syncAll(); });
    afMenu.querySelector('.af-search').addEventListener('input', e => {
      const q = e.target.value.trim().toLowerCase();
      afMenu.querySelectorAll('.af-item:not(.af-all)').forEach(el => {
        el.style.display = !q || el.dataset.text.includes(q) ? '' : 'none';
      });
      // like Excel: searching selects only the matches
      if (q) boxes().forEach(b => { b.checked = b.closest('.af-item').style.display !== 'none'; });
      syncAll();
    });
    afMenu.addEventListener('click', e => {
      const act = e.target.closest('[data-af-act]')?.dataset.afAct;
      if (!act) return;
      if (act === 'cancel') return closeFilterMenu();
      if (act === 'asc' || act === 'desc') { closeFilterMenu(); return sortFilterRange(fi, colId, act === 'desc'); }
      pushUndoState([activeSheet]);
      f.columns = f.columns || {};
      if (act === 'clear') delete f.columns[String(colId)];
      if (act === 'ok') {
        const selected = boxes().filter(b => b.matches('[data-af-val]') && b.checked).map(b => items[Number(b.dataset.afVal)].text);
        const blankBox = afMenu.querySelector('[data-af-blank]');
        const blanks = blankBox ? blankBox.checked : false;
        const all = selected.length === items.length && (!hasBlank || blanks);
        if (all) delete f.columns[String(colId)];
        else f.columns[String(colId)] = blanks ? { values: selected, blanks: true } : { values: selected };
      }
      closeFilterMenu();
      recomputeFilterHidden(sheet);
      engine.recalculateAll(); // SUBTOTAL / AGGREGATE skip the rows the filter hides
      renderSheet();
      updateStatusMetrics();
      saveWorkbookState();
    });
    setTimeout(() => afMenu && afMenu.querySelector('.af-search').focus(), 0);
  }

  // Sort the visible rows of a filter range by one column (hidden rows stay where they are)
  function sortFilterRange(fi, colId, desc) {
    const sheet = workbook.sheets[activeSheet];
    const f = sheet.filters[fi];
    const rg = parseRef(f.ref);
    const keyCol = rg.c1 + colId;
    const hidden = new Set([...(sheet.hiddenRows || []), ...(sheet.filterHiddenRows || [])]);
    const rows = [];
    for (let r = rg.r1 + 1; r <= rg.r2; r++) if (!hidden.has(r)) rows.push(r);
    const keyOf = r => {
      const c = sheet.cells[`${indexToCol(keyCol)}${r}`];
      const v = c && c.v !== undefined && c.v !== null ? c.v : '';
      return v;
    };
    const rank = v => (v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'boolean' ? 2 : 1);
    const order = rows.slice().sort((a, b) => {
      const va = keyOf(a), vb = keyOf(b);
      const ra = rank(va), rb = rank(vb);
      if (ra === 3 || rb === 3) return ra === rb ? a - b : (ra === 3 ? 1 : -1); // blanks always last
      let cmp = ra - rb;
      if (cmp === 0) cmp = ra === 0 ? va - vb : String(va).localeCompare(String(vb), 'es', { sensitivity: 'base' });
      if (cmp === 0) return a - b;
      return desc ? -cmp : cmp;
    });
    pushUndoState([activeSheet]);
    const snapshot = {};
    for (const r of rows) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const coord = `${indexToCol(c)}${r}`;
        snapshot[coord] = sheet.cells[coord] ? JSON.parse(JSON.stringify(sheet.cells[coord])) : null;
      }
    }
    rows.forEach((destRow, i) => {
      const srcRow = order[i];
      for (let c = rg.c1; c <= rg.c2; c++) {
        const src = snapshot[`${indexToCol(c)}${srcRow}`];
        const destCoord = `${indexToCol(c)}${destRow}`;
        if (!src) { delete sheet.cells[destCoord]; continue; }
        if (src.f) { src.f = shiftFormula(src.f, destRow - srcRow, 0); src.v = ''; }
        sheet.cells[destCoord] = src;
      }
    });
    engine.recalculateAll();
    renderSheet();
    updateStatusMetrics();
    saveWorkbookState();
    showToast(`Ordenado ${desc ? 'de Z a A' : 'de A a Z'}`);
  }

  // Reunion: flower rows whose numbers are all 0 (or empty) are hidden automatically;
  // sheet.showZeroRows (button "Flores en 0") shows them again
  function zeroRowsArea(sheet) {
    if (WB_KEY !== 'reunion' || activeSheet !== 'Reunion' || sheet !== workbook.sheets.Reunion) return null;
    const ref = (sheet.filters && sheet.filters[0] && sheet.filters[0].ref) || 'A2:BK209';
    const [a, b] = ref.split(':').map(parseCoord);
    if (!a || !b) return null;
    return { r1: a.row + 1, r2: b.row, c1: a.col + 1, c2: b.col };
  }
  function zeroRows(sheet, hiddenCols) {
    const area = zeroRowsArea(sheet);
    if (!area || !engine) return [];
    const cols = [];
    for (let c = area.c1; c <= area.c2; c++) if (!hiddenCols.has(indexToCol(c))) cols.push(indexToCol(c));
    const out = [];
    for (let r = area.r1; r <= area.r2; r++) {
      const label = sheet.cells[`A${r}`];
      if (!label || label.v === undefined || label.v === '') continue;
      const allZero = cols.every(col => {
        const v = engine.getCellValue(activeSheet, `${col}${r}`);
        return v === 0 || v === '' || v === null || v === undefined;
      });
      if (allZero) out.push(r);
    }
    return out;
  }
  function updateZeroRowsButton(sheet, count) {
    const btn = document.getElementById('btnZeroRows');
    if (!btn) return;
    const area = zeroRowsArea(sheet);
    btn.style.display = area ? '' : 'none';
    if (!area) return;
    const hidden = !sheet.showZeroRows;
    btn.classList.toggle('checked', hidden);
    btn.querySelector('.chk').textContent = hidden ? '✓' : '';
    btn.querySelector('.lbl').textContent = `Ocultar flores en 0 (${count})`;
  }

  function computeLayout(sheet) {
    const hiddenRows = new Set([...(sheet.hiddenRows || []), ...(sheet.filterHiddenRows || [])]);
    const hiddenCols = new Set(sheet.hiddenCols || []);
    view.rowCache = new Map(); // data or layout may have changed: rows are rebuilt
    const zeros = zeroRows(sheet, hiddenCols);
    if (!sheet.showZeroRows) zeros.forEach(r => hiddenRows.add(r));
    updateZeroRowsButton(sheet, zeros.length);
    const visRows = [];
    const offsets = [0];
    for (let r = 1; r <= sheet.maxRow; r++) {
      if (hiddenRows.has(r)) continue;
      visRows.push(r);
      offsets.push(offsets[offsets.length - 1] + rowHeightPx(sheet, r));
    }
    const visCols = [];
    const colLeft = {};
    const colPx = {};
    let left = ROW_HEADER_PX;
    let colOrder = [];
    for (let c = 1; c <= sheet.maxCol; c++) colOrder.push(c);
    // expandable day columns (weeks.js) are shown next to their week, not at their real position
    if (window.MauriWeeks) colOrder = window.MauriWeeks.orderColumns(sheet, colOrder);
    for (const c of colOrder) {
      const colStr = indexToCol(c);
      if (hiddenCols.has(colStr)) continue;
      visCols.push(c);
      colLeft[c] = left;
      colPx[c] = colWidthPx(sheet, colStr);
      if (window.MauriReq) colPx[c] = Math.max(colPx[c], window.MauriReq.minColPx(activeSheet, sheet, c));
      left += colPx[c];
    }
    const fr = sheet.freeze ? sheet.freeze.rows || 0 : 0;
    const fc = sheet.freeze ? sheet.freeze.cols || 0 : 0;
    const colIdx = new Map(visCols.map((c, i) => [c, i]));
    const rowIdx = new Map(visRows.map((r, i) => [r, i]));
    const nFrozenCols = visCols.filter(c => c <= fc).length;
    Object.assign(view, {
      sheetName: activeSheet, visRows, offsets, visCols, colLeft, colPx, colIdx, rowIdx, nFrozenCols,
      totalWidth: left,
      colVirtual: visCols.length > COL_VIRTUAL_THRESHOLD,
      frozenRows: visRows.filter(r => r <= fr).length,
      frozenCols: fc,
      virtual: visRows.length > VIRTUAL_THRESHOLD,
      mergeInfo: buildMergeMap(sheet),
      hiddenRows, hiddenCols,
      cfIndex: buildConditionalIndex(sheet),
      filterHeaders: filterHeaderMap(sheet)
    });
  }

  function visibleSpan(list, hiddenSet, from, to, isCol) {
    let n = 0;
    for (let i = from; i <= to; i++) if (!hiddenSet.has(isCol ? indexToCol(i) : i)) n++;
    return n;
  }

  function renderRowHtml(sheet, r, frozenTop) {
    const { skipCells, spanMap } = view.mergeInfo;
    const h = rowHeightPx(sheet, r);
    const stickyRow = frozenTop !== null;
    const rowSticky = stickyRow ? `position:sticky;top:${frozenTop}px;z-index:9;` : '';
    let tr = `<tr data-row="${r}" style="height:${h}px">`;
    tr += `<th class="row-header" data-row="${r}" style="${stickyRow ? `top:${frozenTop}px;z-index:12;` : ''}">${r}</th>`;
    const cw = view.colWin;
    for (let k = 0; k < cw.cols.length; k++) {
      if (k === cw.nFrozen && cw.leftPad > 0) tr += `<td class="h-spacer" style="width:${cw.leftPad}px"></td>`;
      const c = cw.cols[k];
      const colStr = indexToCol(c);
      const coord = `${colStr}${r}`;
      if (skipCells.has(coord)) continue;
      const span = spanMap[coord];
      let spanAttrs = '';
      if (span) {
        let cs = 0; // columns of the block that are rendered
        for (let x = c; x < c + span.colspan; x++) if (cw.set.has(x)) cs++;
        const rs = visibleSpan(null, view.hiddenRows, r, r + span.rowspan - 1, false);
        if (cs > 1) spanAttrs += ` colspan="${cs}"`;
        if (rs > 1) spanAttrs += ` rowspan="${rs}"`;
      }
      const cellObj = sheet.cells[coord];
      const style = styleAt(sheet, r, c);
      const disp = cellDisplay(cellObj, style);
      let cond = conditionalStyleFor(sheet, r, c, cellObj ? cellObj.v : null);
      // Plan de cosecha: Flor in light red while its combination is not homologated (homolog.js)
      const hm = window.MauriHomolog && window.MauriHomolog.cellStyle(activeSheet, r, colStr, sheet);
      if (hm) cond = { ...(cond || {}), ...hm };
      // Requerimiento Vs Proyeccion: highlight negative pivot values only in
      // the Weeks x FechaProduccion table. Keep the warning local to the cell.
      if (
        activeSheet === 'Weeks x FechaProduccion' &&
        cellObj && typeof cellObj.v === 'number' && cellObj.v < 0 &&
        window.MauriPivot && window.MauriPivot.pivotAt &&
        window.MauriPivot.pivotAt(activeSheet, coord)
      ) {
        cond = { ...(cond || {}), bg: '#fde2e2', color: '#b91c1c' };
      }
      const merged = cond ? { ...(style || {}), ...cond } : style;
      // cells draw their own borders (separate model, so sticky panes keep them): a shared edge is drawn
      // once, as the bottom/right of the visible cell before it. When hidden rows/columns sit in between,
      // that cell only knows its own bottom/right, so this cell keeps the line otherwise.
      const edge = getEffectiveBorder(sheet, c, r, span ? span.rowspan : 1, span ? span.colspan : 1);
      const bOf = (cc, rr) => { const st = styleAt(sheet, rr, cc); return (st && st.border) || {}; };
      const prevRow = view.visRows[view.rowIdx.get(r) - 1];
      if (prevRow === r - 1) edge.top = false;
      else if (prevRow !== undefined && edge.top && (bOf(c, prevRow).bottom || bOf(c, prevRow + 1).top)) edge.top = false;
      const prevCol = view.visCols[view.colIdx.get(c) - 1];
      if (prevCol === c - 1) edge.left = false;
      else if (prevCol !== undefined && edge.left && (bOf(prevCol, r).right || bOf(prevCol + 1, r).left)) edge.left = false;
      let styleAttr = buildCellStyleAttr(merged, edge);
      const extra = [];
      if (disp.color && !(cond && cond.color)) extra.push(`color:${disp.color}`);
      const frozenCol = c <= view.frozenCols;
      if (frozenCol) extra.push(`position:sticky;left:${view.colLeft[c]}px;z-index:${stickyRow ? 11 : 7}`);
      else if (stickyRow) extra.push(rowSticky.slice(0, -1));
      if (extra.length) {
        styleAttr = styleAttr
          ? styleAttr.replace(/"$/, `;${extra.join(';')}"`)
          : ` style="${extra.join(';')}"`;
      }
      if (frozenCol && stickyRow) styleAttr = styleAttr.replace(/"$/, `;top:${frozenTop}px"`);
      const cls = [];
      if (disp.kind === 'num') cls.push('num-cell');
      else if (disp.kind === 'error' || disp.kind === 'bool') cls.push('center-cell');
      if (frozenCol) cls.push('frozen-col');
      if (frozenCol && c === view.frozenCols) cls.push('frozen-col-edge');
      if (stickyRow && r === view.visRows[view.frozenRows - 1]) cls.push('frozen-row-edge');
      const af = view.filterHeaders && view.filterHeaders.get(coord);
      const afBtn = af ? `<button type="button" class="af-btn${af.active ? ' active' : ''}" data-af="${af.fi}:${af.colId}" title="Filtro">${af.active ? '⏷' : '▾'}</button>` : '';
      if (af) cls.push('af-header');
      // Plan de cosecha: small Req box from Reunion inside day / week cells (reqbadges.js)
      const reqBadge = (window.MauriPivot && window.MauriPivot.cellExtra ? window.MauriPivot.cellExtra(activeSheet, r, c) : '') // pivot +/-
        + (window.MauriReq ? window.MauriReq.cellExtra(activeSheet, r, c, sheet) : '')
        + (window.MauriCorMarks ? window.MauriCorMarks.cellExtra(activeSheet, r, c, sheet) : ''); // Reunion: Cor changed in Plan de cosecha
      if (reqBadge) cls.push('has-req');
      tr += `<td data-coord="${coord}"${spanAttrs}${cls.length ? ` class="${cls.join(' ')}"` : ''}${styleAttr}>${reqBadge}${escapeHtml(disp.text)}${afBtn}</td>`;
    }
    if (cw.rightPad > 0) tr += `<td class="h-spacer" style="width:${cw.rightPad}px"></td>`;
    return tr + '</tr>';
  }

  // Columns rendered for the current horizontal scroll: frozen ones + the ones on screen (with margin)
  function computeColWindow() {
    const { visCols, colLeft, colPx, nFrozenCols: nf } = view;
    const n = visCols.length;
    let cs = nf, ce = n - 1, coreS = nf, coreE = n - 1;
    if (view.colVirtual && n > nf) {
      const vp = document.getElementById('spreadsheetViewport');
      const z = zoomFactor();
      const frozenW = nf ? colLeft[visCols[nf]] - ROW_HEADER_PX : 0;
      const target = vp.scrollLeft / z + ROW_HEADER_PX + frozenW;
      const width = vp.clientWidth / z;
      let lo = nf, hi = n - 1;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (colLeft[visCols[mid]] + colPx[visCols[mid]] <= target) lo = mid + 1; else hi = mid;
      }
      let endIdx = lo;
      while (endIdx < n - 1 && colLeft[visCols[endIdx]] < target + width) endIdx++;
      coreS = lo; coreE = endIdx;
      cs = Math.max(nf, lo - OVERSCAN_COLS);
      ce = Math.min(n - 1, endIdx + OVERSCAN_COLS);
      // never cut a merged block: start at the anchor of any merge crossing the window's left edge
      const startCol = visCols[cs];
      for (const m of workbook.sheets[activeSheet].merges || []) {
        const [a, b] = m.split(':').map(parseCoord);
        if (a && b && a.col < startCol && b.col >= startCol && view.colIdx.has(a.col)) cs = Math.max(nf, Math.min(cs, view.colIdx.get(a.col)));
      }
    }
    const cols = visCols.slice(0, nf).concat(visCols.slice(cs, ce + 1));
    const firstScrollLeft = n > nf ? colLeft[visCols[nf]] : view.totalWidth;
    const leftPad = cs > nf ? colLeft[visCols[cs]] - firstScrollLeft : 0;
    const rightPad = ce < n - 1 ? view.totalWidth - colLeft[visCols[ce + 1]] : 0;
    return { cs, ce, coreS, coreE, nFrozen: nf, cols, set: new Set(cols), leftPad, rightPad };
  }

  function renderHead(sheet) {
    const cw = view.colWin;
    let headHtml = '<tr><th class="corner-header" title="Select all cells">◢</th>';
    cw.cols.forEach((c, k) => {
      if (k === cw.nFrozen && cw.leftPad > 0) headHtml += `<th class="h-spacer" style="width:${cw.leftPad}px;min-width:${cw.leftPad}px"></th>`;
      const colStr = indexToCol(c);
      const w = view.colPx[c];
      const frozen = c <= view.frozenCols ? `position:sticky;left:${view.colLeft[c]}px;z-index:18;` : '';
      const extraHead = window.MauriWeeks ? window.MauriWeeks.headerHtml(sheet, c) : '';
      headHtml += `<th data-col="${colStr}" style="${frozen}width:${w}px;min-width:${w}px;max-width:${w}px;">${colStr}${extraHead}</th>`;
    });
    if (cw.rightPad > 0) headHtml += `<th class="h-spacer" style="width:${cw.rightPad}px;min-width:${cw.rightPad}px"></th>`;
    headHtml += '</tr>';
    gridHead.innerHTML = headHtml;
  }

  // the horizontal window moved past the rendered columns: rebuild header and rows
  function refreshColWindow() {
    const sheet = workbook.sheets[activeSheet];
    view.colWin = computeColWindow();
    view.rowCache = new Map();
    renderHead(sheet);
    renderBody();
  }

  function computeWindow() {
    const vp = document.getElementById('spreadsheetViewport');
    const n = view.visRows.length;
    if (!view.virtual) return { start: view.frozenRows, end: n - 1 };
    const z = zoomFactor();
    const top = vp.scrollTop / z;
    const height = vp.clientHeight / z;
    const frozenH = view.offsets[view.frozenRows];
    // first visible (non-frozen) row index at the scroll position
    let lo = view.frozenRows, hi = n - 1;
    const target = top + frozenH;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (view.offsets[mid + 1] <= target) lo = mid + 1; else hi = mid;
    }
    let start = Math.max(view.frozenRows, lo - OVERSCAN_ROWS);
    let endIdx = lo;
    while (endIdx < n - 1 && view.offsets[endIdx] < target + height) endIdx++;
    const end = Math.min(n - 1, endIdx + OVERSCAN_ROWS);
    const core = { coreStart: lo, coreEnd: endIdx };
    // never cut a merged block: start at the anchor of any merge crossing the window top
    const startRow = view.visRows[start];
    const sheet = workbook.sheets[activeSheet];
    let minRow = startRow;
    for (const m of sheet.merges || []) {
      const [a, b] = m.split(':').map(parseCoord);
      if (a && b && a.row < startRow && b.row >= startRow) minRow = Math.min(minRow, a.row);
    }
    if (minRow < startRow) {
      while (start > view.frozenRows && view.visRows[start] > minRow) start--;
    }
    return { start, end, ...core };
  }

  function renderBody() {
    const sheet = workbook.sheets[activeSheet];
    const { start, end } = computeWindow();
    view.winStart = start;
    view.winEnd = end;
    const parts = [];
    // frozen rows (sticky below the column header)
    let frozenTop = HEADER_ROW_PX;
    for (let i = 0; i < view.frozenRows; i++) {
      const r = view.visRows[i];
      parts.push(renderRowHtml(sheet, r, frozenTop));
      frozenTop += rowHeightPx(sheet, r);
    }
    const colCount = view.colWin.cols.length + 3;
    const topPad = view.offsets[start] - view.offsets[view.frozenRows];
    if (view.virtual) parts.push(`<tr class="v-spacer v-top" style="height:${Math.max(0, topPad)}px"><td colspan="${colCount}"></td></tr>`);
    // rows already built since the last full render are reused (scrolling only moves the window)
    for (let i = start; i <= end; i++) parts.push(rowHtmlCached(sheet, view.visRows[i]));
    const bottomPad = view.offsets[view.visRows.length] - view.offsets[end + 1];
    if (view.virtual) parts.push(`<tr class="v-spacer v-bottom" style="height:${Math.max(0, bottomPad)}px"><td colspan="${colCount}"></td></tr>`);
    gridBody.innerHTML = parts.join('');
  }

  // Scrolling a virtualized sheet: only the rows entering / leaving the window change in the DOM
  function rowHtmlCached(sheet, r) {
    const cache = view.rowCache || (view.rowCache = new Map());
    let html = cache.get(r);
    if (html === undefined) { html = renderRowHtml(sheet, r, null); cache.set(r, html); }
    return html;
  }
  function shiftWindow() {
    const sheet = workbook.sheets[activeSheet];
    const { start, end } = computeWindow();
    const oldS = view.winStart, oldE = view.winEnd;
    const topSp = gridBody.querySelector('tr.v-top');
    const botSp = gridBody.querySelector('tr.v-bottom');
    if (!topSp || !botSp || end < oldS || start > oldE) { renderBody(); return; }
    for (let i = oldS; i < start; i++) topSp.nextElementSibling.remove();
    for (let i = oldE; i > end; i--) botSp.previousElementSibling.remove();
    if (start < oldS) {
      const parts = [];
      for (let i = start; i < oldS; i++) parts.push(rowHtmlCached(sheet, view.visRows[i]));
      topSp.insertAdjacentHTML('afterend', parts.join(''));
    }
    if (end > oldE) {
      const parts = [];
      for (let i = oldE + 1; i <= end; i++) parts.push(rowHtmlCached(sheet, view.visRows[i]));
      botSp.insertAdjacentHTML('beforebegin', parts.join(''));
    }
    topSp.style.height = `${Math.max(0, view.offsets[start] - view.offsets[view.frozenRows])}px`;
    botSp.style.height = `${Math.max(0, view.offsets[view.visRows.length] - view.offsets[end + 1])}px`;
    view.winStart = start;
    view.winEnd = end;
  }

  // Render Active Sheet Grid
  function renderSheet() {
    closeInlineEditor(false);

    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    const lastReal = getLastRealRow(activeSheet);
    sheet.maxCol = Math.max(sheet.maxCol || 20, 20);
    sheet.maxRow = Math.max(sheet.maxRow || 100, lastReal + 40, 100);

    computeLayout(sheet);

    view.colWin = computeColWindow();
    renderHead(sheet);

    renderBody();
    renderSheetImages(sheet);
    highlightSelectedCellInDOM();
  }

  // Floating pictures (e.g. the macro button on PegarData); click runs the assigned macro
  const MACRO_BY_NAME = {
    Actualizar_ColumnaT_PegarData: () => window.MauriMacros && window.MauriMacros.actualizarColumnaT(),
    ConsolidarDatos_Rapido_Optimizado: () => window.MauriMacros && window.MauriMacros.consolidarDatos()
  };
  function renderSheetImages(sheet) {
    sheetCanvasWrapper.querySelectorAll('.sheet-image').forEach(el => el.remove());
    for (const img of sheet.images || []) {
      if (view.colLeft[img.col] === undefined) continue;
      const idx = view.visRows.findIndex(r => r >= img.row);
      if (idx < 0) continue;
      const el = document.createElement('img');
      el.className = 'sheet-image';
      el.src = img.src;
      el.alt = img.macro ? `Macro ${img.macro}` : 'Imagen';
      el.draggable = false;
      el.style.left = `${view.colLeft[img.col] + (img.dx || 0)}px`;
      el.style.top = `${HEADER_ROW_PX + view.offsets[idx] + (img.dy || 0)}px`;
      if (img.w) el.style.width = `${img.w}px`;
      if (img.h) el.style.height = `${img.h}px`;
      if (img.macro && MACRO_BY_NAME[img.macro]) {
        el.classList.add('has-macro');
        el.title = `Ejecutar macro ${img.macro}`;
        el.addEventListener('mousedown', e => { e.preventDefault(); e.stopPropagation(); });
        el.addEventListener('click', () => MACRO_BY_NAME[img.macro]());
      }
      sheetCanvasWrapper.appendChild(el);
    }
  }

  // Re-render the row window while scrolling a virtualized sheet
  let scrollRaf = null;
  function onViewportScroll() {
    if ((!view.virtual && !view.colVirtual) || view.sheetName !== activeSheet) return;
    if (scrollRaf) return;
    scrollRaf = requestAnimationFrame(() => {
      scrollRaf = null;
      if (inlineEditor) return;
      if (view.colVirtual) {
        const cw = computeColWindow();
        if (cw.coreS < view.colWin.cs || cw.coreE > view.colWin.ce) {
          refreshColWindow();
          highlightSelectedCellInDOM(false);
          return;
        }
      }
      if (!view.virtual) return;
      const { coreStart, coreEnd } = computeWindow();
      if (coreStart >= view.winStart && coreEnd <= view.winEnd) return; // still inside the rendered rows
      shiftWindow();
      highlightSelectedCellInDOM(false);
    });
  }

  // Make sure a row is inside the rendered window (keyboard navigation / Name Box)
  function ensureRowRendered(row) {
    if (!view.virtual) return;
    const idx = view.visRows.indexOf(row);
    if (idx < 0 || idx < view.frozenRows) return;
    if (idx >= view.winStart && idx <= view.winEnd) return;
    const vp = document.getElementById('spreadsheetViewport');
    const z = zoomFactor();
    const frozenH = view.offsets[view.frozenRows];
    vp.scrollTop = Math.max(0, (view.offsets[idx] - frozenH - 60) * z);
    renderBody();
  }

  function getSelectionBounds(startC = selectedCoord, endC = selectionEndCoord) {
    const p1 = parseCoord(startC) || parseCoord('A1');
    const p2 = parseCoord(endC || startC) || p1;
    let minRow = Math.min(p1.row, p2.row);
    let maxRow = Math.max(p1.row, p2.row);
    let minCol = Math.min(p1.col, p2.col);
    let maxCol = Math.max(p1.col, p2.col);

    // Like Excel, a selection always covers whole merged blocks it touches
    const merges = (workbook && workbook.sheets[activeSheet] && workbook.sheets[activeSheet].merges) || [];
    let grown = true;
    while (grown && merges.length) {
      grown = false;
      for (const m of merges) {
        const [a, b] = m.split(':').map(parseCoord);
        if (!a || !b) continue;
        const mr1 = Math.min(a.row, b.row), mr2 = Math.max(a.row, b.row);
        const mc1 = Math.min(a.col, b.col), mc2 = Math.max(a.col, b.col);
        if (mr2 < minRow || mr1 > maxRow || mc2 < minCol || mc1 > maxCol) continue;
        if (mr1 < minRow) { minRow = mr1; grown = true; }
        if (mr2 > maxRow) { maxRow = mr2; grown = true; }
        if (mc1 < minCol) { minCol = mc1; grown = true; }
        if (mc2 > maxCol) { maxCol = mc2; grown = true; }
      }
    }
    const rangeStr = `${indexToCol(minCol)}${minRow}:${indexToCol(maxCol)}${maxRow}`;
    return {
      minRow,
      maxRow,
      minCol,
      maxCol,
      isMulti: minRow !== maxRow || minCol !== maxCol,
      singleMerge: merges.includes(rangeStr) ? rangeStr : null,
      startParsed: p1,
      endParsed: p2
    };
  }

  function findMergeAt(coord) {
    const p = parseCoord(coord);
    const sheet = workbook && workbook.sheets[activeSheet];
    if (!p || !sheet) return null;
    for (const m of sheet.merges || []) {
      const [a, b] = m.split(':').map(parseCoord);
      if (!a || !b) continue;
      const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row);
      const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
      if (p.row >= r1 && p.row <= r2 && p.col >= c1 && p.col <= c2) {
        return { r1, r2, c1, c2, anchor: `${indexToCol(c1)}${r1}` };
      }
    }
    return null;
  }

  function selectCell(coord, extendSelection = false, scroll = true) {
    const parsed = parseCoord(coord);
    if (!parsed) return;

    if (extendSelection) {
      selectionEndCoord = parsed.coord;
    } else {
      const merge = findMergeAt(parsed.coord);
      selectedCoord = merge ? merge.anchor : parsed.coord;
      selectionEndCoord = selectedCoord;
    }

    const bounds = getSelectionBounds();
    if (bounds.singleMerge) {
      cellNameBox.value = bounds.singleMerge.split(':')[0];
    } else if (bounds.isMulti) {
      cellNameBox.value = `${indexToCol(bounds.minCol)}${bounds.minRow}:${indexToCol(bounds.maxCol)}${bounds.maxRow}`;
    } else {
      cellNameBox.value = selectedCoord;
    }

    const sheet = workbook.sheets[activeSheet];
    const cellObj = sheet?.cells[selectedCoord];

    if (cellObj && cellObj.f) {
      formulaInput.value = cellObj.f;
      cellMetaBadge.style.display = 'inline-block';
      cellMetaBadge.className = 'cell-meta-badge formula';
      cellMetaBadge.textContent = `= ${cellObj.v ?? ''}`;
    } else {
      formulaInput.value = cellEditText(sheet, selectedCoord);
      cellMetaBadge.style.display = 'none';
    }

    highlightSelectedCellInDOM(scroll);
    updateStatusMetrics();
    updateFormatToolbarState();
    if (window.MauriPivot) window.MauriPivot.onSelectionChange(activeSheet, selectedCoord);
    if (window.MauriCorMarks) window.MauriCorMarks.onSelect(activeSheet, selectedCoord);
  }

  const NUM_FMT_LABELS = { general: 'General', number: 'Número', currency: 'Moneda', percent: 'Porcentaje' };

  function updateFormatToolbarState() {
    const sheet = workbook.sheets[activeSheet];
    const s = (sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s) || {};
    btnBold.classList.toggle('active', !!s.bold);
    btnItalic.classList.toggle('active', !!s.italic);
    btnUnderline.classList.toggle('active', !!s.underline);
    btnStrike.classList.toggle('active', !!s.strike);
    btnWrapText.classList.toggle('active', !!s.wrap);
    btnAlignLeft.classList.toggle('active', s.align === 'left');
    btnAlignCenter.classList.toggle('active', s.align === 'center');
    btnAlignRight.classList.toggle('active', s.align === 'right');
    btnValignTop.classList.toggle('active', s.valign === 'top');
    btnValignMiddle.classList.toggle('active', s.valign === 'middle');
    btnValignBottom.classList.toggle('active', s.valign === 'bottom');
    btnMergeCenter.classList.toggle('active', !!getSelectionBounds().singleMerge);
    fontFamilySelect.value = FONT_FAMILY_WHITELIST.has(s.fontFamily) ? s.fontFamily : 'Calibri';
    fontSizeSelect.value = String(s.fontSize || 11);
    numFmtLabel.textContent = NUM_FMT_LABELS[s.numFmt] || 'General';
  }

  // keepView: whole row/column selections leave the scroll position alone, like Excel
  function selectRange(startCoordStr, endCoordStr, keepView = false) {
    const p1 = parseCoord(startCoordStr);
    const p2 = parseCoord(endCoordStr);
    if (!p1 || !p2) return;
    selectedCoord = p1.coord;
    selectionEndCoord = p2.coord;
    selectCell(p2.coord, true, !keepView);
  }

  // scrollIntoView ignores the sticky column/row headers, so scroll manually.
  function scrollCellIntoView(td) {
    const vp = document.getElementById('spreadsheetViewport');
    const vr = vp.getBoundingClientRect();
    const r = td.getBoundingClientRect();
    const z = zoomFactor();
    const headH = gridHead.getBoundingClientRect().height;
    const rowHeader = gridBody.querySelector('th.row-header');
    const rowHW = rowHeader ? rowHeader.getBoundingClientRect().width : 0;
    const sheet = workbook.sheets[activeSheet];
    let frozenH = 0;
    for (let i = 0; i < view.frozenRows; i++) frozenH += rowHeightPx(sheet, view.visRows[i]);
    let frozenW = 0;
    for (const c of view.visCols) if (c <= view.frozenCols) frozenW += view.colPx[c];
    const top = vr.top + headH + frozenH * z;
    const left = vr.left + rowHW + frozenW * z;
    const bottom = vr.top + vp.clientHeight;
    const right = vr.left + vp.clientWidth;
    const row = Number(td.parentElement && td.parentElement.dataset.row);
    const p = parseCoord(td.dataset.coord || '');
    const inFrozenRow = view.frozenRows && row <= view.visRows[view.frozenRows - 1];
    const inFrozenCol = p && p.col <= view.frozenCols;
    if (!inFrozenRow) {
      if (r.top < top) vp.scrollTop -= top - r.top;
      else if (r.bottom > bottom) vp.scrollTop += Math.min(r.bottom - bottom, r.top - top);
    }
    if (!inFrozenCol) {
      if (r.left < left) vp.scrollLeft -= left - r.left;
      else if (r.right > right) vp.scrollLeft += Math.min(r.right - right, r.left - left);
    }
  }

  // only selection changes (selectCell) scroll; repainting the selection keeps the view
  function ensureColRendered(col) {
    if (!view.colVirtual || !view.colWin || view.colWin.set.has(col) || !view.colIdx.has(col) || col <= view.frozenCols) return;
    const vp = document.getElementById('spreadsheetViewport');
    const z = zoomFactor();
    const nf = view.nFrozenCols;
    const frozenW = nf ? view.colLeft[view.visCols[nf]] - ROW_HEADER_PX : 0;
    vp.scrollLeft = Math.max(0, (view.colLeft[col] - ROW_HEADER_PX - frozenW - 60) * z);
    refreshColWindow();
  }

  function highlightSelectedCellInDOM(scroll = false) {
    if (scroll) {
      const focusP = parseCoord(selectionEndCoord || selectedCoord);
      if (focusP) { ensureRowRendered(focusP.row); ensureColRendered(focusP.col); }
    }
    gridBody.querySelectorAll('td.selected-cell, td.in-selection-range, td.sel-edge-top, td.sel-edge-bottom, td.sel-edge-left, td.sel-edge-right, td.copied-range-cell, td.autofill-preview-cell').forEach(el => {
      el.classList.remove('selected-cell', 'in-selection-range', 'sel-edge-top', 'sel-edge-bottom', 'sel-edge-left', 'sel-edge-right', 'copied-range-cell', 'autofill-preview-cell');
    });
    gridBody.querySelectorAll('.cell-fill-handle').forEach(el => el.remove());
    gridHead.querySelectorAll('th.active-col-header').forEach(el => el.classList.remove('active-col-header'));
    gridBody.querySelectorAll('th.active-row-header').forEach(el => el.classList.remove('active-row-header'));

    const bounds = getSelectionBounds();
    const { minRow, maxRow, minCol, maxCol } = bounds;

    const anchorTd = gridBody.querySelector(`td[data-coord="${selectedCoord}"]`);
    if (anchorTd) anchorTd.classList.add('selected-cell');

    const focusTd = gridBody.querySelector(`td[data-coord="${selectionEndCoord || selectedCoord}"]`);
    if (focusTd && scroll) scrollCellIntoView(focusTd);

    for (let c = minCol; c <= maxCol; c++) {
      const colStr = indexToCol(c);
      const colTh = gridHead.querySelector(`th[data-col="${colStr}"]`);
      if (colTh) colTh.classList.add('active-col-header');
    }

    // walk the rendered rows once instead of looking up every selected cell
    const clip = clipboardState && clipboardState.sheetName === activeSheet ? clipboardState : null;
    const colOf = coord => colToIndex(coord.replace(/\d+$/, ''));
    for (const tr of gridBody.rows) {
      const r = Number(tr.dataset.row);
      if (!r) continue;
      const inSel = r >= minRow && r <= maxRow;
      const inClip = clip && r >= clip.minRow && r <= clip.maxRow;
      if (!inSel && !inClip) continue;
      for (const cell of tr.cells) {
        if (cell.tagName === 'TH') { if (inSel) cell.classList.add('active-row-header'); continue; }
        const coord = cell.dataset.coord;
        if (!coord) continue;
        const c = colOf(coord);
        if (inClip && c >= clip.minCol && c <= clip.maxCol) cell.classList.add('copied-range-cell');
        if (!inSel || c < minCol || c > maxCol) continue;
        cell.classList.add('in-selection-range');
        if (r === minRow) cell.classList.add('sel-edge-top');
        if (r + (cell.rowSpan || 1) - 1 === maxRow) cell.classList.add('sel-edge-bottom');
        if (c === minCol) cell.classList.add('sel-edge-left');
        if (c + (cell.colSpan || 1) - 1 === maxCol) cell.classList.add('sel-edge-right');
      }
    }

    const brCoord = `${indexToCol(maxCol)}${maxRow}`;
    const brTd = gridBody.querySelector(`td[data-coord="${brCoord}"]`) || anchorTd;
    if (brTd && !inlineEditor) {
      const handle = document.createElement('div');
      handle.className = 'cell-fill-handle';
      handle.title = 'Drag to auto-fill / copy cells';
      brTd.appendChild(handle);
    }

  }

  // ---- dates typed / shown like Excel (stored as serial numbers)
  const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
  function parseDateInput(text) {
    let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/.exec(text);
    let y, mo, d, h = 0, mi = 0, se = 0, ap = null;
    if (m) {
      mo = +m[1]; d = +m[2]; y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
      if (m[4] !== undefined) { h = +m[4]; mi = +m[5]; se = m[6] ? +m[6] : 0; ap = m[7]; }
    } else if ((m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text))) {
      y = +m[1]; mo = +m[2]; d = +m[3];
    } else {
      return null;
    }
    if (ap) { const pm = /p/i.test(ap); if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
    if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCMonth() !== mo - 1) return null;
    const serial = (Date.UTC(y, mo - 1, d) - EXCEL_EPOCH_UTC) / 86400000 + (h * 3600 + mi * 60 + se) / 86400;
    return { serial, hasTime: m[4] !== undefined };
  }

  function cellEditText(sheet, coord) {
    const cellObj = sheet && sheet.cells[coord];
    if (!cellObj) return '';
    if (cellObj.f) return cellObj.f;
    const v = cellObj.v;
    if (v === undefined || v === null) return '';
    const p = parseCoord(coord);
    const st = p ? styleAt(sheet, p.row, p.col) : cellObj.s;
    if (typeof v === 'number' && st && st.nf && window.ExcelNumberFormat.isDateFormat(st.nf)) {
      const fmt = Math.abs(v - Math.floor(v)) > 1e-9 ? 'm/d/yyyy h:mm:ss AM/PM' : 'm/d/yyyy';
      return window.ExcelNumberFormat.formatValue(v, fmt).text;
    }
    return v;
  }

  function applyCellValueRaw(sheetName, coord, rawInput) {
    writeCellValue(sheetName, coord, rawInput);
    markCellDirty(sheetName, coord);
    if (window.MauriCorMarks && window.MauriCorMarks.onUserEdit) window.MauriCorMarks.onUserEdit(sheetName, coord);
  }

  function writeCellValue(sheetName, coord, rawInput) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet) return;
    const trimmed = String(rawInput ?? '').trim();
    const existingStyle = sheet.cells[coord] && sheet.cells[coord].s ? sheet.cells[coord].s : null;

    if (trimmed === '') {
      if (existingStyle) sheet.cells[coord] = { s: existingStyle };
      else delete sheet.cells[coord];
    } else if (trimmed.startsWith('=')) {
      sheet.cells[coord] = existingStyle ? { f: trimmed, v: '', s: existingStyle } : { f: trimmed, v: '' };
    } else {
      let val = trimmed;
      let style = existingStyle;
      const date = parseDateInput(trimmed);
      if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
        val = Number(trimmed);
      } else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(trimmed)) {
        val = Number(trimmed.replace(/,/g, ''));
      } else if (date) {
        // like Excel: a typed date becomes a date value; give it a date format if it has none
        val = date.serial;
        const p = parseCoord(coord);
        const current = p ? styleAt(sheet, p.row, p.col) : existingStyle;
        if (!(current && current.nf && window.ExcelNumberFormat.isDateFormat(current.nf))) {
          style = { ...(current ? JSON.parse(JSON.stringify(current)) : {}), nf: date.hasTime ? 'm/d/yyyy h:mm' : 'm/d/yyyy' };
          delete style.numFmt;
        }
      }
      sheet.cells[coord] = style ? { v: val, s: style } : { v: val };
    }
  }

  function setCellContent(sheetName, coord, rawInput) {
    const sheet = workbook.sheets[sheetName];
    if (!sheet) return;
    pushUndoState();
    applyCellValueRaw(sheetName, coord, rawInput);
    engine.recalculateAll();
    renderSheet();
    selectCell(coord, false);
    saveWorkbookState();
  }

  function clearSelectedRange() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        applyCellValueRaw(activeSheet, `${indexToCol(c)}${r}`, '');
      }
    }
    engine.recalculateAll();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
  }

  // ============================================================================
  // CELL FORMATTING (fill, font, borders, number format, format painter)
  // ============================================================================

  function applyStyleToSelection(mutator, bounds = getSelectionBounds(), withUndo = true) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    if (withUndo) pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = bounds;
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        if (!sheet.cells[coord]) sheet.cells[coord] = {};
        if (!sheet.cells[coord].s) sheet.cells[coord].s = baseStyle(sheet);
        mutator(sheet.cells[coord].s);
        if (Object.keys(sheet.cells[coord].s).length === 0) delete sheet.cells[coord].s;
        if (Object.keys(sheet.cells[coord]).length === 0) delete sheet.cells[coord];
      }
    }
    renderSheet();
    highlightSelectedCellInDOM();
    updateFormatToolbarState();
    saveWorkbookState();
  }

  function toggleBoldSelection() {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const turnOn = !(anchorStyle && anchorStyle.bold);
    applyStyleToSelection(s => { if (turnOn) s.bold = true; else delete s.bold; });
  }

  function toggleItalicSelection() {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const turnOn = !(anchorStyle && anchorStyle.italic);
    applyStyleToSelection(s => { if (turnOn) s.italic = true; else delete s.italic; });
  }

  function toggleUnderlineSelection() {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const turnOn = !(anchorStyle && anchorStyle.underline);
    applyStyleToSelection(s => { if (turnOn) s.underline = true; else delete s.underline; });
  }

  function toggleStrikeSelection() {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const turnOn = !(anchorStyle && anchorStyle.strike);
    applyStyleToSelection(s => { if (turnOn) s.strike = true; else delete s.strike; });
  }

  function toggleWrapSelection() {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const turnOn = !(anchorStyle && anchorStyle.wrap);
    applyStyleToSelection(s => { if (turnOn) s.wrap = true; else delete s.wrap; });
  }

  function setAlignSelection(align) {
    applyStyleToSelection(s => { s.align = align; });
  }

  function setValignSelection(valign) {
    applyStyleToSelection(s => { s.valign = valign; });
  }

  function setFontFamilySelection(family) {
    applyStyleToSelection(s => { s.fontFamily = family; });
  }

  function setFontSizeSelection(size) {
    applyStyleToSelection(s => { s.fontSize = size; });
  }

  function growShrinkFontSelection(delta) {
    const sheet = workbook.sheets[activeSheet];
    const anchorStyle = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    const current = (anchorStyle && anchorStyle.fontSize) || 11;
    const next = Math.max(6, Math.min(72, current + delta));
    applyStyleToSelection(s => { s.fontSize = next; });
  }

  function setQuickNumFmtSelection(numFmt) {
    applyStyleToSelection(s => { delete s.nf; s.numFmt = numFmt; });
  }

  function adjustDecimalsSelection(delta) {
    applyStyleToSelection(s => {
      delete s.nf;
      if (!s.numFmt || s.numFmt === 'general') s.numFmt = 'number';
      const current = typeof s.decimals === 'number' ? s.decimals : 2;
      s.decimals = Math.max(0, Math.min(10, current + delta));
    });
  }

  function mergeRangeOverlaps(mergeStr, minRow, maxRow, minCol, maxCol) {
    const parts = mergeStr.split(':');
    const p1 = parseCoord(parts[0]);
    const p2 = parseCoord(parts[1]);
    if (!p1 || !p2) return false;
    const mMinRow = Math.min(p1.row, p2.row), mMaxRow = Math.max(p1.row, p2.row);
    const mMinCol = Math.min(p1.col, p2.col), mMaxCol = Math.max(p1.col, p2.col);
    return !(mMaxRow < minRow || mMinRow > maxRow || mMaxCol < minCol || mMinCol > maxCol);
  }

  function toggleMergeCenterSelection() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    sheet.merges = sheet.merges || [];
    const bounds = getSelectionBounds();
    const { minRow, maxRow, minCol, maxCol } = bounds;
    const anchorCoord = `${indexToCol(minCol)}${minRow}`;

    if (!bounds.isMulti) return; // a single plain cell: nothing to merge

    pushUndoState();

    const hasMergeInside = sheet.merges.some(m => mergeRangeOverlaps(m, minRow, maxRow, minCol, maxCol));
    if (hasMergeInside) {
      // Excel: when the selection contains merged cells, the button unmerges them
      sheet.merges = sheet.merges.filter(m => !mergeRangeOverlaps(m, minRow, maxRow, minCol, maxCol));
    } else {
      sheet.merges = sheet.merges.filter(m => !mergeRangeOverlaps(m, minRow, maxRow, minCol, maxCol));
      sheet.merges.push(`${anchorCoord}:${indexToCol(maxCol)}${maxRow}`);
      if (!sheet.cells[anchorCoord]) sheet.cells[anchorCoord] = {};
      if (!sheet.cells[anchorCoord].s) sheet.cells[anchorCoord].s = baseStyle(sheet);
      sheet.cells[anchorCoord].s.align = 'center';
      sheet.cells[anchorCoord].s.valign = 'middle';
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const coord = `${indexToCol(c)}${r}`;
          if (coord === anchorCoord) continue;
          applyCellValueRaw(activeSheet, coord, '');
        }
      }
    }

    engine.recalculateAll();
    renderSheet();
    selectCell(anchorCoord, false);
    saveWorkbookState();
  }

  function applyBorderPreset(preset) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        if (preset === 'none') {
          const cellObj = sheet.cells[coord];
          if (cellObj && cellObj.s) {
            delete cellObj.s.border;
            if (Object.keys(cellObj.s).length === 0) delete cellObj.s;
            if (Object.keys(cellObj).length === 0) delete sheet.cells[coord];
          }
          continue;
        }
        if (!sheet.cells[coord]) sheet.cells[coord] = {};
        if (!sheet.cells[coord].s) sheet.cells[coord].s = baseStyle(sheet);
        const style = sheet.cells[coord].s;
        const border = style.border || {};
        if (preset === 'all') {
          border.top = border.right = border.bottom = border.left = true;
        } else if (preset === 'outer') {
          if (r === minRow) border.top = true;
          if (r === maxRow) border.bottom = true;
          if (c === minCol) border.left = true;
          if (c === maxCol) border.right = true;
        } else if (preset === 'bottom') {
          if (r === maxRow) border.bottom = true;
        }
        style.border = border;
      }
    }
    if (preset === 'none') {
      const clearSide = (coord, side) => {
        const cellObj = sheet.cells[coord];
        if (!cellObj || !cellObj.s || !cellObj.s.border) return;
        delete cellObj.s.border[side];
        if (Object.keys(cellObj.s.border).length === 0) delete cellObj.s.border;
        if (Object.keys(cellObj.s).length === 0) delete cellObj.s;
        if (Object.keys(cellObj).length === 0) delete sheet.cells[coord];
      };
      for (let c = minCol; c <= maxCol; c++) {
        if (minRow > 1) clearSide(`${indexToCol(c)}${minRow - 1}`, 'bottom');
        clearSide(`${indexToCol(c)}${maxRow + 1}`, 'top');
      }
      for (let r = minRow; r <= maxRow; r++) {
        if (minCol > 1) clearSide(`${indexToCol(minCol - 1)}${r}`, 'right');
        clearSide(`${indexToCol(maxCol + 1)}${r}`, 'left');
      }
    }
    renderSheet();
    highlightSelectedCellInDOM();
    saveWorkbookState();
  }

  function clearFormatSelection() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        const cellObj = sheet.cells[coord];
        if (cellObj && cellObj.s) {
          delete cellObj.s;
          if (Object.keys(cellObj).length === 0) delete sheet.cells[coord];
        }
      }
    }
    renderSheet();
    highlightSelectedCellInDOM();
    updateFormatToolbarState();
    saveWorkbookState();
    showToast('Formatting cleared');
  }

  function startFormatPainter() {
    const sheet = workbook.sheets[activeSheet];
    const src = sheet && sheet.cells[selectedCoord] && sheet.cells[selectedCoord].s;
    formatPainterSource = src ? JSON.parse(JSON.stringify(src)) : {};
    formatPainterActive = true;
    gridBody.classList.add('format-painter-active');
    btnFormatPainter.classList.add('active');
    showToast('Format Painter: select the cell(s) to paint');
  }

  function stopFormatPainter() {
    formatPainterActive = false;
    formatPainterSource = null;
    gridBody.classList.remove('format-painter-active');
    btnFormatPainter.classList.remove('active');
  }

  function applyFormatPainterToSelection() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) { stopFormatPainter(); return; }
    pushUndoState();
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    const hasStyle = Object.keys(formatPainterSource || {}).length > 0;
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const coord = `${indexToCol(c)}${r}`;
        if (!sheet.cells[coord]) sheet.cells[coord] = {};
        if (hasStyle) {
          sheet.cells[coord].s = JSON.parse(JSON.stringify(formatPainterSource));
        } else {
          delete sheet.cells[coord].s;
        }
        if (Object.keys(sheet.cells[coord]).length === 0) delete sheet.cells[coord];
      }
    }
    stopFormatPainter();
    renderSheet();
    highlightSelectedCellInDOM();
    updateFormatToolbarState();
    saveWorkbookState();
    showToast('Format applied');
  }

  // Shift relative references of a formula by (dRow, dCol); `$` parts stay fixed (Excel semantics).
  const SHIFT_REF_RE = /(?<![A-Za-z0-9_.$])((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/g;
  function shiftFormula(formula, dRow, dCol) {
    if (!formula || (dRow === 0 && dCol === 0)) return formula;
    return formula.split(/("[^"]*")/).map(part => {
      if (part.startsWith('"')) return part;
      return part.replace(SHIFT_REF_RE, (m, sheetPrefix, colAbs, col, rowAbs, row) => {
        const newCol = colAbs ? colToIndex(col.toUpperCase()) : colToIndex(col.toUpperCase()) + dCol;
        const newRow = rowAbs ? Number(row) : Number(row) + dRow;
        if (newCol < 1 || newRow < 1) return '#REF!';
        return `${sheetPrefix || ''}${colAbs}${indexToCol(newCol)}${rowAbs}${newRow}`;
      });
    }).join('');
  }

  function copyCellWithFormat(sheet, srcR, srcC, destR, destC) {
    const src = sheet.cells[`${indexToCol(srcC)}${srcR}`];
    const destCoord = `${indexToCol(destC)}${destR}`;
    if (!src) { delete sheet.cells[destCoord]; return; }
    const copy = JSON.parse(JSON.stringify(src));
    if (copy.f) {
      copy.f = shiftFormula(copy.f, destR - srcR, destC - srcC);
      copy.v = '';
    }
    sheet.cells[destCoord] = copy;
  }

  function cellDisplayText(cellObj) {
    if (!cellObj || cellObj.v === undefined || cellObj.v === null) return '';
    return String(cellObj.s && cellObj.s.numFmt ? formatCellDisplay(cellObj.v, cellObj.s) : cellObj.v);
  }

  function copySelection(isCut = false) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    const grid = [];
    const tsvRows = [];

    for (let r = minRow; r <= maxRow; r++) {
      const rowCells = [];
      const rowText = [];
      for (let c = minCol; c <= maxCol; c++) {
        const cellObj = sheet.cells[`${indexToCol(c)}${r}`];
        rowCells.push(cellObj ? JSON.parse(JSON.stringify(cellObj)) : null);
        rowText.push(cellDisplayText(cellObj));
      }
      grid.push(rowCells);
      tsvRows.push(rowText.join('\t'));
    }

    const merges = (sheet.merges || [])
      .map(m => m.split(':').map(parseCoord))
      .filter(([a, b]) => a && b && a.row >= minRow && b.row <= maxRow && a.col >= minCol && b.col <= maxCol)
      .map(([a, b]) => ({ r1: a.row - minRow, c1: a.col - minCol, r2: b.row - minRow, c2: b.col - minCol }));

    const tsv = tsvRows.join('\r\n');
    clipboardState = {
      sheetName: activeSheet,
      minRow,
      maxRow,
      minCol,
      maxCol,
      isCut,
      grid,
      merges,
      tsv
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).catch(() => {});
    }

    highlightSelectedCellInDOM();
    const count = (maxRow - minRow + 1) * (maxCol - minCol + 1);
    showToast(`${isCut ? 'Cut' : 'Copied'}: ${count} cell(s) (${indexToCol(minCol)}${minRow}:${indexToCol(maxCol)}${maxRow})`);
  }

  function normalizeClipboardText(txt) {
    return String(txt || '').replace(/\r\n?/g, '\n').replace(/\n$/, '');
  }

  function pasteFromClipboard() {
    if (!(navigator.clipboard && navigator.clipboard.readText)) {
      if (clipboardState) pasteInternalClipboard();
      return;
    }
    const timeout = new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 1500));
    Promise.race([navigator.clipboard.readText(), timeout]).then(txt => {
      if (clipboardState && (!txt || normalizeClipboardText(txt) === normalizeClipboardText(clipboardState.tsv))) {
        pasteInternalClipboard();
      } else if (txt) {
        const rows = normalizeClipboardText(txt).split('\n');
        pasteDataMatrix(rows.map(line => line.split('\t')));
      } else if (clipboardState) {
        pasteInternalClipboard();
      }
    }).catch(() => {
      if (clipboardState) pasteInternalClipboard();
    });
  }

  // Excel-style paste: full cells (value, formula, format, merges), relative formula shift, tiling.
  function pasteInternalClipboard() {
    const clip = clipboardState;
    const sheet = workbook.sheets[activeSheet];
    if (!clip || !sheet) return;
    const srcSheet = workbook.sheets[clip.sheetName];

    pushUndoState([activeSheet, clip.sheetName]);
    markSheetDirty(clip.sheetName);
    const bounds = getSelectionBounds();
    const srcRows = clip.grid.length;
    const srcCols = clip.grid[0].length;
    const selRows = bounds.maxRow - bounds.minRow + 1;
    const selCols = bounds.maxCol - bounds.minCol + 1;
    const tile = selRows % srcRows === 0 && selCols % srcCols === 0;
    const targetRows = tile ? selRows : srcRows;
    const targetCols = tile ? selCols : srcCols;
    const endR = bounds.minRow + targetRows - 1;
    const endC = bounds.minCol + targetCols - 1;

    // Cut: empty the source first (it may overlap the destination)
    if (clip.isCut && srcSheet) {
      for (let r = clip.minRow; r <= clip.maxRow; r++) {
        for (let c = clip.minCol; c <= clip.maxCol; c++) delete srcSheet.cells[`${indexToCol(c)}${r}`];
      }
      srcSheet.merges = (srcSheet.merges || []).filter(m => !mergeRangeOverlaps(m, clip.minRow, clip.maxRow, clip.minCol, clip.maxCol));
    }

    sheet.merges = (sheet.merges || []).filter(m => !mergeRangeOverlaps(m, bounds.minRow, endR, bounds.minCol, endC));

    for (let dr = 0; dr < targetRows; dr++) {
      for (let dc = 0; dc < targetCols; dc++) {
        const destR = bounds.minRow + dr;
        const destC = bounds.minCol + dc;
        const destCoord = `${indexToCol(destC)}${destR}`;
        const src = clip.grid[dr % srcRows][dc % srcCols];
        if (!src) { delete sheet.cells[destCoord]; continue; }
        const copy = JSON.parse(JSON.stringify(src));
        if (copy.f && !clip.isCut) {
          const srcR = clip.minRow + (dr % srcRows);
          const srcC = clip.minCol + (dc % srcCols);
          copy.f = shiftFormula(copy.f, destR - srcR, destC - srcC);
          copy.v = '';
        }
        sheet.cells[destCoord] = copy;
      }
    }

    for (let tr = 0; tr < targetRows; tr += srcRows) {
      for (let tc = 0; tc < targetCols; tc += srcCols) {
        for (const m of clip.merges || []) {
          const r1 = bounds.minRow + tr + m.r1, c1 = bounds.minCol + tc + m.c1;
          sheet.merges.push(`${indexToCol(c1)}${r1}:${indexToCol(bounds.minCol + tc + m.c2)}${bounds.minRow + tr + m.r2}`);
        }
      }
    }

    if (clip.isCut) clipboardState = null;
    if (endR + 5 > sheet.maxRow) sheet.maxRow = endR + 20;
    if (endC + 2 > sheet.maxCol) sheet.maxCol = endC + 5;

    engine.recalculateAll();
    renderSheet();
    selectedCoord = `${indexToCol(bounds.minCol)}${bounds.minRow}`;
    selectionEndCoord = `${indexToCol(endC)}${endR}`;
    selectCell(selectionEndCoord, true, false); // Excel keeps the view after pasting
    saveWorkbookState();
  }

  // Plain text paste from outside the app: values only, destination formatting is kept.
  function pasteDataMatrix(matrix) {
    if (!matrix || matrix.length === 0) return;
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    pushUndoState();
    const bounds = getSelectionBounds();
    const srcRows = matrix.length;
    const srcCols = matrix[0].length;
    const selRows = bounds.maxRow - bounds.minRow + 1;
    const selCols = bounds.maxCol - bounds.minCol + 1;

    const targetRows = (srcRows === 1 && srcCols === 1 && (selRows > 1 || selCols > 1)) ? selRows : srcRows;
    const targetCols = (srcRows === 1 && srcCols === 1 && (selRows > 1 || selCols > 1)) ? selCols : srcCols;

    for (let dr = 0; dr < targetRows; dr++) {
      for (let dc = 0; dc < targetCols; dc++) {
        const destR = bounds.minRow + dr;
        const destC = bounds.minCol + dc;
        const destCoord = `${indexToCol(destC)}${destR}`;
        applyCellValueRaw(activeSheet, destCoord, matrix[dr % srcRows][dc % srcCols]);
      }
    }

    const endR = bounds.minRow + targetRows - 1;
    const endC = bounds.minCol + targetCols - 1;
    if (endR + 5 > sheet.maxRow) sheet.maxRow = endR + 20;
    if (endC + 2 > sheet.maxCol) sheet.maxCol = endC + 5;

    engine.recalculateAll();
    renderSheet();
    selectedCoord = `${indexToCol(bounds.minCol)}${bounds.minRow}`;
    selectionEndCoord = `${indexToCol(endC)}${endR}`;
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
  }

  function fillSelectionDown() {
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    if (maxRow <= minRow) return;
    const sheet = workbook.sheets[activeSheet];
    pushUndoState();
    for (let c = minCol; c <= maxCol; c++) {
      for (let r = minRow + 1; r <= maxRow; r++) copyCellWithFormat(sheet, minRow, c, r, c);
    }
    engine.recalculateAll();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast('Filled down (Ctrl+D)');
  }

  function fillSelectionRight() {
    const { minRow, maxRow, minCol, maxCol } = getSelectionBounds();
    if (maxCol <= minCol) return;
    const sheet = workbook.sheets[activeSheet];
    pushUndoState();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol + 1; c <= maxCol; c++) copyCellWithFormat(sheet, r, minCol, r, c);
    }
    engine.recalculateAll();
    renderSheet();
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast('Filled right (Ctrl+R)');
  }

  // Inline Editor inside Cell
  // Formula reference picking (click/drag cells while typing a formula)
  let refPick = null; // { input, start, end, anchor, focus, dragging }

  function getFormulaEditingInput() {
    if (inlineEditor && document.activeElement === inlineEditor.input) return inlineEditor.input;
    if (document.activeElement === formulaInput) return formulaInput;
    return null;
  }

  function canInsertReference(input) {
    if (!input || !input.value.startsWith('=')) return false;
    if (refPick && refPick.input === input && input.selectionStart === refPick.end) return true;
    const before = input.value.slice(0, input.selectionStart).trimEnd();
    return /[=(,;+\-*/^&<>:]$/.test(before);
  }

  function writePickedReference() {
    if (!refPick) return;
    const p1 = parseCoord(refPick.anchor);
    const p2 = parseCoord(refPick.focus);
    if (!p1 || !p2) return;
    const a = `${indexToCol(Math.min(p1.col, p2.col))}${Math.min(p1.row, p2.row)}`;
    const b = `${indexToCol(Math.max(p1.col, p2.col))}${Math.max(p1.row, p2.row)}`;
    const ref = a === b ? a : `${a}:${b}`;
    const input = refPick.input;
    input.value = input.value.slice(0, refPick.start) + ref + input.value.slice(refPick.end);
    refPick.end = refPick.start + ref.length;
    input.setSelectionRange(refPick.end, refPick.end);
    if (input !== formulaInput) formulaInput.value = input.value;
    else if (inlineEditor) inlineEditor.input.value = input.value;
    updateFormulaRefHighlight(input.value);
    highlightPickedReference(p1, p2);
  }

  function highlightPickedReference(p1, p2) {
    gridBody.querySelectorAll('td.ref-pick-cell').forEach(el => el.classList.remove('ref-pick-cell'));
    if (!p1 || !p2) return;
    for (let r = Math.min(p1.row, p2.row); r <= Math.max(p1.row, p2.row); r++) {
      for (let c = Math.min(p1.col, p2.col); c <= Math.max(p1.col, p2.col); c++) {
        const td = gridBody.querySelector(`td[data-coord="${indexToCol(c)}${r}"]`);
        if (td) td.classList.add('ref-pick-cell');
      }
    }
  }

  function endReferencePick() {
    refPick = null;
    highlightPickedReference(null, null);
  }

  // Excel-style colored outlines around every range referenced by the formula being edited
  const FORMULA_REF_COLORS = ['#1f6fd1', '#d13438', '#8e44ad', '#107c41', '#ca5010', '#00838f', '#c2185b'];
  const FORMULA_REF_RE = /(?<![A-Za-z0-9_$.])(?:(?:'([^']+)'|([A-Za-z_][\w.]*))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?(?![\w(])/g;

  function clearFormulaRefHighlight() {
    gridBody.querySelectorAll('td.formula-ref-cell').forEach(td => {
      td.classList.remove('formula-ref-cell');
      ['--ref-color', '--ref-bt', '--ref-br', '--ref-bb', '--ref-bl'].forEach(p => td.style.removeProperty(p));
    });
  }

  function updateFormulaRefHighlight(text) {
    clearFormulaRefHighlight();
    if (!text || !text.startsWith('=')) return;
    const withoutStrings = text.replace(/"[^"]*"/g, m => ' '.repeat(m.length));
    let colorIdx = 0;
    for (const m of withoutStrings.matchAll(FORMULA_REF_RE)) {
      const sheetName = m[1] || m[2];
      if (sheetName && sheetName !== activeSheet) continue;
      const p1 = parseCoord(`${m[3].toUpperCase()}${m[4]}`);
      const p2 = m[5] ? parseCoord(`${m[5].toUpperCase()}${m[6]}`) : p1;
      if (!p1 || !p2) continue;
      const color = FORMULA_REF_COLORS[colorIdx++ % FORMULA_REF_COLORS.length];
      const minR = Math.min(p1.row, p2.row), maxR = Math.max(p1.row, p2.row);
      const minC = Math.min(p1.col, p2.col), maxC = Math.max(p1.col, p2.col);
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const td = gridBody.querySelector(`td[data-coord="${indexToCol(c)}${r}"]`);
          if (!td) continue;
          td.classList.add('formula-ref-cell');
          td.style.setProperty('--ref-color', color);
          td.style.setProperty('--ref-bt', r === minR ? '2px' : '0');
          td.style.setProperty('--ref-bb', r === maxR ? '2px' : '0');
          td.style.setProperty('--ref-bl', c === minC ? '2px' : '0');
          td.style.setProperty('--ref-br', c === maxC ? '2px' : '0');
        }
      }
    }
  }

  function startInlineEdit(coord, initialText = null) {
    const td = gridBody.querySelector(`td[data-coord="${coord}"]`);
    if (!td) return;
    closeInlineEditor(false);

    const sheet = workbook.sheets[activeSheet];
    const cellObj = sheet?.cells[coord];
    const startVal = initialText !== null ? initialText : cellEditText(sheet, coord);

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-inline-editor';
    input.value = startVal;

    td.appendChild(input);
    input.focus();
    if (initialText === null) input.select();
    else input.setSelectionRange(input.value.length, input.value.length);

    inlineEditor = { coord, input };
    updateFormulaRefHighlight(input.value);

    input.addEventListener('input', () => {
      formulaInput.value = input.value;
      updateFormulaRefHighlight(input.value);
    });

    input.addEventListener('keydown', (e) => {
      e.stopPropagation();

      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
        e.preventDefault();
        const val = input.value;
        closeInlineEditor(false);
        setCellContent(activeSheet, coord, val);
        if (e.key === 'ArrowDown') moveSelection(1, 0);
        else if (e.key === 'ArrowUp') moveSelection(-1, 0);
        else if (e.key === 'ArrowRight') moveSelection(0, 1);
        else if (e.key === 'ArrowLeft') moveSelection(0, -1);
        return;
      }

      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        const val = input.value;
        closeInlineEditor(false);
        setCellContent(activeSheet, coord, val);
        if (e.key === 'Enter') moveSelection(1, 0);
        else moveSelection(0, e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        closeInlineEditor(false);
        selectCell(coord, false);
      }
    });

    input.addEventListener('blur', () => {
      setTimeout(() => {
        if (inlineEditor && inlineEditor.coord === coord) {
          const val = input.value;
          closeInlineEditor(false);
          setCellContent(activeSheet, coord, val);
        }
      }, 120);
    });
  }

  function closeInlineEditor(commit = false) {
    endReferencePick();
    clearFormulaRefHighlight();
    if (!inlineEditor) return;
    const { coord, input } = inlineEditor;
    const val = input.value;
    inlineEditor = null;
    if (input.parentNode) input.parentNode.removeChild(input);
    if (commit) setCellContent(activeSheet, coord, val);
  }

  function moveSelection(dRow, dCol, extendSelection = false) {
    const baseCoord = extendSelection ? (selectionEndCoord || selectedCoord) : selectedCoord;
    const parsed = parseCoord(baseCoord);
    if (!parsed) return;
    const sheet = workbook.sheets[activeSheet];
    let needsRender = false;

    // Step out of a merged block from its edge, like Excel
    const m = findMergeAt(parsed.coord);
    let fromRow = parsed.row, fromCol = parsed.col;
    if (m) {
      if (dRow > 0) fromRow = m.r2; else if (dRow < 0) fromRow = m.r1;
      if (dCol > 0) fromCol = m.c2; else if (dCol < 0) fromCol = m.c1;
    }
    const targetRow = Math.max(1, fromRow + dRow);
    let targetCol = Math.max(1, fromCol + dCol);
    // move along the columns as displayed (skips hidden ones, follows expanded day columns)
    const visIdx = dCol ? view.visCols.indexOf(fromCol) : -1;
    if (visIdx >= 0 && view.visCols[visIdx + dCol] !== undefined) targetCol = view.visCols[visIdx + dCol];
    else if (visIdx >= 0 && dCol < 0) targetCol = fromCol;

    if (targetRow > (sheet.maxRow || 100)) {
      sheet.maxRow = targetRow + 25;
      needsRender = true;
    }
    if (targetCol > (sheet.maxCol || 20)) {
      sheet.maxCol = targetCol + 5;
      needsRender = true;
    }

    if (needsRender) renderSheet();
    selectCell(`${indexToCol(targetCol)}${targetRow}`, extendSelection);
  }

  // Row / Column Insert & Delete
  // ---- Excel-like row insert / delete: every reference to the moved rows follows them
  //   op = { sheet, at, n } inserts n rows before `at`; op = { sheet, d1, d2 } deletes rows d1..d2.
  //   Applies to the formulas of every sheet (unqualified refs of the same sheet, or Sheet!refs),
  //   and to the sheet's merges, row heights, hidden rows, filters, conditional-format ranges, pivots and tables.
  const RANGE_REF_RE = /(?<![A-Za-z0-9_.$!])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?)(\d+)(?::(\$?[A-Za-z]{1,3}\$?)(\d+))?(?![\w(])/g;
  const unquoteSheet = prefix => {
    const n = prefix.slice(0, -1);
    return n.startsWith("'") ? n.slice(1, -1).replace(/''/g, "'") : n;
  };
  function mapRowRange(r1, r2, op) {
    if (op.n) return [r1 >= op.at ? r1 + op.n : r1, r2 >= op.at ? r2 + op.n : r2];
    const n = op.d2 - op.d1 + 1;
    if (r1 >= op.d1 && r2 <= op.d2) return null; // the whole range was deleted
    const a = r1 < op.d1 ? r1 : (r1 > op.d2 ? r1 - n : op.d1);
    const b = r2 < op.d1 ? r2 : (r2 > op.d2 ? r2 - n : op.d1 - 1);
    return [a, b];
  }
  function adjustRowRefs(text, homeSheet, op) {
    if (!text) return text;
    return String(text).split(/("[^"]*")/).map(part => {
      if (part.startsWith('"')) return part;
      return part.replace(RANGE_REF_RE, (m, prefix, c1, r1, c2, r2) => {
        const target = prefix ? unquoteSheet(prefix) : homeSheet;
        if (target !== op.sheet) return m;
        if (c2 === undefined) {
          const res = mapRowRange(Number(r1), Number(r1), op);
          return res ? `${prefix || ''}${c1}${res[0]}` : '#REF!';
        }
        const res = mapRowRange(Number(r1), Number(r2), op);
        return res ? `${prefix || ''}${c1}${res[0]}:${c2}${res[1]}` : '#REF!';
      });
    }).join('');
  }
  // sheets whose formulas mention `sheetName` (for the undo snapshot)
  function sheetsReferencing(sheetName) {
    const names = new Set([sheetName]);
    for (const [name, sh] of Object.entries(workbook.sheets)) {
      if (names.has(name)) continue;
      for (const cell of Object.values(sh.cells || {})) {
        if (cell.f && cell.f.includes(sheetName)) { names.add(name); break; }
      }
    }
    return [...names];
  }
  function applyRowOp(op) {
    // 1) formulas everywhere (by their old positions)
    for (const [name, sh] of Object.entries(workbook.sheets)) {
      let changed = false;
      for (const cell of Object.values(sh.cells || {})) {
        if (!cell.f) continue;
        const f = adjustRowRefs(cell.f, name, op);
        if (f !== cell.f) { cell.f = f; changed = true; }
      }
      if (changed) markSheetDirty(name);
    }
    const sheet = workbook.sheets[op.sheet];
    const mapRow = r => { const res = mapRowRange(r, r, op); return res ? res[0] : null; };
    // 2) structures of the sheet itself
    const adjustList = list => (list || []).map(ref => adjustRowRefs(ref, op.sheet, op)).filter(ref => ref && !ref.includes('#REF!'));
    if (sheet.merges) sheet.merges = adjustList(sheet.merges);
    if (sheet.filters) sheet.filters.forEach(f => { if (f.ref) f.ref = adjustRowRefs(f.ref, op.sheet, op); });
    if (sheet.condFormats) {
      sheet.condFormats.forEach(cf => { cf.ranges = adjustList(cf.ranges); });
      sheet.condFormats = sheet.condFormats.filter(cf => cf.ranges.length);
    }
    for (const key of ['hiddenRows', 'filterHiddenRows']) {
      if (sheet[key]) sheet[key] = sheet[key].map(mapRow).filter(r => r !== null);
    }
    if (sheet.rowHeights) {
      const h = {};
      for (const [r, v] of Object.entries(sheet.rowHeights)) { const nr = mapRow(Number(r)); if (nr !== null) h[nr] = v; }
      sheet.rowHeights = h;
    }
    if (sheet.styleBounds) {
      const b = sheet.styleBounds.rows;
      sheet.styleBounds.rows = op.n ? (b >= op.at ? b + op.n : b) : (b > op.d2 ? b - (op.d2 - op.d1 + 1) : Math.min(b, op.d1 - 1));
    }
    for (const t of Object.values(workbook.tables || {})) if (t.sheet === op.sheet && t.ref) t.ref = adjustRowRefs(t.ref, op.sheet, op);
    for (const pv of workbook.pivots || []) {
      if (pv.source && pv.source.sheet === op.sheet && pv.source.ref) pv.source.ref = adjustRowRefs(pv.source.ref, op.sheet, op);
      if (pv.sheet === op.sheet) for (const k of ['location', 'area', 'anchor']) if (pv[k]) pv[k] = adjustRowRefs(pv[k], op.sheet, op);
    }
  }

  function insertRowsAt(insertRow, count = 1) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet || count < 1) return;
    pushUndoState(sheetsReferencing(activeSheet));
    applyRowOp({ sheet: activeSheet, at: insertRow, n: count });

    const maxR = Math.max(sheet.maxRow || 100, getLastRealRow(activeSheet) + 5);
    for (let r = maxR; r >= insertRow; r--) {
      for (let c = 1; c <= (sheet.maxCol || 20); c++) {
        const colStr = indexToCol(c);
        const srcCell = sheet.cells[`${colStr}${r}`];
        if (srcCell) {
          sheet.cells[`${colStr}${r + count}`] = JSON.parse(JSON.stringify(srcCell));
          delete sheet.cells[`${colStr}${r}`];
        } else {
          delete sheet.cells[`${colStr}${r + count}`];
        }
      }
    }

    sheet.maxRow = Math.max(sheet.maxRow || 100, maxR + count + 20);
    engine.recalculateAll();
    renderSheet();
    selectCell(`A${insertRow}`, false, false); // keep the view where it is
    saveWorkbookState();
    showToast(`${count} row(s) inserted at row ${insertRow}`);
  }

  function deleteSelectedRows() {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    const { minRow, maxRow } = getSelectionBounds();
    const count = maxRow - minRow + 1;

    pushUndoState(sheetsReferencing(activeSheet));
    applyRowOp({ sheet: activeSheet, d1: minRow, d2: maxRow });
    const maxR = Math.max(sheet.maxRow || 100, getLastRealRow(activeSheet) + 5);
    for (let r = minRow; r <= maxR; r++) {
      for (let c = 1; c <= (sheet.maxCol || 20); c++) {
        const colStr = indexToCol(c);
        const srcCell = sheet.cells[`${colStr}${r + count}`];
        if (srcCell) {
          sheet.cells[`${colStr}${r}`] = JSON.parse(JSON.stringify(srcCell));
        } else {
          delete sheet.cells[`${colStr}${r}`];
        }
      }
    }

    engine.recalculateAll();
    renderSheet();
    selectCell(`A${minRow}`, false, false);
    saveWorkbookState();
    showToast(`${count} row(s) deleted`);
  }

  function expandGridDimensions(extraRows = 0, extraCols = 0) {
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;
    if (extraRows > 0) sheet.maxRow = (sheet.maxRow || 100) + extraRows;
    if (extraCols > 0) sheet.maxCol = (sheet.maxCol || 20) + extraCols;
    renderSheet();
    saveWorkbookState();
    showToast(
      extraRows > 0
        ? `Added ${extraRows} more rows at the bottom (Total: ${sheet.maxRow} rows)`
        : `Added ${extraCols} more columns to the right (Up to column ${indexToCol(sheet.maxCol)})`
    );
  }

  // Fill Handle Preview & Commit
  function updateFillHandlePreview(hoverCoord) {
    gridBody.querySelectorAll('td.autofill-preview-cell').forEach(el => el.classList.remove('autofill-preview-cell'));
    const pHover = parseCoord(hoverCoord);
    if (!pHover) return;
    fillPreviewTargetCoord = hoverCoord;

    const bounds = getSelectionBounds();
    const minR = Math.min(bounds.minRow, pHover.row);
    const maxR = Math.max(bounds.maxRow, pHover.row);
    const minC = Math.min(bounds.minCol, pHover.col);
    const maxC = Math.max(bounds.maxCol, pHover.col);

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (r >= bounds.minRow && r <= bounds.maxRow && c >= bounds.minCol && c <= bounds.maxCol) continue;
        const td = gridBody.querySelector(`td[data-coord="${indexToCol(c)}${r}"]`);
        if (td) td.classList.add('autofill-preview-cell');
      }
    }
  }

  function commitFillHandleDrag() {
    if (!fillPreviewTargetCoord) return;
    const pHover = parseCoord(fillPreviewTargetCoord);
    fillPreviewTargetCoord = null;
    if (!pHover) return;

    const bounds = getSelectionBounds();
    const sheet = workbook.sheets[activeSheet];
    if (!sheet) return;

    const srcRows = bounds.maxRow - bounds.minRow + 1;
    const srcCols = bounds.maxCol - bounds.minCol + 1;

    const minR = Math.min(bounds.minRow, pHover.row);
    const maxR = Math.max(bounds.maxRow, pHover.row);
    const minC = Math.min(bounds.minCol, pHover.col);
    const maxC = Math.max(bounds.maxCol, pHover.col);

    if (minR === bounds.minRow && maxR === bounds.maxRow && minC === bounds.minCol && maxC === bounds.maxCol) {
      return;
    }

    pushUndoState();
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (r >= bounds.minRow && r <= bounds.maxRow && c >= bounds.minCol && c <= bounds.maxCol) continue;
        const sr = bounds.minRow + ((r - bounds.minRow) % srcRows + srcRows) % srcRows;
        const sc = bounds.minCol + ((c - bounds.minCol) % srcCols + srcCols) % srcCols;
        copyCellWithFormat(sheet, sr, sc, r, c);
      }
    }

    engine.recalculateAll();
    renderSheet();
    selectedCoord = `${indexToCol(minC)}${minR}`;
    selectionEndCoord = `${indexToCol(maxC)}${maxR}`;
    highlightSelectedCellInDOM();
    updateStatusMetrics();
    saveWorkbookState();
    showToast(`Auto-fill applied (${indexToCol(minC)}${minR}:${indexToCol(maxC)}${maxR})`);
  }

  async function switchSheet(sheetName) {
    if (!workbook.sheets[sheetName]) return;
    if (!Object.prototype.hasOwnProperty.call(workbook.sheets[sheetName], 'cells')) {
      const requestedWorkbook = workbook;
      const requestedKey = WB_KEY;
      try {
        const sheet = await loadSheetFromNeon(requestedKey, sheetName);
        if (requestedWorkbook !== workbook || requestedKey !== WB_KEY) return;
        workbook.sheets[sheetName] = sheet;
        workbookCache.set(WB_KEY, workbook);
        engine = new window.FormulaEngine(workbook);
        engine.recalculateAll();
      } catch (error) {
        showToast(error.message || `No se pudo cargar ${sheetName}.`, true);
        return;
      }
    }
    activeSheet = sheetName;
    renderSheetTabs();
    renderSheet();
    selectCell('A1', false);
    sheetTabsList.querySelector('.sheet-tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  let workspaceSwitchPromise = null;
  async function switchWorkbookApartado(key, updateHistory = true) {
    if (!VALID_WB_KEYS.includes(key) || key === WB_KEY) return;
    if (workspaceSwitchPromise) return workspaceSwitchPromise;

    workspaceSwitchPromise = (async () => {
      const previousKey = WB_KEY;
      await flushPendingSave();
      clearTimeout(actualizarDataPollTimer);
      clearTimeout(actualizarInventarioPollTimer);
      actualizarDataPollTimer = null;
      actualizarInventarioPollTimer = null;
      closeInlineEditor(true);

      let nextWorkbook = workbookCache.get(key);
      if (!nextWorkbook) nextWorkbook = await prefetchWorkbook(key);
      if (!nextWorkbook) {
        setWorkbookKey(previousKey);
        showToast(`No se pudo cargar ${key}.`, true);
        return;
      }

      undoStack.length = 0;
      redoStack.length = 0;
      dirtySheets.clear();
      activateWorkbook(nextWorkbook, key);
      updateUndoRedoButtons();
      if (updateHistory) history.pushState({ wb: key }, '', `/?wb=${key}`);
      if (WB_KEY === 'reunion') {
        pollActualizarData();
        pollActualizarInventario();
      }
    })().finally(() => { workspaceSwitchPromise = null; });

    return workspaceSwitchPromise;
  }

  function bindWorkspaceNavigation() {
    document.querySelectorAll('.wb-switch [data-wb]').forEach(link => {
      link.addEventListener('mouseenter', () => prefetchWorkbook(link.dataset.wb));
      link.addEventListener('focus', () => prefetchWorkbook(link.dataset.wb));
      link.addEventListener('click', event => {
        event.preventDefault();
        switchWorkbookApartado(link.dataset.wb);
      });
    });
    window.addEventListener('popstate', () => switchWorkbookApartado(wbFromUrl(), false));
  }

  function addNewSheet() {
    pushUndoState();
    let idx = (workbook.sheetOrder?.length || 1) + 1;
    let newName = `Sheet${idx}`;
    while (workbook.sheets[newName]) {
      idx++;
      newName = `Sheet${idx}`;
    }
    workbook.sheets[newName] = {
      maxRow: 100,
      maxCol: 20,
      colWidths: {},
      merges: [],
      cells: {}
    };
    workbook.sheetOrder.push(newName);
    switchSheet(newName);
    saveWorkbookState();
    showToast(`New sheet "${newName}" created`);
  }

  function renameSheet(oldName) {
    const nextName = prompt('New sheet name:', oldName);
    if (!nextName) return;
    const trimmed = nextName.trim();
    if (!trimmed || trimmed === oldName) return;
    if (workbook.sheets[trimmed]) {
      showToast('A sheet with that name already exists', true);
      return;
    }
    // like Excel, formulas that point at the old name are updated
    const esc = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const refRe = new RegExp(`(^|[^A-Za-z0-9_.'])(?:'${esc.replace(/'/g, "''")}'|${esc})!`, 'g');
    const needsQuotes = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed);
    const newRef = needsQuotes ? `'${trimmed.replace(/'/g, "''")}'!` : `${trimmed}!`;
    const touched = Object.keys(workbook.sheets).filter(n => Object.values(workbook.sheets[n].cells || {}).some(c => c && c.f && refRe.test(c.f) && ((refRe.lastIndex = 0) || true)));
    pushUndoState([oldName, ...touched]);
    for (const n of touched) {
      for (const cell of Object.values(workbook.sheets[n].cells)) {
        if (cell && cell.f) { refRe.lastIndex = 0; cell.f = cell.f.replace(refRe, (m, pre) => `${pre}${newRef}`); }
      }
      markSheetDirty(n);
    }
    for (const t of Object.values(workbook.tables || {})) if (t.sheet === oldName) t.sheet = trimmed;
    if (window.MauriPivot) window.MauriPivot.onSheetRenamed(oldName, trimmed);
    workbook.sheets[trimmed] = workbook.sheets[oldName];
    delete workbook.sheets[oldName];
    const idx = workbook.sheetOrder.indexOf(oldName);
    if (idx !== -1) workbook.sheetOrder[idx] = trimmed;
    if (activeSheet === oldName) activeSheet = trimmed;
    markSheetDirty(oldName);
    markSheetDirty(trimmed);
    // "P Cosecha YYWW": title and week columns follow the new name (weeks.js)
    if (window.MauriWeeks) window.MauriWeeks.onSheetRenamed(oldName, trimmed);
    engine.recalculateAll();
    renderSheetTabs();
    renderSheet();
    updateStatusMetrics();
    saveWorkbookState();
    showToast(`Sheet renamed to "${trimmed}"`);
  }

  function deleteSheet(sheetName) {
    const order = workbook.sheetOrder || Object.keys(workbook.sheets);
    if (visibleSheets().length <= 1 && !workbook.sheets[sheetName]?.hidden) {
      showToast('A workbook must contain at least one visible worksheet', true);
      return;
    }
    if (!workbook.sheets[sheetName]) return;

    const confirmed = confirm(`⚠️ Are you sure you want to permanently delete sheet "${sheetName}"?`);
    if (!confirmed) return;

    pushUndoState([sheetName]);
    markSheetDirty(sheetName);
    const idx = order.indexOf(sheetName);
    delete workbook.sheets[sheetName];
    if (window.MauriPivot) window.MauriPivot.onSheetDeleted(sheetName);
    workbook.sheetOrder = order.filter(n => n !== sheetName);

    if (activeSheet === sheetName) {
      const vis = visibleSheets();
      activeSheet = vis[Math.min(Math.max(0, idx), vis.length - 1)] || vis[0];
    }

    engine.recalculateAll();
    renderSheetTabs();
    renderSheet();
    selectCell('A1', false);
    saveWorkbookState();
    showToast(`Sheet "${sheetName}" deleted (Ctrl+Z to undo)`);
  }

  function duplicateSheet(sheetName) {
    const src = workbook.sheets[sheetName];
    if (!src) return;
    pushUndoState();
    let copyIdx = 2;
    let newName = `${sheetName} (${copyIdx})`;
    while (workbook.sheets[newName]) {
      copyIdx++;
      newName = `${sheetName} (${copyIdx})`;
    }
    workbook.sheets[newName] = JSON.parse(JSON.stringify(src));
    const idx = workbook.sheetOrder.indexOf(sheetName);
    if (idx !== -1) {
      workbook.sheetOrder.splice(idx + 1, 0, newName);
    } else {
      workbook.sheetOrder.push(newName);
    }
    switchSheet(newName);
    saveWorkbookState();
    showToast(`Sheet "${sheetName}" duplicated as "${newName}"`);
  }

  // the status bar (count / average / sum) is computed once things settle, not on every drag step
  let statusTimer = null;
  function updateStatusMetrics() {
    clearTimeout(statusTimer);
    statusTimer = setTimeout(computeStatusMetrics, 60);
  }
  function computeStatusMetrics() {
    const bounds = getSelectionBounds();
    const sheet = workbook.sheets[activeSheet];
    let nonEmptyCount = 0;
    let numericCount = 0;
    let numericSum = 0;

    // like Excel's status bar, only visible cells count (filtered / hidden rows and columns are skipped)
    const hiddenR = view.sheetName === activeSheet && view.hiddenRows ? view.hiddenRows : new Set();
    const hiddenC = view.sheetName === activeSheet && view.hiddenCols ? view.hiddenCols : new Set();
    for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
      if (hiddenR.has(r)) continue;
      for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
        if (hiddenC.has(indexToCol(c))) continue;
        const val = sheet?.cells[`${indexToCol(c)}${r}`]?.v;
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          nonEmptyCount++;
          const num = Number(val);
          if (!isNaN(num)) {
            numericCount++;
            numericSum += num;
          }
        }
      }
    }

    if (bounds.isMulti) {
      const rangeStr = `${indexToCol(bounds.minCol)}${bounds.minRow}:${indexToCol(bounds.maxCol)}${bounds.maxRow}`;
      document.getElementById('metricCount').textContent = `Range: ${activeSheet}!${rangeStr} (Count: ${nonEmptyCount})`;
    } else {
      document.getElementById('metricCount').textContent = `Cell: ${activeSheet}!${selectedCoord}`;
    }

    const avg = numericCount > 0 ? numericSum / numericCount : 0;
    document.getElementById('metricAvg').textContent = `Average: ${avg.toLocaleString('en-US', { maximumFractionDigits: 2 })}`;
    document.getElementById('metricSum').textContent = `Sum: ${numericSum.toLocaleString('en-US', { maximumFractionDigits: 2 })}`;
  }

  function showToast(msg, isWarn = false) {
    const container = document.getElementById('toastContainer');
    const el = document.createElement('div');
    el.className = `excel-toast${isWarn ? ' warn' : ''}`;
    el.innerHTML = `<span>${escapeHtml(msg)}</span>`;
    container.appendChild(el);
    setTimeout(() => {
      if (el.parentNode) el.remove();
    }, 3800);
  }

  function exportActiveSheetCsv() {
    const sheet = workbook.sheets[activeSheet];
    const lastR = getLastRealRow(activeSheet);
    const maxC = sheet.maxCol || 20;
    const lines = [];
    for (let r = 1; r <= lastR; r++) {
      const rowVals = [];
      for (let c = 1; c <= maxC; c++) {
        const coord = `${indexToCol(c)}${r}`;
        const v = sheet.cells[coord]?.v ?? '';
        rowVals.push(`"${String(v).replace(/"/g, '""')}"`);
      }
      lines.push(rowVals.join(','));
    }
    const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${activeSheet}_FormatoMauri.csv`;
    a.click();
    showToast(`Sheet "${activeSheet}" exported to CSV`);
  }

  // Ribbon collapse (Excel's "^" / Ctrl+F1); remembered per browser
  function setRibbonCollapsed(collapsed) {
    document.body.classList.toggle('ribbon-collapsed', collapsed);
    const btn = document.getElementById('btnToggleRibbon');
    if (btn) {
      btn.classList.toggle('checked', collapsed);
      btn.querySelector('.chk').textContent = collapsed ? '✓' : '';
    }
    try { localStorage.setItem('reunion123_ribbon_collapsed', collapsed ? '1' : '0'); } catch (e) { /* ignore */ }
    // the grid area changed size: refresh the rendered rows
    if (activeSheet && workbook) onViewportScroll();
  }

  // "☰" list of every visible sheet (workbooks with many tabs)
  let sheetListMenu = null;
  function closeSheetList() { if (sheetListMenu) { sheetListMenu.remove(); sheetListMenu = null; } }
  function openSheetList(anchor) {
    closeSheetList();
    sheetListMenu = document.createElement('div');
    sheetListMenu.className = 'sheet-list-menu no-print';
    const names = visibleSheets();
    sheetListMenu.innerHTML = `<input type="search" placeholder="Buscar hoja (${names.length})" autocomplete="off" />
      <div class="sheet-list-items">${names.map(n => `<button type="button" data-sheet="${escapeHtml(n)}"${n === activeSheet ? ' class="active"' : ''}>${escapeHtml(n)}</button>`).join('')}</div>`;
    document.body.appendChild(sheetListMenu);
    const r = anchor.getBoundingClientRect();
    sheetListMenu.style.left = `${Math.max(8, r.left)}px`;
    sheetListMenu.style.bottom = `${Math.max(8, window.innerHeight - r.top + 4)}px`;
    const input = sheetListMenu.querySelector('input');
    input.addEventListener('input', () => {
      const q = input.value.trim().toLowerCase();
      sheetListMenu.querySelectorAll('.sheet-list-items button').forEach(b => {
        b.style.display = !q || b.dataset.sheet.toLowerCase().includes(q) ? '' : 'none';
      });
    });
    input.addEventListener('keydown', e => {
      if (e.key === 'Escape') closeSheetList();
      if (e.key === 'Enter') {
        const first = [...sheetListMenu.querySelectorAll('.sheet-list-items button')].find(b => b.style.display !== 'none');
        if (first) { switchSheet(first.dataset.sheet); closeSheetList(); }
      }
    });
    sheetListMenu.addEventListener('click', e => {
      const b = e.target.closest('[data-sheet]');
      if (!b) return;
      switchSheet(b.dataset.sheet);
      closeSheetList();
    });
    sheetListMenu.querySelector('.active')?.scrollIntoView({ block: 'center' });
    setTimeout(() => input.focus(), 0);
  }

  function bindEvents() {
    document.getElementById('btnSheetList').addEventListener('click', (e) => {
      if (sheetListMenu) closeSheetList(); else openSheetList(e.currentTarget);
    });
    window.addEventListener('mousedown', (e) => {
      if (sheetListMenu && !sheetListMenu.contains(e.target) && e.target.id !== 'btnSheetList') closeSheetList();
    });

    let startCollapsed = false;
    try { startCollapsed = localStorage.getItem('reunion123_ribbon_collapsed') === '1'; } catch (e) { /* ignore */ }
    setRibbonCollapsed(startCollapsed);
    document.getElementById('btnToggleRibbon').addEventListener('click', () => {
      setRibbonCollapsed(!document.body.classList.contains('ribbon-collapsed'));
    });

    // "Ocultar" menu (flowers in 0 / tools)
    const hideMenu = document.getElementById('hideMenu');
    document.getElementById('btnHideMenu').addEventListener('click', (e) => {
      e.stopPropagation();
      hideMenu.hidden = !hideMenu.hidden;
    });
    hideMenu.addEventListener('click', (e) => {
      if (e.target.closest('.hide-menu-item')) hideMenu.hidden = true;
    });
    document.addEventListener('mousedown', (e) => {
      if (!hideMenu.hidden && !e.target.closest('.hide-menu-wrap')) hideMenu.hidden = true;
    });
    window.addEventListener('keydown', (e) => {
      if (e.ctrlKey && e.key === 'F1') {
        e.preventDefault();
        setRibbonCollapsed(!document.body.classList.contains('ribbon-collapsed'));
      }
    }, true);

    // Autofilter dropdown buttons (captured before the grid's own selection logic)
    gridBody.addEventListener('mousedown', (e) => {
      const btn = e.target.closest('.af-btn');
      if (!btn) return;
      e.preventDefault();
      e.stopPropagation();
      closeInlineEditor(true);
      if (afMenu && afMenu.dataset.for === btn.dataset.af) { closeFilterMenu(); return; }
      openFilterMenu(btn);
      if (afMenu) afMenu.dataset.for = btn.dataset.af;
    }, true);
    window.addEventListener('mousedown', (e) => {
      if (afMenu && !afMenu.contains(e.target) && !e.target.closest('.af-btn')) closeFilterMenu();
    });
    window.addEventListener('keydown', (e) => {
      if (afMenu && e.key === 'Escape') { e.stopPropagation(); closeFilterMenu(); }
    }, true);

    // Column Header Click -> Select Entire Column
    gridHead.addEventListener('mousedown', (e) => {
      const cornerTh = e.target.closest('th.corner-header');
      if (cornerTh) {
        e.preventDefault();
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 20);
        const maxR = Math.max(getLastRealRow(activeSheet), sheet.maxRow || 100);
        selectRange('A1', `${maxColStr}${maxR}`, true);
        return;
      }
      const colTh = e.target.closest('th[data-col]');
      if (!colTh) return;
      e.preventDefault();
      closeInlineEditor(true);
      const colStr = colTh.getAttribute('data-col');
      const sheet = workbook.sheets[activeSheet];
      const maxR = Math.max(getLastRealRow(activeSheet), sheet.maxRow || 100);
      if (e.shiftKey) {
        const startP = parseCoord(selectedCoord) || { colStr: 'A' };
        selectRange(`${startP.colStr}1`, `${colStr}${maxR}`, true);
      } else {
        selectRange(`${colStr}1`, `${colStr}${maxR}`, true);
      }
    });

    // Grid click, drag-to-select, fill handle drag & row header click
    gridBody.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;

      const fillHandle = e.target.closest('.cell-fill-handle');
      if (fillHandle) {
        e.preventDefault();
        e.stopPropagation();
        closeInlineEditor(true);
        isDraggingFillHandle = true;
        fillPreviewTargetCoord = null;
        return;
      }

      const rowTh = e.target.closest('th.row-header[data-row]');
      if (rowTh) {
        e.preventDefault();
        closeInlineEditor(true);
        const r = Number(rowTh.getAttribute('data-row'));
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 20);
        if (e.shiftKey) {
          const startP = parseCoord(selectedCoord) || { row: r };
          selectRange(`A${startP.row}`, `${maxColStr}${r}`, true);
        } else {
          selectRange(`A${r}`, `${maxColStr}${r}`, true);
        }
        return;
      }

      const td = e.target.closest('td[data-coord]');
      if (!td) return;
      const coord = td.getAttribute('data-coord');

      const editingInput = getFormulaEditingInput();
      if (editingInput && canInsertReference(editingInput) && !(inlineEditor && e.target === inlineEditor.input)) {
        e.preventDefault();
        const continuing = refPick && refPick.input === editingInput && editingInput.selectionStart === refPick.end;
        if (!continuing) {
          const pos = editingInput.selectionStart;
          refPick = { input: editingInput, start: pos, end: pos };
        }
        refPick.anchor = e.shiftKey && continuing ? refPick.anchor : coord;
        refPick.focus = coord;
        refPick.dragging = true;
        writePickedReference();
        return;
      }

      if (inlineEditor && inlineEditor.coord !== coord) {
        closeInlineEditor(true);
      }
      if (!inlineEditor) {
        e.preventDefault();
        if (document.activeElement === formulaInput || document.activeElement === cellNameBox) document.activeElement.blur();
        isDraggingSelection = true;
        selectCell(coord, e.shiftKey);
      }
    });

    const dragActive = () => isDraggingSelection || isDraggingFillHandle || !!(refPick && refPick.dragging);
    const dragOverCell = (td) => {
      if (!td) return;
      const coord = td.getAttribute('data-coord');
      if (refPick && refPick.dragging) {
        if (coord !== refPick.focus) {
          refPick.focus = coord;
          writePickedReference();
        }
        return;
      }
      if (isDraggingFillHandle) {
        updateFillHandlePreview(coord);
      } else if (isDraggingSelection) {
        if (selectionEndCoord !== coord) {
          selectCell(coord, true);
        }
      }
    };

    gridBody.addEventListener('mouseover', (e) => {
      if (!dragActive()) return;
      dragOverCell(e.target.closest('td[data-coord]'));
    });

    // Excel-like auto-scroll while dragging past the edge of the scrollable area
    let dragPointer = null;
    let dragScrollTimer = null;
    const stopDragScroll = () => {
      if (dragScrollTimer) clearInterval(dragScrollTimer);
      dragScrollTimer = null;
      dragPointer = null;
    };
    const dragAutoScroll = () => {
      if (!dragActive() || !dragPointer) { stopDragScroll(); return; }
      const vp = document.getElementById('spreadsheetViewport');
      const r = vp.getBoundingClientRect();
      const head = gridHead;
      const frozenRow = gridBody.querySelector('td.frozen-row-edge');
      const frozenCol = gridBody.querySelector('td.frozen-col-edge');
      const rowHead = gridBody.querySelector('th.row-header');
      const top = frozenRow ? frozenRow.getBoundingClientRect().bottom : (head ? head.getBoundingClientRect().bottom : r.top);
      const left = frozenCol ? frozenCol.getBoundingClientRect().right : (rowHead ? rowHead.getBoundingClientRect().right : r.left);
      const right = r.left + vp.clientWidth;
      const bottom = r.top + vp.clientHeight;
      const { x, y } = dragPointer;
      const speed = d => Math.min(80, Math.max(8, d));
      let dx = 0, dy = 0;
      if (x > right - 12) dx = speed(x - (right - 12));
      else if (x < left + 4 && vp.scrollLeft > 0) dx = -speed(left + 4 - x);
      if (y > bottom - 12) dy = speed(y - (bottom - 12));
      else if (y < top + 4 && vp.scrollTop > 0) dy = -speed(top + 4 - y);
      if (!dx && !dy) return;
      vp.scrollLeft += dx;
      vp.scrollTop += dy;
      const cx = Math.min(Math.max(x, left + 3), right - 3);
      const cy = Math.min(Math.max(y, top + 3), bottom - 3);
      setTimeout(() => {
        const el = document.elementFromPoint(cx, cy);
        dragOverCell(el && el.closest ? el.closest('td[data-coord]') : null);
      }, 0);
    };
    window.addEventListener('mousemove', (e) => {
      if (!dragActive()) return;
      dragPointer = { x: e.clientX, y: e.clientY };
      if (!dragScrollTimer) dragScrollTimer = setInterval(dragAutoScroll, 40);
    });

    window.addEventListener('mouseup', () => {
      stopDragScroll();
      if (refPick && refPick.dragging) {
        refPick.dragging = false;
        refPick.input.focus();
        refPick.input.setSelectionRange(refPick.end, refPick.end);
      }
      if (isDraggingFillHandle) {
        isDraggingFillHandle = false;
        commitFillHandleDrag();
      }
      if (isDraggingSelection) {
        isDraggingSelection = false;
        if (formatPainterActive) applyFormatPainterToSelection();
      }
    });

    gridBody.addEventListener('dblclick', (e) => {
      const td = e.target.closest('td[data-coord]');
      if (!td) return;
      startInlineEdit(td.getAttribute('data-coord'));
    });

    // Formula Bar & Name Box
    cellNameBox.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        e.preventDefault();
        const raw = cellNameBox.value.trim().toUpperCase();
        if (raw.includes(':')) {
          const [startStr, endStr] = raw.split(':');
          selectRange(startStr, endStr);
        } else {
          selectCell(raw, false);
        }
        cellNameBox.blur();
      }
    });

    formulaInput.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        e.preventDefault();
        endReferencePick();
        const val = formulaInput.value;
        const target = selectedCoord;
        if (inlineEditor) closeInlineEditor(false);
        setCellContent(activeSheet, target, val);
        formulaInput.blur();
        moveSelection(1, 0);
      } else if (e.key === 'Escape') {
        endReferencePick();
        if (inlineEditor) closeInlineEditor(false);
        selectCell(selectedCoord, false);
        formulaInput.blur();
      }
    });

    formulaInput.addEventListener('input', () => {
      if (inlineEditor) inlineEditor.input.value = formulaInput.value;
      updateFormulaRefHighlight(formulaInput.value);
    });
    formulaInput.addEventListener('focus', () => updateFormulaRefHighlight(formulaInput.value));
    formulaInput.addEventListener('blur', () => {
      setTimeout(() => {
        if (!inlineEditor && document.activeElement !== formulaInput) clearFormulaRefHighlight();
      }, 0);
    });

    document.getElementById('formatToolbar').addEventListener('mousedown', (e) => {
      if (inlineEditor && !e.target.closest('select')) {
        e.preventDefault();
        closeInlineEditor(true);
      }
    });

    // System Clipboard Copy / Cut
    const onNativeCopy = (isCut) => (e) => {
      if (document.body.classList.contains('modal-open')) return;
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'SELECT' || activeTag === 'TEXTAREA') return;
      copySelection(isCut);
      if (e.clipboardData && clipboardState) {
        e.preventDefault();
        e.clipboardData.setData('text/plain', clipboardState.tsv);
      }
    };
    window.addEventListener('copy', onNativeCopy(false));
    window.addEventListener('cut', onNativeCopy(true));

    // System Clipboard Paste
    window.addEventListener('paste', (e) => {
      if (document.body.classList.contains('modal-open')) return;
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'SELECT' || activeTag === 'TEXTAREA') return;

      e.preventDefault();
      const pastedText = (e.clipboardData || window.clipboardData)?.getData('text/plain') || '';
      if (clipboardState && (!pastedText || normalizeClipboardText(pastedText) === normalizeClipboardText(clipboardState.tsv))) {
        pasteInternalClipboard();
      } else if (pastedText) {
        pasteDataMatrix(normalizeClipboardText(pastedText).split('\n').map(line => line.split('\t')));
      }
    });

    // Global Keyboard Navigation & Excel Shortcuts
    window.addEventListener('keydown', (e) => {
      if (document.body.classList.contains('modal-open')) return;
      const activeTag = document.activeElement?.tagName;
      if (activeTag === 'INPUT' || activeTag === 'SELECT' || activeTag === 'TEXTAREA') return;

      const keyLow = e.key.toLowerCase();

      if ((e.ctrlKey || e.metaKey) && keyLow === 'z') {
        e.preventDefault();
        undo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'y') {
        e.preventDefault();
        redo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (keyLow === 'c' || keyLow === 'x')) {
        return; // handled by the native 'copy' / 'cut' events
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'v') {
        return; // handled by the native 'paste' event (no clipboard permission needed)
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'd') {
        e.preventDefault();
        fillSelectionDown();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'r') {
        e.preventDefault();
        fillSelectionRight();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'a') {
        e.preventDefault();
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 20);
        const maxR = Math.max(getLastRealRow(activeSheet), sheet.maxRow || 100);
        selectRange('A1', `${maxColStr}${maxR}`, true);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 's') {
        e.preventDefault();
        saveWorkbookState();
        showToast('Workbook saved');
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'b') {
        e.preventDefault();
        toggleBoldSelection();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'i') {
        e.preventDefault();
        toggleItalicSelection();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && keyLow === 'u') {
        e.preventDefault();
        toggleUnderlineSelection();
        return;
      }
      if (e.key === 'Escape' && formatPainterActive) {
        e.preventDefault();
        stopFormatPainter();
        return;
      }

      if (e.key === 'Escape') {
        if (clipboardState) {
          clipboardState = null;
          highlightSelectedCellInDOM();
          showToast('Copy selection cleared');
        } else {
          selectCell(selectedCoord, false);
        }
        return;
      }

      if (e.key === 'ArrowDown') { e.preventDefault(); moveSelection(1, 0, e.shiftKey); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); moveSelection(-1, 0, e.shiftKey); }
      else if (e.key === 'ArrowRight') { e.preventDefault(); moveSelection(0, 1, e.shiftKey); }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); moveSelection(0, -1, e.shiftKey); }
      else if (e.key === 'Tab') { e.preventDefault(); moveSelection(0, e.shiftKey ? -1 : 1, false); }
      else if ((e.ctrlKey || e.metaKey) && (e.key === '+' || e.key === '=')) {
        e.preventDefault();
        const { minRow } = getSelectionBounds();
        insertRowsAt(minRow, 1);
      } else if ((e.ctrlKey || e.metaKey) && e.key === '-') {
        e.preventDefault();
        deleteSelectedRows();
      } else if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault();
        startInlineEdit(selectedCoord);
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        clearSelectedRange();
      } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        startInlineEdit(selectedCoord, e.key);
      }
    });

    // Infinite scroll-to-bottom auto-expansion
    const viewportEl = document.getElementById('spreadsheetViewport');
    if (viewportEl) {
      viewportEl.addEventListener('scroll', onViewportScroll, { passive: true });
      viewportEl.addEventListener('scroll', () => {
        if (viewportEl.scrollTop + viewportEl.clientHeight >= viewportEl.scrollHeight - 25) {
          const sheet = workbook?.sheets?.[activeSheet];
          if (sheet && (sheet.maxRow || 100) < 2500) {
            sheet.maxRow = (sheet.maxRow || 100) + 150; // fewer, larger extensions while scrolling down
            renderSheet();
          }
        }
      });
    }

    // Excel Right-Click Context Menu
    const ctxMenu = document.createElement('div');
    ctxMenu.className = 'excel-context-menu no-print';
    document.body.appendChild(ctxMenu);

    const hideCtxMenu = () => {
      ctxMenu.style.display = 'none';
    };

    window.addEventListener('mousedown', (e) => {
      if (!ctxMenu.contains(e.target)) hideCtxMenu();
    });

    gridBody.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      closeInlineEditor(true);

      const rowTh = e.target.closest('th.row-header[data-row]');
      const td = e.target.closest('td[data-coord]');
      if (rowTh) {
        const r = Number(rowTh.getAttribute('data-row'));
        const sheet = workbook.sheets[activeSheet];
        const maxColStr = indexToCol(sheet.maxCol || 20);
        selectRange(`A${r}`, `${maxColStr}${r}`, true);
      } else if (td) {
        const coord = td.getAttribute('data-coord');
        const p = parseCoord(coord);
        const b = getSelectionBounds();
        if (p && (p.row < b.minRow || p.row > b.maxRow || p.col < b.minCol || p.col > b.maxCol)) {
          selectCell(coord, false);
        }
      }

      const bounds = getSelectionBounds();

      ctxMenu.innerHTML = `
        <button class="excel-context-item" data-ctx="copy">
          <span>📋 Copy</span><span class="excel-context-shortcut">Ctrl+C</span>
        </button>
        <button class="excel-context-item" data-ctx="cut">
          <span>✂️ Cut</span><span class="excel-context-shortcut">Ctrl+X</span>
        </button>
        <button class="excel-context-item" data-ctx="paste">
          <span>📥 Paste</span><span class="excel-context-shortcut">Ctrl+V</span>
        </button>
        <button class="excel-context-item" data-ctx="clear">
          <span>🧹 Clear Contents</span><span class="excel-context-shortcut">Del</span>
        </button>
        <div class="excel-context-divider"></div>
        <button class="excel-context-item" data-ctx="insert-above">
          <span>➕ Insert 1 Row Above (Row ${bounds.minRow})</span><span class="excel-context-shortcut">Ctrl++</span>
        </button>
        <button class="excel-context-item" data-ctx="insert-below">
          <span>➕ Insert 1 Row Below</span>
        </button>
        <button class="excel-context-item" data-ctx="insert-5">
          <span>➕ Insert 5 Rows Here</span>
        </button>
        <button class="excel-context-item danger" data-ctx="delete-row">
          <span>🗑️ Delete Selected Row(s)</span><span class="excel-context-shortcut">Ctrl+-</span>
        </button>
        <div class="excel-context-divider"></div>
        <button class="excel-context-item" data-ctx="add-50-rows">
          <span>⬇️ Add 50 More Rows at Bottom</span>
        </button>
        <button class="excel-context-item" data-ctx="add-5-cols">
          <span>➡️ Add 5 More Columns to Right</span>
        </button>
      `;

      ctxMenu.style.display = 'flex';
      const menuW = 265;
      const menuH = 340;
      ctxMenu.style.left = `${Math.min(window.innerWidth - menuW - 10, e.clientX)}px`;
      ctxMenu.style.top = `${Math.min(window.innerHeight - menuH - 10, e.clientY)}px`;
    });

    ctxMenu.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-ctx]');
      if (!btn) return;
      const action = btn.getAttribute('data-ctx');
      hideCtxMenu();
      const bounds = getSelectionBounds();

      if (action === 'copy') copySelection(false);
      else if (action === 'cut') copySelection(true);
      else if (action === 'paste') pasteFromClipboard();
      else if (action === 'clear') clearSelectedRange();
      else if (action === 'insert-above') insertRowsAt(bounds.minRow, 1);
      else if (action === 'insert-below') insertRowsAt(bounds.maxRow + 1, 1);
      else if (action === 'insert-5') insertRowsAt(bounds.minRow, 5);
      else if (action === 'delete-row') deleteSelectedRows();
      else if (action === 'add-50-rows') expandGridDimensions(50, 0);
      else if (action === 'add-5-cols') expandGridDimensions(0, 5);
      else if (action === 'sheet-add') addNewSheet();
      else if (action === 'sheet-rename') {
        const targetSheet = btn.getAttribute('data-target-sheet') || activeSheet;
        renameSheet(targetSheet);
      } else if (action === 'sheet-duplicate') {
        const targetSheet = btn.getAttribute('data-target-sheet') || activeSheet;
        duplicateSheet(targetSheet);
      } else if (action === 'sheet-delete') {
        const targetSheet = btn.getAttribute('data-target-sheet') || activeSheet;
        deleteSheet(targetSheet);
      } else if (action === 'sheet-hide' || action === 'sheet-unhide') {
        const targetSheet = btn.getAttribute('data-target-sheet');
        const target = workbook.sheets[targetSheet];
        if (!target) return;
        pushUndoState([targetSheet]);
        if (action === 'sheet-hide') {
          target.hidden = 'hidden';
          if (activeSheet === targetSheet) activeSheet = visibleSheets()[0];
        } else {
          delete target.hidden;
          activeSheet = targetSheet;
        }
        renderSheetTabs();
        renderSheet();
        selectCell('A1', false);
        markSheetDirty(targetSheet);
        saveWorkbookState();
      }
    });

    // Sheet Tabs Click, Close Button & Double-Click Rename
    // Drag a tab to reorder the sheets (before / after the tab you drop on, or at the end)
    let draggingTab = null;
    const clearTabDrop = () => sheetTabsList.querySelectorAll('.drop-before, .drop-after').forEach(el => el.classList.remove('drop-before', 'drop-after'));
    const tabDropTarget = (e) => {
      const tab = e.target.closest('.sheet-tab[data-sheet]');
      if (tab) {
        const r = tab.getBoundingClientRect();
        return { tab, after: e.clientX > r.left + r.width / 2 };
      }
      const tabs = sheetTabsList.querySelectorAll('.sheet-tab[data-sheet]');
      return tabs.length ? { tab: tabs[tabs.length - 1], after: true } : null;
    };
    sheetTabsList.addEventListener('dragstart', (e) => {
      const tab = e.target.closest('.sheet-tab[data-sheet]');
      if (!tab) return;
      draggingTab = tab.getAttribute('data-sheet');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', draggingTab);
      tab.classList.add('dragging');
    });
    sheetTabsList.addEventListener('dragover', (e) => {
      if (!draggingTab) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      const t = tabDropTarget(e);
      clearTabDrop();
      if (t && t.tab.getAttribute('data-sheet') !== draggingTab) t.tab.classList.add(t.after ? 'drop-after' : 'drop-before');
      // near the edges, scroll the tab strip
      const r = sheetTabsList.getBoundingClientRect();
      if (e.clientX < r.left + 30) sheetTabsList.scrollLeft -= 12;
      else if (e.clientX > r.right - 30) sheetTabsList.scrollLeft += 12;
    });
    sheetTabsList.addEventListener('drop', (e) => {
      if (!draggingTab) return;
      e.preventDefault();
      const t = tabDropTarget(e);
      const target = t && t.tab.getAttribute('data-sheet');
      if (target && target !== draggingTab) {
        const order = workbook.sheetOrder.filter(n => n !== draggingTab);
        order.splice(order.indexOf(target) + (t.after ? 1 : 0), 0, draggingTab);
        workbook.sheetOrder = order;
        renderSheetTabs();
        saveWorkbookState();
        showToast(`"${draggingTab}" movida`);
      }
    });
    sheetTabsList.addEventListener('dragend', () => {
      draggingTab = null;
      clearTabDrop();
      sheetTabsList.querySelectorAll('.sheet-tab.dragging').forEach(el => el.classList.remove('dragging'));
    });

    sheetTabsList.addEventListener('click', (e) => {
      const closeBtn = e.target.closest('[data-delete-sheet]');
      if (closeBtn) {
        e.stopPropagation();
        deleteSheet(closeBtn.getAttribute('data-delete-sheet'));
        return;
      }
      const btn = e.target.closest('.sheet-tab[data-sheet]');
      if (!btn) return;
      switchSheet(btn.getAttribute('data-sheet'));
    });

    sheetTabsList.addEventListener('dblclick', (e) => {
      if (e.target.closest('[data-delete-sheet]')) return;
      const btn = e.target.closest('.sheet-tab[data-sheet]');
      if (!btn) return;
      renameSheet(btn.getAttribute('data-sheet'));
    });

    sheetTabsList.addEventListener('contextmenu', (e) => {
      const btn = e.target.closest('.sheet-tab[data-sheet]');
      if (!btn) return;
      e.preventDefault();
      closeInlineEditor(true);
      const targetSheet = btn.getAttribute('data-sheet');
      const order = visibleSheets();
      const canDelete = order.length > 1;
      const hiddenList = (workbook.sheetOrder || []).filter(n => workbook.sheets[n] && workbook.sheets[n].hidden === 'hidden');

      ctxMenu.innerHTML = `
        <button class="excel-context-item" data-ctx="sheet-add">
          <span>➕ Insert New Sheet</span>
        </button>
        <button class="excel-context-item" data-ctx="sheet-duplicate" data-target-sheet="${escapeHtml(targetSheet)}">
          <span>📑 Duplicate "${escapeHtml(targetSheet)}"</span>
        </button>
        <button class="excel-context-item" data-ctx="sheet-rename" data-target-sheet="${escapeHtml(targetSheet)}">
          <span>✏️ Rename</span>
        </button>
        ${canDelete ? `
        <button class="excel-context-item" data-ctx="sheet-hide" data-target-sheet="${escapeHtml(targetSheet)}">
          <span>🙈 Ocultar "${escapeHtml(targetSheet)}"</span>
        </button>` : ''}
        ${hiddenList.map(n => `
        <button class="excel-context-item" data-ctx="sheet-unhide" data-target-sheet="${escapeHtml(n)}">
          <span>👁 Mostrar "${escapeHtml(n)}"</span>
        </button>`).join('')}
        ${canDelete ? `
        <div class="excel-context-divider"></div>
        <button class="excel-context-item danger" data-ctx="sheet-delete" data-target-sheet="${escapeHtml(targetSheet)}">
          <span>🗑️ Delete "${escapeHtml(targetSheet)}"</span>
        </button>` : ''}
      `;

      ctxMenu.style.display = 'flex';
      const menuW = 230;
      const menuH = 200 + hiddenList.length * 34;
      ctxMenu.style.left = `${Math.min(window.innerWidth - menuW - 10, e.clientX)}px`;
      ctxMenu.style.top = `${Math.max(10, e.clientY - menuH)}px`;
    });

    document.getElementById('btnAddSheet').addEventListener('click', addNewSheet);

    document.getElementById('btnPrevSheet').addEventListener('click', () => {
      const order = visibleSheets();
      const idx = order.indexOf(activeSheet);
      if (idx > 0) switchSheet(order[idx - 1]);
    });
    document.getElementById('btnNextSheet').addEventListener('click', () => {
      const order = visibleSheets();
      const idx = order.indexOf(activeSheet);
      if (idx < order.length - 1) switchSheet(order[idx + 1]);
    });

    // Utility buttons
    document.getElementById('btnUndo').addEventListener('click', undo);
    document.getElementById('btnRedo').addEventListener('click', redo);
    document.getElementById('btnExportCsv').addEventListener('click', exportActiveSheetCsv);
    document.getElementById('btnActualizarData')?.addEventListener('click', actualizarDataDesdePosco);
    document.getElementById('btnActualizarInventario')?.addEventListener('click', actualizarInventarioDesdePosco);

    document.getElementById('btnResetWorkbook')?.addEventListener('click', async () => {
      if (!confirm('Clear and reset workbook to blank?')) return;
      localStorage.removeItem(BACKUP_KEY);
      const res = await fetch(`/api/reset${API_Q}`, { method: 'POST' });
      workbook = await res.json();
      activeSheet = workbook.sheetOrder[0] || 'Sheet1';
      engine = new window.FormulaEngine(workbook);
      engine.recalculateAll();
      renderSheetTabs();
      renderSheet();
      selectCell('A1', false);
      showToast('Workbook reset');
    });

    // Zoom controls
    document.getElementById('btnZoomIn').addEventListener('click', () => {
      zoomLevel = Math.min(140, zoomLevel + 10);
      sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
      document.getElementById('zoomLabel').textContent = `${zoomLevel}%`;
    });
    document.getElementById('btnZoomOut').addEventListener('click', () => {
      zoomLevel = Math.max(60, zoomLevel - 10);
      sheetCanvasWrapper.style.zoom = `${zoomLevel}%`;
      document.getElementById('zoomLabel').textContent = `${zoomLevel}%`;
    });

    // Format Toolbar: Clipboard group
    btnPaste.addEventListener('click', pasteFromClipboard);
    btnCut.addEventListener('click', () => copySelection(true));
    btnCopy.addEventListener('click', () => copySelection(false));

    // Format Toolbar: font family / size
    fontFamilySelect.addEventListener('change', () => setFontFamilySelection(fontFamilySelect.value));
    fontSizeSelect.addEventListener('change', () => setFontSizeSelection(Number(fontSizeSelect.value)));
    btnGrowFont.addEventListener('click', () => growShrinkFontSelection(2));
    btnShrinkFont.addEventListener('click', () => growShrinkFontSelection(-2));

    // Format Toolbar: bold / italic / underline / strikethrough
    btnBold.addEventListener('click', toggleBoldSelection);
    btnItalic.addEventListener('click', toggleItalicSelection);
    btnUnderline.addEventListener('click', toggleUnderlineSelection);
    btnStrike.addEventListener('click', toggleStrikeSelection);

    // Format Toolbar: font & fill color
    // Font & fill color: split buttons with Excel-style palettes
    const THEME_COLORS = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
    const THEME_TINTS = ['#F2F2F2', '#7F7F7F', '#D0CECE', '#D6DCE4', '#D9E1F2', '#FCE4D6', '#EDEDED', '#FFF2CC', '#DDEBF7', '#E2EFDA'];
    const STANDARD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

    function setColorOnSelection(target, val, bounds, withUndo = true) {
      applyStyleToSelection(s => { if (val) s[target] = val; else delete s[target]; }, bounds, withUndo);
      if (val) {
        if (target === 'bg') { lastFillColor = val; fillColorSwatch.style.background = val; }
        else { lastFontColor = val; fontColorSwatch.style.background = val; }
      }
    }

    function buildPalette(menu) {
      const target = menu.getAttribute('data-target');
      const swatches = colors => `<div class="fmt-palette-row">${colors.map(c =>
        `<button type="button" class="fmt-swatch" data-color="${c}" title="${c}" style="background:${c}"></button>`).join('')}</div>`;
      menu.innerHTML = `
        <button type="button" class="fmt-palette-action" data-color="">${target === 'bg' ? '⊘ Sin relleno' : '■ Automático'}</button>
        <div class="fmt-palette-title">Colores del tema</div>
        ${swatches(THEME_COLORS)}${swatches(THEME_TINTS)}
        <div class="fmt-palette-title">Colores estándar</div>
        ${swatches(STANDARD_COLORS)}
        <label class="fmt-palette-action">🎨 Más colores…<input type="color" value="${target === 'bg' ? lastFillColor : lastFontColor}"></label>`;

      let pickerBounds = null;
      let pickerUndoPushed = false;
      const picker = menu.querySelector('input[type="color"]');
      picker.addEventListener('click', () => {
        pickerBounds = getSelectionBounds();
        pickerUndoPushed = false;
      });
      picker.addEventListener('input', () => {
        if (!pickerBounds) pickerBounds = getSelectionBounds();
        setColorOnSelection(target, picker.value, pickerBounds, !pickerUndoPushed);
        pickerUndoPushed = true;
      });

      menu.addEventListener('click', (e) => {
        const sw = e.target.closest('[data-color]');
        if (!sw) return;
        setColorOnSelection(target, sw.getAttribute('data-color'));
        closeAllFormatDropdowns();
      });
    }
    buildPalette(fillColorMenu);
    buildPalette(fontColorMenu);

    btnFillColor.addEventListener('click', () => setColorOnSelection('bg', lastFillColor));
    btnFontColor.addEventListener('click', () => setColorOnSelection('color', lastFontColor));
    [[btnFillColorMenu, fillColorMenu], [btnFontColorMenu, fontColorMenu]].forEach(([btn, menu]) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const wasOpen = menu.classList.contains('open');
        closeAllFormatDropdowns();
        if (!wasOpen) menu.classList.add('open');
      });
    });

    // Format Toolbar: horizontal & vertical alignment
    btnAlignLeft.addEventListener('click', () => setAlignSelection('left'));
    btnAlignCenter.addEventListener('click', () => setAlignSelection('center'));
    btnAlignRight.addEventListener('click', () => setAlignSelection('right'));
    btnValignTop.addEventListener('click', () => setValignSelection('top'));
    btnValignMiddle.addEventListener('click', () => setValignSelection('middle'));
    btnValignBottom.addEventListener('click', () => setValignSelection('bottom'));
    btnWrapText.addEventListener('click', toggleWrapSelection);
    btnMergeCenter.addEventListener('click', toggleMergeCenterSelection);

    // Format Toolbar: dropdowns (borders / number format)
    function closeAllFormatDropdowns() {
      document.querySelectorAll('.fmt-dropdown-menu.open').forEach(el => el.classList.remove('open'));
    }
    btnBorders.addEventListener('click', (e) => {
      e.stopPropagation();
      const wasOpen = borderMenu.classList.contains('open');
      closeAllFormatDropdowns();
      if (!wasOpen) borderMenu.classList.add('open');
    });
    btnNumFmt.addEventListener('click', (e) => {
      e.stopPropagation();
      const wasOpen = numFmtMenu.classList.contains('open');
      closeAllFormatDropdowns();
      if (!wasOpen) numFmtMenu.classList.add('open');
    });
    borderMenu.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-border]');
      if (!btn) return;
      applyBorderPreset(btn.getAttribute('data-border'));
      closeAllFormatDropdowns();
    });
    numFmtMenu.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-numfmt]');
      if (!btn) return;
      const fmt = btn.getAttribute('data-numfmt');
      applyStyleToSelection(s => { delete s.nf; if (fmt === 'general') delete s.numFmt; else s.numFmt = fmt; });
      closeAllFormatDropdowns();
    });
    window.addEventListener('mousedown', (e) => {
      if (!e.target.closest('.fmt-dropdown')) closeAllFormatDropdowns();
    });

    // Format Toolbar: quick number format & decimals
    btnCurrencyFmt.addEventListener('click', () => setQuickNumFmtSelection('currency'));
    btnPercentFmt.addEventListener('click', () => setQuickNumFmtSelection('percent'));
    btnCommaFmt.addEventListener('click', () => setQuickNumFmtSelection('number'));
    btnDecimalInc.addEventListener('click', () => adjustDecimalsSelection(1));
    btnDecimalDec.addEventListener('click', () => adjustDecimalsSelection(-1));

    // Format Toolbar: format painter & clear formatting
    btnFormatPainter.addEventListener('click', () => {
      if (formatPainterActive) stopFormatPainter();
      else startFormatPainter();
    });
    btnClearFormat.addEventListener('click', clearFormatSelection);

    document.getElementById('btnZeroRows')?.addEventListener('click', () => {
      const sheet = workbook.sheets[activeSheet];
      if (!sheet) return;
      sheet.showZeroRows = !sheet.showZeroRows;
      renderSheet();
      saveWorkbookState();
      showToast(sheet.showZeroRows ? 'Se muestran las flores en 0' : 'Flores en 0 ocultas');
    });
  }

  window.MauriApp = {
    getWorkbook: () => workbook,
    getWorkbookKey: () => WB_KEY,
    toast: (msg, isWarn) => showToast(msg, isWarn),
    getSelectionBounds: () => getSelectionBounds(),
    selectCoord: coord => selectCell(coord, false),
    getActiveSheet: () => activeSheet,
    getEngine: () => engine,
    pushUndo: names => pushUndoState(names),
    markDirty: name => markSheetDirty(name),
    save: () => saveWorkbookState(),
    recalcAndRender: () => {
      engine.recalculateAll();
      renderSheetTabs();
      renderSheet();
      highlightSelectedCellInDOM();
      updateStatusMetrics();
    },
    switchSheet: name => switchSheet(name),
    cellDisplayText: (sheetName, coord) => {
      const sheet = workbook.sheets[sheetName];
      const p = parseCoord(coord);
      return sheet && p ? cellText(sheet, p.row, p.col) : '';
    },
    previewFilterHidden: (name) => {
      const sheet = workbook.sheets[name];
      const keep = sheet.filterHiddenRows;
      recomputeFilterHidden(sheet);
      const computed = sheet.filterHiddenRows;
      sheet.filterHiddenRows = keep;
      return computed;
    },
    showToast: (msg, isWarn) => showToast(msg, isWarn)
  };

  init();
})();
