// ============================================================================
// REQUERIMIENTO VS PROYECCION: keeps the "Append" sheet live (the server rebuilds it from
// PegarData and Plan de cosecha and announces it on /api/events)
// ============================================================================

(function () {
  const App = () => window.MauriApp;

  // Append is big and can take a while to arrive: by then the user may be in another book, so it goes
  // into the Requerimiento Vs Proyeccion book (never the open one) and is only drawn if that is still open
  let inFlight = null;
  let again = false;
  function refreshAppend() {
    if (inFlight) { again = true; return inFlight; }
    inFlight = (async () => {
      const app = App();
      if (!app || app.getWorkbookKey?.() !== 'rvp') return;
      if (!app.getWorkbook() || !app.getEngine()) { setTimeout(refreshAppend, 500); return; }
      const res = await fetch('/api/sheet?wb=rvp&name=Append', { cache: 'no-store' });
      if (!res.ok) return;
      const sheet = await res.json();
      const wb = app.getCachedWorkbook ? app.getCachedWorkbook('rvp') : null;
      if (!wb || !wb.sheets) return;
      wb.sheets.Append = sheet;
      if (!wb.sheetOrder.includes('Append')) wb.sheetOrder.push('Append');
      if (app.getWorkbookKey() !== 'rvp' || app.getWorkbook() !== wb) return; // drawn when the book is opened again
      app.recalcAndRender();
      if (window.MauriPivot && window.MauriPivot.refreshAll) window.MauriPivot.refreshAll('Append');
    })().catch(err => console.error('Append:', err)).finally(() => {
      inFlight = null;
      if (again) { again = false; refreshAppend(); }
    });
    return inFlight;
  }

  // the Append may have changed while this page was closed: bring the pivots up to date once loaded
  (function refreshOnLoad() {
    const app = App();
    if (!app || !app.getWorkbook() || !app.getEngine() || !window.MauriPivot || !window.MauriPivot.refreshAll) { setTimeout(refreshOnLoad, 500); return; }
    if (app.getWorkbookKey?.() === 'rvp') window.MauriPivot.refreshAll('Append');
  })();

  window.addEventListener('mauri:workbook-switch', event => {
    if (event.detail?.key !== 'rvp') return;
    refreshAppend();
  });

  window.MauriEvents.on('append', () => { refreshAppend(); });
})();
