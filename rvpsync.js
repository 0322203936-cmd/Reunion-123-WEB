// ============================================================================
// REQUERIMIENTO VS PROYECCION: keeps the "Append" sheet live (the server rebuilds it from
// PegarData and Plan de cosecha and announces it on /api/events)
// ============================================================================

(function () {
  if (new URLSearchParams(location.search).get('wb') !== 'rvp') return;
  const App = () => window.MauriApp;

  async function refreshAppend() {
    const app = App();
    if (!app || !app.getWorkbook() || !app.getEngine()) { setTimeout(refreshAppend, 500); return; }
    const res = await fetch('/api/sheet?wb=rvp&name=Append', { cache: 'no-store' });
    if (!res.ok) return;
    const sheet = await res.json();
    const wb = app.getWorkbook();
    wb.sheets.Append = sheet;
    if (!wb.sheetOrder.includes('Append')) wb.sheetOrder.push('Append');
    app.recalcAndRender();
    if (window.MauriPivot && window.MauriPivot.refreshAll) window.MauriPivot.refreshAll('Append');
  }

  // the Append may have changed while this page was closed: bring the pivots up to date once loaded
  (function refreshOnLoad() {
    const app = App();
    if (!app || !app.getWorkbook() || !app.getEngine() || !window.MauriPivot || !window.MauriPivot.refreshAll) { setTimeout(refreshOnLoad, 500); return; }
    window.MauriPivot.refreshAll('Append');
  })();

  function listen() {
    const es = new EventSource('/api/events?wb=rvp');
    es.addEventListener('append', () => { refreshAppend().catch(err => console.error('Append:', err)); });
    es.onerror = () => { es.close(); setTimeout(listen, 3000); };
  }
  listen();
})();
