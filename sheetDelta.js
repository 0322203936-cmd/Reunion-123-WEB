// ============================================================================
// SHEET DELTAS (shared by the page, server.js and neonStore.js)
//   A delta is what changed in a sheet: { cells: { A1: cell | null }, props: { key: value }, removed: [key] }
//   (null = cell deleted; props = any other key of the sheet such as colWidths, merges...)
//   The page sends deltas instead of whole sheets, and Neon stores them, so a typed number
//   travels and is written as a few bytes instead of the whole sheet.
// ============================================================================

(function (root) {
  function applyDelta(sheet, delta) {
    if (delta.cells) {
      sheet.cells = sheet.cells || {};
      for (const coord of Object.keys(delta.cells)) {
        const cell = delta.cells[coord];
        if (cell === null) delete sheet.cells[coord];
        else sheet.cells[coord] = cell;
      }
    }
    if (delta.props) Object.assign(sheet, delta.props);
    for (const key of delta.removed || []) delete sheet[key];
    return sheet;
  }

  function propsOf(sheet) {
    const knownProps = new Map();
    for (const key of Object.keys(sheet)) if (key !== 'cells') knownProps.set(key, JSON.stringify(sheet[key]));
    return knownProps;
  }

  // remembers a sheet as last sent / stored: { known: Map coord -> JSON, knownProps: Map key -> JSON }
  function snapshotOf(sheet) {
    const known = new Map();
    const cells = sheet.cells || {};
    for (const coord of Object.keys(cells)) known.set(coord, JSON.stringify(cells[coord]));
    const knownProps = new Map();
    for (const key of Object.keys(sheet)) if (key !== 'cells') knownProps.set(key, JSON.stringify(sheet[key]));
    return { known, knownProps };
  }

  // what changed since `snap` (which is updated to the new state). `coords` (Set) limits the cells
  // compared to the ones known to be touched; without it every cell is compared.
  // Returns null when nothing changed, or { delta, size } (size: number of changed entries).
  function diffSheet(snap, sheet, coords) {
    const { known, knownProps } = snap;
    const cells = {};
    const now = sheet.cells || {};
    let n = 0;
    const check = coord => {
      const cell = now[coord];
      if (cell === undefined) {
        if (known.has(coord) || snap.partial) { known.delete(coord); cells[coord] = null; n++; }
        return;
      }
      const s = JSON.stringify(cell);
      if (known.get(coord) === s) return;
      known.set(coord, s);
      cells[coord] = cell;
      n++;
    };
    if (coords) coords.forEach(check);
    else {
      for (const coord of Object.keys(now)) check(coord);
      for (const coord of [...known.keys()]) if (!(coord in now)) check(coord);
    }
    const props = {};
    const removed = [];
    const keys = new Set([...Object.keys(sheet), ...knownProps.keys()]);
    for (const key of keys) {
      if (key === 'cells') continue;
      if (!(key in sheet)) { knownProps.delete(key); removed.push(key); n++; continue; }
      const s = JSON.stringify(sheet[key]);
      if (knownProps.get(key) === s) continue;
      knownProps.set(key, s);
      props[key] = sheet[key];
      n++;
    }
    if (!n) return null;
    const delta = {};
    if (Object.keys(cells).length) delta.cells = cells;
    if (Object.keys(props).length) delta.props = props;
    if (removed.length) delta.removed = removed;
    return { delta, size: n };
  }

  // a newer delta of the same sheet on top of an older one
  function mergeDelta(into, delta) {
    if (delta.cells) into.cells = Object.assign(into.cells || {}, delta.cells);
    if (delta.props) {
      into.props = Object.assign(into.props || {}, delta.props);
      if (into.removed) into.removed = into.removed.filter(k => !(k in delta.props));
    }
    if (delta.removed) {
      into.removed = [...new Set([...(into.removed || []), ...delta.removed])];
      if (into.props) for (const k of delta.removed) delete into.props[k];
    }
    return into;
  }

  const api = { applyDelta, propsOf, snapshotOf, diffSheet, mergeDelta };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SheetDelta = api;
})(typeof window !== 'undefined' ? window : globalThis);
