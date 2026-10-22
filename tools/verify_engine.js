// Compare the app formula engine against the values Excel cached in the file.
// Usage: node tools/verify_engine.js <workbook.json> <expected.json>
const fs = require('fs');
const path = require('path');
const { FormulaEngine } = require(path.join(__dirname, '..', 'formulaEngine.js'));

const [, , wbPath, expPath] = process.argv;
const workbook = JSON.parse(fs.readFileSync(wbPath, 'utf8'));
const expected = JSON.parse(fs.readFileSync(expPath, 'utf8'));

const t0 = Date.now();
const engine = new FormulaEngine(workbook);
engine.recalculateAll();
const t1 = Date.now();

let ok = 0;
const bad = [];
for (const [key, exp] of Object.entries(expected)) {
  const [sheet, coord] = key.split('!');
  const cell = workbook.sheets[sheet].cells[coord];
  const got = cell ? cell.v : undefined;
  const same = (() => {
    if (exp === null) return got === '' || got === 0 || got === null;
    if (typeof exp === 'number' && typeof got === 'number') return Math.abs(exp - got) <= 1e-9 * Math.max(1, Math.abs(exp));
    if (typeof exp === 'boolean') return got === exp;
    return String(exp) === String(got);
  })();
  if (same) ok++; else bad.push({ key, exp, got, f: cell && cell.f });
}
console.log(`recalculateAll: ${t1 - t0} ms`);
console.log(`iguales: ${ok} / ${Object.keys(expected).length}   distintas: ${bad.length}`);
const bySheet = {};
for (const b of bad) bySheet[b.key.split('!')[0]] = (bySheet[b.key.split('!')[0]] || 0) + 1;
console.log('distintas por hoja:', bySheet);
for (const b of bad.slice(0, 25)) console.log(JSON.stringify(b));
