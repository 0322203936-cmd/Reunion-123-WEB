// Adds a sheet (converted from Excel with tools/xlsx_to_workbook.py) to Plan de cosecha, with the
// system running: it goes through the server, so it is saved in Neon and every open page sees it.
// A "P Cosecha NNNN" tab is placed in week order (before the first tab with a higher number).
//
// Usage:  node tools/importar_hoja.js <hoja.json> ["Nombre de la pestaña"]
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const WB = 'cosecha';

function loadLocalEnv() {
  const envPath = path.join(ROOT, '.env');
  if (!fs.existsSync(envPath)) return;
  for (const rawLine of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const separator = line.indexOf('=');
    if (separator < 1) continue;
    const key = line.slice(0, separator).trim();
    let value = line.slice(separator + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (process.env[key] === undefined) process.env[key] = value;
  }
}

const weekNumber = name => { const m = /p\s*cosecha\s*(\d+)/i.exec(name); return m ? Number(m[1]) : null; };

async function main() {
  loadLocalEnv();
  const [file, wanted] = process.argv.slice(2);
  if (!file) throw new Error('Falta el archivo: node tools/importar_hoja.js <hoja.json> ["Nombre"]');
  const source = JSON.parse(fs.readFileSync(path.resolve(file), 'utf8'));
  const sourceName = (source.sheetOrder || Object.keys(source.sheets))[0];
  const sheet = source.sheets[sourceName];
  const name = String(wanted || sourceName).trim();

  const base = `http://127.0.0.1:${Number(process.env.PORT || 3070)}`;
  let res;
  try {
    res = await fetch(`${base}/api/workbook?wb=${WB}`, { cache: 'no-store' });
  } catch (error) {
    throw new Error('El sistema no está prendido. Préndelo con npm.cmd start y vuelve a correr esto.');
  }
  if (!res.ok) throw new Error(`No se pudo leer Plan de cosecha (${res.status})`);
  const book = await res.json();
  const order = (book.sheetOrder || Object.keys(book.sheets)).slice();
  const clash = order.find(n => n.trim().toUpperCase() === name.toUpperCase());
  if (clash) throw new Error(`Ya existe una pestaña "${clash}". Cámbiale el nombre en el sistema o usa otro nombre.`);

  const week = weekNumber(name);
  let at = order.length;
  if (week !== null) {
    const next = order.findIndex(n => weekNumber(n) !== null && weekNumber(n) > week);
    if (next >= 0) at = next;
  }
  order.splice(at, 0, name);

  res = await fetch(`${base}/api/workbook?wb=${WB}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ meta: { sheetOrder: order }, sheets: { [name]: sheet } })
  });
  if (!res.ok) throw new Error(`El sistema no aceptó la hoja (${res.status}): ${await res.text()}`);
  console.log(` Listo: "${name}" agregada a Plan de cosecha (${Object.keys(sheet.cells).length} celdas).`);
  console.log(` Orden de pestañas: ${order.map(n => n.trim()).join(' | ')}`);
}

main().catch(err => { console.error(' Error:', err.message); process.exit(1); });
