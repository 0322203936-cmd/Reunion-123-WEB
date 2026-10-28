const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

const POSCO_URL = 'http://3.132.9.174/Posco/';
const SHEET_NAME = 'PegarData';
const TABLE_NAME = 'Table2';
const SOURCE_COLUMNS = 19;
const TARGET_COLUMNS = 18;
const CALCULATED_HEADERS = ['Flor Tallos', 'TxR2', 'Flor Color2', 'Flor Color 3', 'Flor Color 4'];
const MATRICARIA_5 = new Set([
  'STATER BROS', 'GALLERIA #3758', 'NOR-CAL PRODUCE INC.', 'MAYESH',
  'GALLERIA #710269', 'SAFEWAY NORCAL', "GELSON'S MARKETS /  11703", "CHARLIE'S PRODUCE SEATTLE"
]);

function requiredEnv(name) {
  const value = String(process.env[name] || '').trim();
  if (!value) throw new Error(`Falta configurar ${name} en el entorno local.`);
  return value;
}

function envFlag(name, fallback) {
  const value = process.env[name];
  if (value === undefined) return fallback;
  return ['1', 'true', 'yes', 'si', 'sí', 'on'].includes(String(value).trim().toLowerCase());
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function text(value) {
  return String(value ?? '').trim();
}

function numberValue(value) {
  const parsed = Number(String(value ?? '').trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function excelSerial(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  }
  const raw = text(value);
  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(raw);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(raw);
  const year = iso ? Number(iso[1]) : us ? Number(us[3]) : null;
  const month = iso ? Number(iso[2]) : us ? Number(us[1]) : null;
  const day = iso ? Number(iso[3]) : us ? Number(us[2]) : null;
  if (!year || !month || !day) return value;
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000;
}

function localToday() {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Tijuana', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.filter(part => part.type !== 'literal').map(part => [part.type, Number(part.value)]));
  return new Date(values.year, values.month - 1, values.day);
}

function dateRange() {
  const today = localToday();
  const daysSinceFriday = (today.getDay() + 2) % 7;
  const daysBack = daysSinceFriday === 0 ? 7 : daysSinceFriday;
  const start = new Date(today);
  start.setDate(start.getDate() - daysBack);
  const end = new Date(start);
  end.setDate(end.getDate() + 21);
  return { start, end };
}

function poscoDate(value) {
  const month = String(value.getMonth() + 1).padStart(2, '0');
  const day = String(value.getDate()).padStart(2, '0');
  return `${month}/${day}/${value.getFullYear()}`;
}

async function launchBrowser() {
  const headless = envFlag('POSCO_HEADLESS', true);
  const executablePath = String(process.env.POSCO_BROWSER_EXECUTABLE_PATH || '').trim();
  // Render does not provide Microsoft Edge. Leave the channel empty there so
  // Playwright uses its bundled Chromium; local .env can still opt into Edge.
  const channel = String(process.env.POSCO_BROWSER_CHANNEL || '').trim();
  const options = { headless };
  if (executablePath) options.executablePath = executablePath;
  else if (channel) options.channel = channel;
  try {
    return await chromium.launch(options);
  } catch (error) {
    if (!executablePath && channel) {
      try {
        return await chromium.launch({ headless });
      } catch (fallbackError) {
        throw new Error(`No se pudo abrir el navegador de POSCO (${error.message}; respaldo: ${fallbackError.message}).`);
      }
    }
    throw new Error(`No se pudo abrir el navegador de POSCO: ${error.message}`);
  }
}

async function visibleText(page, value) {
  const candidates = page.getByText(value, { exact: true });
  for (let index = 0; index < await candidates.count(); index += 1) {
    const element = candidates.nth(index);
    if (await element.isVisible()) {
      await element.click({ timeout: 15000 });
      return;
    }
  }
  throw new Error(`No se encontró el control visible ${value}.`);
}

async function openOrders(page) {
  const candidates = [
    page.getByRole('link', { name: 'Ordenes', exact: true }),
    page.getByRole('button', { name: 'Ordenes', exact: true }),
    page.locator('[ngbdropdowntoggle]:has-text("Ordenes")'),
    page.getByText('Ordenes', { exact: true })
  ];
  for (const candidate of candidates) {
    for (let index = 0; index < await candidate.count(); index += 1) {
      const element = candidate.nth(index);
      if (await element.isVisible()) {
        await element.click({ timeout: 15000 });
        return;
      }
    }
  }
  throw new Error('No se encontró el menú Órdenes de POSCO.');
}

async function selectOrders(page) {
  const route = page.locator('a[href="#/list-orden-detalle"]');
  for (let index = 0; index < await route.count(); index += 1) {
    const element = route.nth(index);
    if (await element.isVisible()) {
      await element.click({ timeout: 15000 });
      return;
    }
  }
  throw new Error('No se encontró la opción Órdenes de POSCO.');
}

async function dateInput(page, label, fallbackIndex) {
  const group = page.locator(`.input-group:has-text("${label}") input`);
  if (await group.count()) return group.first();
  const labelNode = page.getByText(label, { exact: true });
  if (await labelNode.count()) {
    const sibling = labelNode.first().locator('xpath=..').locator('input');
    if (await sibling.count()) return sibling.first();
  }
  const dateInputs = page.locator('input[type="date"]');
  if (await dateInputs.count() > fallbackIndex) return dateInputs.nth(fallbackIndex);
  throw new Error(`No se encontró el campo ${label}.`);
}

async function fillDateInput(locator, value) {
  const type = String(await locator.getAttribute('type') || 'text').toLowerCase();
  await locator.fill(type === 'date' ? `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}` : poscoDate(value));
  await locator.press('Tab');
}

async function selectActive(page) {
  const selects = page.locator('select');
  for (let index = 0; index < await selects.count(); index += 1) {
    const select = selects.nth(index);
    const options = await select.locator('option').allTextContents();
    if (options.map(item => item.trim()).includes('ACTIVO')) {
      await select.selectOption({ label: 'ACTIVO' });
      return;
    }
  }
  throw new Error('No se encontró el filtro ACTIVO de POSCO.');
}

async function clickSearch(page) {
  const candidates = [
    page.locator('button:has-text("Buscar"), a:has-text("Buscar")'),
    page.locator('[class*="btn"]:has-text("Buscar")'),
    page.getByText('Buscar', { exact: true })
  ];
  for (const candidate of candidates) {
    for (let index = 0; index < await candidate.count(); index += 1) {
      const element = candidate.nth(index);
      if (await element.isVisible()) {
        await element.click({ timeout: 15000 });
        return;
      }
    }
  }
  throw new Error('No se encontró el botón Buscar de POSCO.');
}

async function downloadPoscoReport() {
  const user = requiredEnv('POSCO_USER');
  const password = requiredEnv('POSCO_PASSWORD');
  const browser = await launchBrowser();
  const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1920, height: 1080 } });
  const page = await context.newPage();
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'reunion123-posco-'));
  try {
    await page.goto(POSCO_URL, { waitUntil: 'networkidle', timeout: 60000 });
    await page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first().fill(user);
    await page.locator('input[placeholder*="Password" i], input[type="password"]').first().fill(password);
    await page.getByRole('button', { name: 'Iniciar Sesión' }).click({ timeout: 15000 });
    await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(2000);

    await openOrders(page);
    await page.waitForTimeout(800);
    await selectOrders(page);
    await page.waitForURL('**/#/list-orden-detalle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(3000);

    const { start, end } = dateRange();
    await fillDateInput(await dateInput(page, 'Load Date Menor', 0), start);
    await fillDateInput(await dateInput(page, 'Load Date Mayor', 1), end);
    await clickSearch(page);
    await page.waitForTimeout(Number(process.env.POSCO_LOAD_WAIT_MS || 120000));
    await selectActive(page);
    await page.waitForTimeout(6000);

    await visibleText(page, 'Exportar');
    await page.waitForTimeout(500);
    const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
    await visibleText(page, 'Exportar Color filtro');
    const download = await downloadPromise;
    const destination = path.join(temporaryDirectory, path.basename(download.suggestedFilename() || 'posco-reporte.xlsx'));
    await download.saveAs(destination);
    console.log(`Reporte POSCO descargado: ${destination}`);
    return { path: destination, temporaryDirectory, start, end };
  } catch (error) {
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    throw error;
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

function correctedTxr(category, value) {
  return text(category).toUpperCase() === 'BOUQUET' ? 8 : 10;
}

function formatReport(reportPath) {
  const sourceWorkbook = XLSX.readFile(reportPath, { cellDates: false, raw: true });
  const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
  const sourceRows = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, raw: true, defval: '' });
  if (!sourceRows.length || sourceRows[0].length < SOURCE_COLUMNS) {
    throw new Error('El reporte POSCO no contiene las 19 columnas esperadas.');
  }
  if (text(sourceRows[0][13]) !== 'txr_orden') {
    throw new Error(`La columna N del reporte no es txr_orden; se encontró ${text(sourceRows[0][13]) || 'vacía'}.`);
  }
  if (text(sourceRows[0][18]) !== 'Horas') {
    throw new Error(`La columna S del reporte no es Horas; se encontró ${text(sourceRows[0][18]) || 'vacía'}.`);
  }

  const rows = sourceRows.map((sourceRow, rowIndex) => {
    const row = Array.from({ length: SOURCE_COLUMNS }, (_, index) => sourceRow[index] ?? '');
    // POSCO exporta A:M, N:R y S=Horas. La macro mueve N al final de N:R
    // y después elimina S, por lo que el resultado web queda en A:R.
    const reordered = row.slice(0, 13).concat(row.slice(14, 18), [row[13]]);
    if (rowIndex === 0) return reordered;

    // Reunion compara las fechas con seriales Excel, no con textos ISO.
    reordered[7] = excelSerial(reordered[7]);
    reordered[14] = excelSerial(reordered[14]);

    const flower = text(reordered[4]).toUpperCase();
    const category = text(reordered[3]).toUpperCase();
    const currentTxr = numberValue(reordered[17]);
    if (flower.startsWith('VERONICA')) reordered[17] = correctedTxr(category, currentTxr);
    else if (flower.startsWith('SNAPDRAGON') && ![5, 10].includes(currentTxr)) reordered[17] = 5;
    else if (flower === 'MATRICARIA' && ![4, 5, 10].includes(currentTxr)) {
      const customer = text(reordered[1]).toUpperCase();
      reordered[17] = customer === 'GALLERIA #710332' ? 4 : MATRICARIA_5.has(customer) ? 5 : 6;
    }
    return reordered;
  });
  return rows;
}

function columnLetter(number) {
  let result = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function cellValue(sheet, column, row) {
  return sheet.cells?.[`${columnLetter(column)}${row}`]?.v ?? '';
}

function ensureCell(sheet, column, row, template) {
  const coord = `${columnLetter(column)}${row}`;
  if (!sheet.cells) sheet.cells = {};
  if (!sheet.cells[coord]) sheet.cells[coord] = template ? clone(template) : {};
  return sheet.cells[coord];
}

function setCell(sheet, column, row, value, formula, template) {
  const cell = ensureCell(sheet, column, row, template);
  if (value === undefined || value === '') delete cell.v;
  else cell.v = value;
  if (formula === undefined) delete cell.f;
  else cell.f = formula;
}

function clearCell(sheet, column, row) {
  const cell = sheet.cells?.[`${columnLetter(column)}${row}`];
  if (!cell) return;
  delete cell.v;
  delete cell.f;
}

function applyToWorkbook(workbook, rows) {
  const sheet = workbook.sheets?.[SHEET_NAME];
  if (!sheet) throw new Error('No existe la hoja PegarData en el sistema.');
  const expectedHeaders = Array.from({ length: TARGET_COLUMNS }, (_, index) => cellValue(sheet, index + 1, 1));
  if (JSON.stringify(expectedHeaders) !== JSON.stringify(rows[0])) {
    throw new Error('Los encabezados del reporte POSCO no coinciden con la hoja PegarData.');
  }
  const calculatedHeaders = CALCULATED_HEADERS.map((_, index) => cellValue(sheet, index + 19, 1));
  if (JSON.stringify(calculatedHeaders) !== JSON.stringify(CALCULATED_HEADERS)) {
    throw new Error('Los encabezados calculados S:W no coinciden con PegarData.');
  }

  const oldLastRow = Math.max(sheet.maxRow || 1, 1);
  const newLastRow = Math.max(rows.length, 1);
  const templates = {};
  for (let column = 1; column <= 23; column += 1) {
    templates[column] = sheet.cells?.[`${columnLetter(column)}2`]?.s ? { s: clone(sheet.cells[`${columnLetter(column)}2`].s) } : {};
  }

  for (let row = 1; row <= newLastRow; row += 1) {
    const values = rows[row - 1];
    for (let column = 1; column <= TARGET_COLUMNS; column += 1) {
      setCell(sheet, column, row, values[column - 1], undefined, templates[column]);
    }
    if (row === 1) continue;
    setCell(sheet, 19, row, '', '=+Table2[[#This Row],[FLOR]]&" "&"X"&" " & Table2[[#This Row],[txr_orden]]', templates[19]);
    setCell(sheet, 20, row, '', '=Table2[[#This Row],[txr_orden]]', templates[20]);
    setCell(sheet, 21, row, '', '=+Table2[[#This Row],[FLOR]] & " " &Table2[[#This Row],[Flor Color]]', templates[21]);
    setCell(sheet, 22, row, '', '=+Table2[[#This Row],[FLOR]] & " " & Table2[[#This Row],[Flor Color]]&"X"&" " & Table2[[#This Row],[TxR2]]', templates[22]);
    setCell(sheet, 23, row, '', `=E${row} & " " & " X " & T${row}`, templates[23]);
  }

  for (let row = newLastRow + 1; row <= oldLastRow; row += 1) {
    for (let column = 1; column <= 23; column += 1) clearCell(sheet, column, row);
  }
  sheet.maxRow = Math.max(oldLastRow, newLastRow + 20, 100);
  sheet.maxCol = Math.max(sheet.maxCol || 23, 23);
  if (workbook.tables?.[TABLE_NAME]) workbook.tables[TABLE_NAME].ref = `A1:W${newLastRow}`;
  return { rows: Math.max(newLastRow - 1, 0), lastRow: newLastRow };
}

async function runActualizarData(workbook) {
  const report = await downloadPoscoReport();
  try {
    const rows = formatReport(report.path);
    const updatedWorkbook = clone(workbook);
    const stats = applyToWorkbook(updatedWorkbook, rows);
    return { workbook: updatedWorkbook, ...stats, dateRange: { start: report.start, end: report.end } };
  } finally {
    fs.rmSync(report.temporaryDirectory, { recursive: true, force: true });
  }
}

module.exports = { runActualizarData, formatReport, applyToWorkbook };
