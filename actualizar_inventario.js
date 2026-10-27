const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

const POSCO_URL = 'http://3.132.9.174/Posco/';
const SHEET_NAME = 'Inventario';
const SOURCE_HEADERS = ['flor', 'color', 'tallos_por_ramo', 'tallos_en_existencia'];

function requiredEnv(name) {
  const value = String(process.env[name] || '').trim();
  if (!value) throw new Error(`Falta configurar ${name} en el entorno local.`);
  return value;
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function cleanValue(value) {
  if (value === null || value === undefined) return '';
  return value;
}

async function launchBrowser() {
  const headless = String(process.env.POSCO_HEADLESS || 'true').trim().toLowerCase() !== 'false';
  const executablePath = String(process.env.POSCO_BROWSER_EXECUTABLE_PATH || '').trim();
  const channel = String(process.env.POSCO_BROWSER_CHANNEL || 'msedge').trim();
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

async function clickExactText(page, value) {
  const candidates = [
    page.getByRole('button', { name: value, exact: true }),
    page.getByRole('link', { name: value, exact: true }),
    page.getByText(value, { exact: true })
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
  throw new Error(`No se encontró el control visible ${value} en POSCO.`);
}

async function downloadInventoryReport() {
  const user = requiredEnv('POSCO_USER');
  const password = requiredEnv('POSCO_PASSWORD');
  const browser = await launchBrowser();
  const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1920, height: 1080 } });
  const page = await context.newPage();
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'reunion123-inventario-'));

  try {
    console.log('Abriendo POSCO para Actualizar Inventario...');
    await page.goto(POSCO_URL, { waitUntil: 'networkidle', timeout: 60000 });
    await page.locator('input[placeholder*="usuario@email.com" i], input[type="text"]').first().fill(user);
    await page.locator('input[placeholder*="Password" i], input[type="password"]').first().fill(password);
    await page.getByRole('button', { name: 'Iniciar Sesión' }).click({ timeout: 15000 });
    await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(2000);

    await clickExactText(page, 'Inventario');
    await page.waitForTimeout(700);
    await clickExactText(page, 'Inventario General');
    await page.waitForURL('**/#/list-inventario-gral', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(Number(process.env.POSCO_INVENTARIO_WAIT_MS || 15000));

    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
    await clickExactText(page, 'Exportar');
    const download = await downloadPromise;
    const filename = path.basename(download.suggestedFilename() || 'inventario_general.xlsx');
    const reportPath = path.join(temporaryDirectory, filename);
    await download.saveAs(reportPath);
    console.log(`Inventario General descargado: ${reportPath}`);
    return { path: reportPath, temporaryDirectory };
  } catch (error) {
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    throw error;
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

function readInventoryRows(reportPath) {
  const sourceWorkbook = XLSX.readFile(reportPath, { cellDates: false, raw: true });
  const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
  const sourceRows = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, raw: true, defval: '' });
  if (!sourceRows.length) throw new Error('El reporte de Inventario General está vacío.');

  const headers = sourceRows[0].map(value => String(value ?? '').trim().toLowerCase());
  const missing = SOURCE_HEADERS.filter(header => !headers.includes(header));
  if (missing.length) throw new Error(`Faltan columnas en el reporte de Inventario: ${missing.join(', ')}.`);

  const indexes = SOURCE_HEADERS.map(header => headers.indexOf(header));
  return sourceRows.slice(1).map(row => indexes.map(index => cleanValue(row[index])));
}

function setCell(sheet, coord, value, formula) {
  if (!sheet.cells) sheet.cells = {};
  if (!sheet.cells[coord]) {
    const column = coord.match(/^[A-Z]+/)?.[0];
    const templateStyle = column ? sheet.cells[`${column}2`]?.s : null;
    sheet.cells[coord] = templateStyle ? { s: clone(templateStyle) } : {};
  }
  const cell = sheet.cells[coord];
  if (value === undefined || value === '') delete cell.v;
  else cell.v = value;
  if (formula === undefined) delete cell.f;
  else cell.f = formula;
}

function clearCell(sheet, coord) {
  const cell = sheet.cells?.[coord];
  if (!cell) return;
  delete cell.v;
  delete cell.f;
}

function applyToInventory(workbook, rows) {
  const sheet = workbook.sheets?.[SHEET_NAME];
  if (!sheet) throw new Error('No existe la hoja Inventario en el sistema.');
  const header = column => String(sheet.cells?.[`${column}1`]?.v ?? '').trim().toLowerCase();
  const expected = { A: 'flor', B: 'color', D: 'tallos_por_ramo', E: 'tallos_en_existencia' };
  for (const [column, expectedValue] of Object.entries(expected)) {
    if (header(column) !== expectedValue) {
      throw new Error(`El encabezado ${column}1 de Inventario no coincide: se esperaba ${expectedValue}.`);
    }
  }

  const oldLastRow = Math.max(Number(sheet.maxRow) || 1, 1);
  const newLastRow = rows.length + 1;
  const lastRowToClear = Math.max(oldLastRow, newLastRow);

  for (let row = 2; row <= lastRowToClear; row += 1) {
    for (const column of ['A', 'B', 'C', 'D', 'E', 'F']) clearCell(sheet, `${column}${row}`);
  }

  for (let index = 0; index < rows.length; index += 1) {
    const rowNumber = index + 2;
    const [flor, color, tallosPorRamo, tallosEnExistencia] = rows[index];
    setCell(sheet, `A${rowNumber}`, flor);
    setCell(sheet, `B${rowNumber}`, color);
    setCell(sheet, `C${rowNumber}`, '', `=+A${rowNumber} & " " & B${rowNumber}`);
    setCell(sheet, `D${rowNumber}`, tallosPorRamo);
    setCell(sheet, `E${rowNumber}`, tallosEnExistencia);
    setCell(sheet, `F${rowNumber}`, '', `=+A${rowNumber} & " " & B${rowNumber} & "X" & " " & D${rowNumber}`);
  }

  sheet.maxRow = Math.max(oldLastRow, newLastRow + 20, 100);
  sheet.maxCol = Math.max(Number(sheet.maxCol) || 0, 6);
  return { rows: rows.length, lastRow: newLastRow };
}

async function runActualizarInventario(workbook) {
  const report = await downloadInventoryReport();
  try {
    const rows = readInventoryRows(report.path);
    const updatedWorkbook = clone(workbook);
    const stats = applyToInventory(updatedWorkbook, rows);
    return { workbook: updatedWorkbook, ...stats };
  } finally {
    fs.rmSync(report.temporaryDirectory, { recursive: true, force: true });
  }
}

module.exports = { runActualizarInventario, readInventoryRows, applyToInventory };
