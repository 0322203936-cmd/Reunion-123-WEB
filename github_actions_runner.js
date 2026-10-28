const { Pool } = require('pg');
const { runActualizarData } = require('./actualizar_data');
const { runActualizarInventario } = require('./actualizar_inventario');

const WORKBOOK_PREFIX = 'reunion123';
const OPERATION = String(process.argv[2] || '').trim().toLowerCase();

function requiredEnv(name) {
  const value = String(process.env[name] || '').trim();
  if (!value) throw new Error(`Falta configurar ${name} en GitHub Actions.`);
  return value;
}

function splitWorkbook(workbook) {
  const { sheets = {}, ...meta } = workbook;
  return { meta, sheets };
}

async function loadWorkbook(pool, prefix = WORKBOOK_PREFIX) {
  const [metaResult, sheetsResult] = await Promise.all([
    pool.query(`SELECT meta FROM ${prefix}_workbook_meta WHERE id = 1`),
    pool.query(`SELECT name, data FROM ${prefix}_workbook_sheets`)
  ]);
  if (!metaResult.rows[0]) throw new Error(`No existe la configuración de ${prefix} en Neon.`);
  const sheets = {};
  for (const row of sheetsResult.rows) sheets[row.name] = row.data;
  return { ...metaResult.rows[0].meta, sheets };
}

async function saveWorkbookSheets(pool, workbook, sheetNames, prefix = WORKBOOK_PREFIX) {
  const { meta } = splitWorkbook(workbook);
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (const name of sheetNames) {
      const sheet = workbook.sheets?.[name];
      if (!sheet) continue;
      await client.query(
        `INSERT INTO ${prefix}_workbook_sheets (name, data) VALUES ($1, $2::jsonb)
         ON CONFLICT (name) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()`,
        [name, JSON.stringify(sheet)]
      );
    }
    await client.query(
      `INSERT INTO ${prefix}_workbook_meta (id, meta) VALUES (1, $1::jsonb)
       ON CONFLICT (id) DO UPDATE SET meta = EXCLUDED.meta, updated_at = NOW()`,
      [JSON.stringify(meta)]
    );
    await client.query('COMMIT');
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}

async function notifyRender(payload) {
  const callbackUrl = String(process.env.REUNION_CALLBACK_URL || '').trim();
  const token = String(process.env.REUNION_SYNC_TOKEN || '').trim();
  if (!callbackUrl || !token) {
    console.log('Callback de Render no configurado; datos guardados en Neon.');
    return;
  }
  const response = await fetch(callbackUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw new Error(`Render callback respondió HTTP ${response.status}.`);
}

async function run() {
  if (!['data', 'inventario'].includes(OPERATION)) {
    throw new Error('La operación debe ser data o inventario.');
  }
  const pool = new Pool({
    connectionString: requiredEnv('DATABASE_URL'),
    ssl: { rejectUnauthorized: false },
    max: 2,
    connectionTimeoutMillis: 15000
  });
  try {
    const reunion = await loadWorkbook(pool);
    const result = OPERATION === 'data'
      ? await runActualizarData(reunion)
      : await runActualizarInventario(reunion);
    const targetSheet = OPERATION === 'data' ? 'PegarData' : 'Inventario';
    await saveWorkbookSheets(pool, result.workbook, [targetSheet]);
    await notifyRender({
      operation: OPERATION,
      status: 'succeeded',
      rows: result.rows,
      lastRow: result.lastRow,
      range: OPERATION === 'data' ? `A1:W${result.lastRow}` : null,
      dateRange: result.dateRange ? {
        start: result.dateRange.start.toISOString().slice(0, 10),
        end: result.dateRange.end.toISOString().slice(0, 10)
      } : null
    });
    console.log(`POSCO ${OPERATION} completado: ${result.rows} filas actualizadas en Neon.`);
  } catch (error) {
    try {
      await notifyRender({ operation: OPERATION, status: 'failed', error: error.message || 'Error desconocido.' });
    } catch (callbackError) {
      console.error('No se pudo avisar a Render:', callbackError.message);
    }
    throw error;
  } finally {
    await pool.end();
  }
}

run().catch(error => {
  console.error(`POSCO ${OPERATION || 'update'} falló:`, error);
  process.exitCode = 1;
});
