// Compares two databases table by table (row count + checksum of every row) and prints the result.
//   node tools/comparar_bd.js NOMBRE_VAR_A NOMBRE_VAR_B      (names of connection variables in .env)
// Exit code 0 when both have the same tables with the same content.
const fs = require('fs');
const path = require('path');
const { Client } = require('pg');

function envValue(name) {
  const text = fs.readFileSync(path.join(__dirname, '..', '.env'), 'utf8');
  const m = new RegExp(`^${name}=(.*)$`, 'm').exec(text);
  if (!m) throw new Error(`Falta ${name} en .env`);
  return m[1].trim().replace(/^"|"$/g, '');
}

async function snapshot(url) {
  const local = /@(localhost|127\.0\.0\.1)[:/]/.test(url);
  const client = new Client({ connectionString: url, ssl: local ? false : { rejectUnauthorized: false } });
  await client.connect();
  try {
    const tables = (await client.query(
      "select table_name from information_schema.tables where table_schema = 'public' and table_type = 'BASE TABLE' order by 1"
    )).rows.map(r => r.table_name);
    const out = {};
    for (const t of tables) {
      const r = await client.query(
        `select count(*)::int as n, coalesce(md5(string_agg(md5(x::text), '' order by md5(x::text))), '') as sum from "${t}" x`
      );
      out[t] = `${r.rows[0].n} filas ${r.rows[0].sum.slice(0, 12)}`;
    }
    return out;
  } finally {
    await client.end();
  }
}

(async () => {
  const [a, b] = process.argv.slice(2);
  const [sa, sb] = await Promise.all([snapshot(envValue(a)), snapshot(envValue(b))]);
  let same = true;
  for (const t of [...new Set([...Object.keys(sa), ...Object.keys(sb)])].sort()) {
    const ok = sa[t] === sb[t];
    if (!ok) same = false;
    console.log(`${ok ? 'OK ' : 'DIF'} ${t.padEnd(28)} ${a}: ${sa[t] || '(no existe)'}   ${b}: ${sb[t] || '(no existe)'}`);
  }
  console.log(same ? 'IGUALES' : 'DIFERENTES');
  process.exit(same ? 0 : 1);
})().catch(err => { console.error('Error:', err.message); process.exit(2); });
