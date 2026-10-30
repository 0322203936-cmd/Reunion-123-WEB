# Backup of the local database of Reunion 123 (scheduled task "Reunion123 Respaldo", see
# tools/instalar_respaldo.ps1): respaldos\reunion123_AAAAMMDD_HHmm.dump, the last 14 are kept.
# Restore one (system stopped):
#   & 'C:\Program Files\PostgreSQL\18\bin\pg_restore.exe' --clean --if-exists --no-owner --dbname=<DATABASE_URL> <archivo.dump>
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Backups = Join-Path $Root 'respaldos'
$Log = Join-Path $Root 'logs\respaldo.log'
New-Item -ItemType Directory -Force $Backups, (Join-Path $Root 'logs') | Out-Null
function Log($msg) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg" | Add-Content -Path $Log -Encoding utf8 }

$line = Get-Content (Join-Path $Root '.env') | Where-Object { $_ -match '^DATABASE_URL=' } | Select-Object -First 1
$url = $line.Substring('DATABASE_URL='.Length).Trim().Trim('"')
if ($url -notmatch '@(localhost|127\.0\.0\.1)[:/]') { Log 'DATABASE_URL no es la base local: sin respaldo'; exit 0 }

$file = Join-Path $Backups ("reunion123_{0}.dump" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
& 'C:\Program Files\PostgreSQL\18\bin\pg_dump.exe' --format=custom --no-owner --no-privileges --file=$file --dbname=$url
if ($LASTEXITCODE -ne 0) { Log "pg_dump fallo ($LASTEXITCODE)"; Remove-Item $file -ErrorAction SilentlyContinue; exit 1 }
Log ("Respaldo {0} ({1:N1} MB)" -f (Split-Path $file -Leaf), ((Get-Item $file).Length / 1MB))

Get-ChildItem $Backups -Filter 'reunion123_*.dump' | Sort-Object Name -Descending | Select-Object -Skip 14 | Remove-Item
