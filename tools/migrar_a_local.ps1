# Moves the data of Reunion 123 from Neon to the PostgreSQL of this PC (tools/instalar_postgres.ps1 first):
#   1. stops the system (everything pending is saved to Neon first)
#   2. copies Neon into the local database (a copy of Neon stays in respaldos\)
#   3. compares both table by table (tools/comparar_bd.js)
#   4. only if they are identical: .env uses the local database (Neon stays as NEON_DATABASE_URL)
#   5. starts the system again (with the local database, or with Neon if something failed)
# Going back to Neon: in .env put the NEON_DATABASE_URL value back as DATABASE_URL and restart.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$EnvFile = Join-Path $Root '.env'
$Bin = 'C:\Program Files\PostgreSQL\18\bin'
$Backups = Join-Path $Root 'respaldos'
New-Item -ItemType Directory -Force $Backups | Out-Null

function Get-EnvValue($name) {
  $line = Get-Content $EnvFile | Where-Object { $_ -match "^$name=" } | Select-Object -First 1
  if ($line) { $line.Substring($name.Length + 1).Trim().Trim('"') } else { $null }
}
function Set-EnvValue($name, $value) {
  $lines = @(Get-Content $EnvFile | Where-Object { $_ -notmatch "^$name=" })
  $lines += "$name=$value"
  Set-Content -Path $EnvFile -Value $lines -Encoding ascii
}
function Test-ServerUp { [bool](Get-NetTCPConnection -LocalPort 3070 -State Listen -ErrorAction SilentlyContinue) }
function Start-System {
  Start-ScheduledTask -TaskName 'Reunion123 Servidor'
  for ($i = 0; $i -lt 60 -and -not (Test-ServerUp); $i++) { Start-Sleep -Seconds 2 }
}

$neon = Get-EnvValue 'DATABASE_URL'
$local = Get-EnvValue 'DATABASE_URL_LOCAL'
if (-not $local) { throw 'Falta DATABASE_URL_LOCAL: corre primero tools\instalar_postgres.ps1' }
if ($neon -match '@(localhost|127\.0\.0\.1)[:/]') { throw 'DATABASE_URL ya apunta a la base local: la mudanza ya se hizo.' }
Set-EnvValue 'NEON_DATABASE_URL' $neon

Write-Host ' 1/5 Deteniendo el sistema...'
New-Item -ItemType File -Force (Join-Path $Root 'data\.apagar') | Out-Null
for ($i = 0; $i -lt 60 -and (Test-ServerUp); $i++) { Start-Sleep -Seconds 2 }
if (Test-ServerUp) { throw 'El sistema no se detuvo en 2 minutos; no se hizo ningun cambio.' }
Start-Sleep -Seconds 3

$switched = $false
try {
  Write-Host ' 2/5 Copiando Neon -> base local...'
  $dump = Join-Path $Backups ("neon_antes_de_mudanza_{0}.dump" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
  & "$Bin\pg_dump.exe" --format=custom --schema=public --no-owner --no-privileges --file=$dump --dbname=$neon
  if ($LASTEXITCODE -ne 0) { throw 'pg_dump (copia de Neon) fallo.' }
  & "$Bin\pg_restore.exe" --no-owner --no-privileges --clean --if-exists --exit-on-error --dbname=$local $dump
  if ($LASTEXITCODE -ne 0) { throw 'pg_restore (carga en la base local) fallo.' }

  Write-Host ' 3/5 Comparando Neon y la base local tabla por tabla...'
  & node (Join-Path $Root 'tools\comparar_bd.js') NEON_DATABASE_URL DATABASE_URL_LOCAL
  if ($LASTEXITCODE -ne 0) { throw 'Las bases no quedaron iguales; se sigue usando Neon.' }

  Write-Host ' 4/5 Cambiando el sistema a la base local...'
  Set-EnvValue 'DATABASE_URL' $local
  $switched = $true
} finally {
  Write-Host ' 5/5 Prendiendo el sistema...'
  Start-System
}
if (-not (Test-ServerUp)) { throw 'El sistema no volvio a prender; revisa logs\server.err.log' }
Write-Host ''
if ($switched) { Write-Host ' Listo: Reunion 123 ya trabaja con la base de datos de esta PC. Neon queda intacto como respaldo.' }
