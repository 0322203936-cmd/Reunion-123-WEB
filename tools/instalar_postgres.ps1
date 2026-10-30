# Installs PostgreSQL 18 on the server PC for Reunion 123 (run once, as administrator):
#   * service "postgresql-x64-18", starts with Windows, listens ONLY inside this PC (localhost)
#   * database "reunion123" owned by the login "reunion"
#   * passwords are generated here and written only to .env (never shown):
#       PG_SUPERPASSWORD     superuser "postgres"
#       DATABASE_URL_LOCAL   connection of the system (tools/migrar_a_local.ps1 makes it DATABASE_URL)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$EnvFile = Join-Path $Root '.env'
$Bin = 'C:\Program Files\PostgreSQL\18\bin'
$Service = 'postgresql-x64-18'

function New-Password {
  $chars = [char[]]'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789'
  $bytes = New-Object byte[] 32
  [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
  -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
}
function Get-EnvValue($name) {
  $line = Get-Content $EnvFile | Where-Object { $_ -match "^$name=" } | Select-Object -First 1
  if ($line) { $line.Substring($name.Length + 1).Trim() } else { $null }
}
function Set-EnvValue($name, $value) {
  $lines = @(Get-Content $EnvFile | Where-Object { $_ -notmatch "^$name=" })
  $lines += "$name=$value"
  Set-Content -Path $EnvFile -Value $lines -Encoding ascii
}

# 1) install (the superuser password is kept in .env so a second run can reuse it)
$super = Get-EnvValue 'PG_SUPERPASSWORD'
if (-not $super) { $super = New-Password; Set-EnvValue 'PG_SUPERPASSWORD' $super }
if (-not (Test-Path "$Bin\psql.exe")) {
  Write-Host ' Instalando PostgreSQL 18 (unos minutos)...'
  # over SSH Windows reports the domain as WORKGROUP and the installer rejects the account
  $env:USERDOMAIN = $env:COMPUTERNAME
  $override = "--mode unattended --unattendedmodeui none --superpassword $super --serverport 5432 " +
    "--enable-components server,commandlinetools --disable-components pgAdmin,stackbuilder"
  winget install --id PostgreSQL.PostgreSQL.18 --exact --silent --accept-package-agreements --accept-source-agreements --override $override | Out-Null
  if (-not (Test-Path "$Bin\psql.exe")) { throw 'La instalación de PostgreSQL no terminó bien.' }
}
Set-Service $Service -StartupType Automatic
if ((Get-Service $Service).Status -ne 'Running') { Start-Service $Service }

# 2) only reachable from this PC
$conf = 'C:\Program Files\PostgreSQL\18\data\postgresql.conf'
$text = Get-Content $conf -Raw
$fixed = [regex]::Replace($text, "(?m)^\s*#?\s*listen_addresses\s*=.*$", "listen_addresses = 'localhost'")
if ($fixed -ne $text) {
  Set-Content -Path $conf -Value $fixed -Encoding ascii
  Restart-Service $Service
}
Get-NetFirewallRule -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Postgre' } | Remove-NetFirewallRule

# 3) login + database of the system
$env:PGPASSWORD = $super
$appPass = $null
$url = Get-EnvValue 'DATABASE_URL_LOCAL'
if ($url -match '^postgres://reunion:([^@]+)@') { $appPass = $Matches[1] }
if (-not $appPass) { $appPass = New-Password }
$exists = & "$Bin\psql.exe" -h localhost -U postgres -tAc "select 1 from pg_roles where rolname='reunion'"
if ($exists -eq '1') { & "$Bin\psql.exe" -h localhost -U postgres -qc "alter role reunion with login password '$appPass'" }
else { & "$Bin\psql.exe" -h localhost -U postgres -qc "create role reunion with login password '$appPass'" }
$db = & "$Bin\psql.exe" -h localhost -U postgres -tAc "select 1 from pg_database where datname='reunion123'"
if ($db -ne '1') { & "$Bin\psql.exe" -h localhost -U postgres -qc "create database reunion123 owner reunion encoding 'UTF8'" }
Remove-Item Env:PGPASSWORD
Set-EnvValue 'DATABASE_URL_LOCAL' "postgres://reunion:$appPass@localhost:5432/reunion123"

Write-Host ''
Write-Host ' Listo: PostgreSQL 18 instalado, solo accesible desde esta PC.'
Write-Host ' Base "reunion123" creada; su conexion quedo en .env como DATABASE_URL_LOCAL.'
