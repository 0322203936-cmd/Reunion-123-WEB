# Keeps the system running on this PC and up to date with GitHub (started by the scheduled task
# "Reunion123 Servidor", see tools/instalar_arranque.ps1):
#   * starts node server.js and starts it again if it stops
#   * every 2 minutes looks at GitHub; when there is something new: stops the server cleanly
#     (everything saved to Neon first), git pull, npm install if the libraries changed, starts again
#   * data\.apagar (tools/apagar.ps1) stops the server and this loop
# Logs: logs\servidor.log (this loop), logs\server.out.log / server.err.log (the server)
$ErrorActionPreference = 'Continue'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$Logs = Join-Path $Root 'logs'
New-Item -ItemType Directory -Force $Logs | Out-Null
$LoopLog = Join-Path $Logs 'servidor.log'
$StopServer = Join-Path $Root 'data\.detener'   # read by server.js: save and exit
$StopAll = Join-Path $Root 'data\.apagar'
$CheckEvery = 120

function Log($msg) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg" | Add-Content -Path $LoopLog -Encoding utf8 }

# only one copy of this loop
$mutex = New-Object System.Threading.Mutex($false, 'Global\Reunion123Servidor')
if (-not $mutex.WaitOne(0)) { exit 0 }

$node = (Get-Command node -ErrorAction Stop).Source
$npm = Join-Path (Split-Path $node) 'npm.cmd'

function Start-App {
  foreach ($f in 'server.out.log', 'server.err.log') {
    $p = Join-Path $Logs $f
    if (Test-Path $p) { Move-Item $p "$p.anterior" -Force }
  }
  Remove-Item $StopServer -ErrorAction SilentlyContinue
  $proc = Start-Process -FilePath $node -ArgumentList '--max-old-space-size=4096', 'server.js' -WorkingDirectory $Root `
    -RedirectStandardOutput (Join-Path $Logs 'server.out.log') -RedirectStandardError (Join-Path $Logs 'server.err.log') `
    -WindowStyle Hidden -PassThru
  Log "Sistema prendido (proceso $($proc.Id), versión $(git rev-parse --short HEAD))"
  return $proc
}

function Stop-App($proc) {
  if (-not $proc -or $proc.HasExited) { return }
  New-Item -ItemType File -Force $StopServer | Out-Null
  if (-not $proc.WaitForExit(20000)) {
    Log 'El sistema no se apagó solo en 20 s; se cierra a la fuerza'
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
  }
  Remove-Item $StopServer -ErrorAction SilentlyContinue
}

function Update-Available {
  git fetch -q origin main 2>&1 | Out-Null
  if ($LASTEXITCODE -ne 0) { return $false }
  return (git rev-parse HEAD) -ne (git rev-parse origin/main)
}

Remove-Item $StopAll -ErrorAction SilentlyContinue
Log 'Arranque'
$app = Start-App
$nextCheck = (Get-Date).AddSeconds(30)
while ($true) {
  Start-Sleep -Seconds 2
  if (Test-Path $StopAll) {
    Remove-Item $StopAll -ErrorAction SilentlyContinue
    Log 'Apagado pedido'
    Stop-App $app
    break
  }
  if ($app.HasExited) {
    Log "El sistema se detuvo (código $($app.ExitCode)); se vuelve a prender en 5 s"
    Start-Sleep -Seconds 5
    $app = Start-App
  }
  if ((Get-Date) -ge $nextCheck) {
    $nextCheck = (Get-Date).AddSeconds($CheckEvery)
    if (Update-Available) {
      $changed = git diff --name-only HEAD origin/main
      Log "Actualización encontrada: $(git log --oneline -1 origin/main)"
      Stop-App $app
      $out = git pull --ff-only -q origin main 2>&1
      if ($LASTEXITCODE -ne 0) { Log "git pull falló: $out" }
      elseif ($changed -match '^package(-lock)?\.json$') {
        Log 'Librerías nuevas: npm install'
        & $npm install --no-audit --no-fund 2>&1 | Out-Null
      }
      $app = Start-App
    }
  }
}
$mutex.ReleaseMutex()
