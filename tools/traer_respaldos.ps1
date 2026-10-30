# Runs on YOUR computer (not the server): brings the database backups of the server PC
# (C:\reunion123\respaldos) to "Respaldos Reunion123" in your user folder, through Tailscale + SSH.
# Copies kept 30 days. If the server PC is off or unreachable it just tries next time.
#   powershell -ExecutionPolicy Bypass -File tools\traer_respaldos.ps1
$ErrorActionPreference = 'Continue'
$Dest = Join-Path $env:USERPROFILE 'Respaldos Reunion123'
$Key = Join-Path $env:USERPROFILE '.ssh\reunion_pc'
$Server = 'CFBC@desktop-289vqpm'
New-Item -ItemType Directory -Force $Dest | Out-Null
$Log = Join-Path $Dest 'traer_respaldos.log'
function Log($msg) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg" | Add-Content -Path $Log -Encoding utf8 }

$list = & ssh -i $Key -o BatchMode=yes -o ConnectTimeout=15 $Server "Get-ChildItem C:\reunion123\respaldos -Filter *.dump | ForEach-Object Name" 2>$null
if ($LASTEXITCODE -ne 0) { Log 'PC servidor no disponible; se reintenta la proxima vez'; exit 0 }
$new = 0
foreach ($name in $list) {
  $name = $name.Trim()
  if (-not $name -or (Test-Path (Join-Path $Dest $name))) { continue }
  & scp -i $Key -o BatchMode=yes -q "${Server}:C:/reunion123/respaldos/$name" (Join-Path $Dest $name)
  if ($LASTEXITCODE -eq 0) { $new++ } else { Log "No se pudo copiar $name" }
}
Get-ChildItem $Dest -Filter *.dump | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item
Log "Respaldos nuevos: $new"
