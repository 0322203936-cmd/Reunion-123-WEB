# Stops the system cleanly (pending changes saved to Neon first). To turn it on again:
#   Start-ScheduledTask -TaskName 'Reunion123 Servidor'   (or restart the PC)
$Root = Split-Path -Parent $PSScriptRoot
New-Item -ItemType File -Force (Join-Path $Root 'data\.apagar') | Out-Null
Write-Host ' Apagando el sistema (guarda lo pendiente en Neon primero)...'
