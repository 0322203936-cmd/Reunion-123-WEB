# Runs on YOUR computer: scheduled task "Reunion123 Traer respaldos" that runs tools/traer_respaldos.ps1
# when you log in and every 4 hours while the computer is on.
#   powershell -ExecutionPolicy Bypass -File tools\instalar_traer_respaldos.ps1
$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot 'traer_respaldos.ps1'
$name = 'Reunion123 Traer respaldos'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$every4h = New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Hours 4)
$atLogon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 20) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $name -Action $action -Trigger @($every4h, $atLogon) -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $name
Write-Host " Listo: '$name' instalada. Los respaldos llegan a $env:USERPROFILE\Respaldos Reunion123" -ForegroundColor Green
