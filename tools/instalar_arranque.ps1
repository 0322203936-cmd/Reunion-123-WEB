# Scheduled task "Reunion123 Servidor": runs tools/servidor.ps1 when the PC turns on, without
# anyone logging in and without a window. Run once, as administrator:
#   powershell -ExecutionPolicy Bypass -File C:\reunion123\tools\instalar_arranque.ps1
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$name = 'Reunion123 Servidor'
$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Root\tools\servidor.ps1`"" -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -AtStartup
# S4U: runs as this user even when nobody is logged in, without storing the password
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType S4U -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
  -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $name
Write-Host " Listo: '$name' instalada y prendida (usuario $user)." -ForegroundColor Green
