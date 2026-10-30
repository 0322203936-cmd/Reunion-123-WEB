# Scheduled task "Reunion123 Respaldo": runs tools/respaldo_bd.ps1 every day at 2:00 a.m. and at 2:00 p.m.
# (if the PC was off at that time, it runs as soon as it turns on). Run once, as administrator:
#   powershell -ExecutionPolicy Bypass -File C:\reunion123\tools\instalar_respaldo.ps1
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$name = 'Reunion123 Respaldo'
$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Root\tools\respaldo_bd.ps1`"" -WorkingDirectory $Root
$triggers = @((New-ScheduledTaskTrigger -Daily -At '02:00'), (New-ScheduledTaskTrigger -Daily -At '14:00'))
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType S4U -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $name -Action $action -Trigger $triggers -Principal $principal -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $name
Write-Host " Listo: '$name' instalada (2:00 y 14:00, usuario $user); primer respaldo en respaldos\." -ForegroundColor Green
