# Turns on SSH in this PC so it can be managed from the other computer through Tailscale.
#   * key only (no passwords): just the key below can enter
#   * port 22 open only for Tailscale addresses (100.64.0.0/10), not for the local network
# Run once, in PowerShell as administrator:
#   powershell -ExecutionPolicy Bypass -File C:\reunion123\tools\instalar_ssh.ps1
$ErrorActionPreference = 'Stop'
$key = 'ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIMYQIcZ/AEkW1JvYMN6kkggWnXafh4Vlt4Gi7HlgMGFU claude-reunion123'

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { Write-Host ' Abre PowerShell como administrador y vuelve a correr esto.' -ForegroundColor Red; exit 1 }

Write-Host ' Instalando OpenSSH Server (puede tardar unos minutos)...'
$cap = Get-WindowsCapability -Online -Name 'OpenSSH.Server*'
if ($cap.State -ne 'Installed') { Add-WindowsCapability -Online -Name $cap.Name | Out-Null }

# the key, for administrators (C:\ProgramData\ssh) and for this user (~\.ssh)
New-Item -ItemType Directory -Force "$env:ProgramData\ssh" | Out-Null
$adminKeys = "$env:ProgramData\ssh\administrators_authorized_keys"
Set-Content -Path $adminKeys -Value $key -Encoding ascii
icacls $adminKeys /inheritance:r /grant 'Administrators:F' /grant 'SYSTEM:F' | Out-Null
New-Item -ItemType Directory -Force "$env:USERPROFILE\.ssh" | Out-Null
Set-Content -Path "$env:USERPROFILE\.ssh\authorized_keys" -Value $key -Encoding ascii

# first start creates sshd_config; then: keys only
Start-Service sshd
$config = "$env:ProgramData\ssh\sshd_config"
$text = Get-Content $config -Raw
$text = $text -replace '(?m)^#?\s*PasswordAuthentication\s+\S+', 'PasswordAuthentication no'
$text = $text -replace '(?m)^#?\s*PubkeyAuthentication\s+\S+', 'PubkeyAuthentication yes'
if ($text -notmatch '(?m)^PasswordAuthentication no') { $text = "PasswordAuthentication no`r`n" + $text }
Set-Content -Path $config -Value $text -Encoding ascii
Set-Service -Name sshd -StartupType Automatic
Restart-Service sshd

# PowerShell as the shell of SSH sessions
New-ItemProperty -Path 'HKLM:\SOFTWARE\OpenSSH' -Name DefaultShell -Value "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -PropertyType String -Force | Out-Null

# port 22 only from Tailscale
Get-NetFirewallRule -Name 'OpenSSH-Server-In-TCP' -ErrorAction SilentlyContinue | Remove-NetFirewallRule
Get-NetFirewallRule -DisplayName 'SSH por Tailscale' -ErrorAction SilentlyContinue | Remove-NetFirewallRule
New-NetFirewallRule -DisplayName 'SSH por Tailscale' -Direction Inbound -Protocol TCP -LocalPort 22 -RemoteAddress 100.64.0.0/10 -Action Allow | Out-Null

Write-Host ''
Write-Host " Listo. Usuario de esta PC: $env:USERNAME" -ForegroundColor Green
Write-Host " Nombre en Tailscale: $(try { (& tailscale status --json | ConvertFrom-Json).Self.DNSName } catch { "(revisa el icono de Tailscale)" })" -ForegroundColor Green
