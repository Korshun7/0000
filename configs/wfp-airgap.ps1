#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Allow MailVault AI processes to talk only to loopback.

.DESCRIPTION
    Does not change rules for OUTLOOK.EXE. Adjust $InstallRoot before use.
    Apply once after MSI install. Re-run after moving binaries.
#>

$ErrorActionPreference = 'Stop'
$InstallRoot = 'C:\Program Files\MailVault'

$binaries = @(
    Join-Path $InstallRoot 'MailVault.Orchestrator.exe'
    Join-Path $InstallRoot 'MailVault.Mapi.exe'
    Join-Path $InstallRoot 'llama-server.exe'
)

foreach ($path in $binaries) {
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Warning "Skip missing $path"
        continue
    }

    $name = Split-Path $path -Leaf

    Get-NetFirewallApplicationFilter -PolicyStore ActiveStore -ErrorAction SilentlyContinue |
        Where-Object { $_.Program -eq $path } |
        ForEach-Object { $_ }

    New-NetFirewallRule -DisplayName "MailVault allow loopback TCP $name" `
        -Direction Outbound -Action Allow -Protocol TCP `
        -RemoteAddress 127.0.0.1 -Program $path -Profile Any | Out-Null

    New-NetFirewallRule -DisplayName "MailVault allow loopback UDP $name" `
        -Direction Outbound -Action Allow -Protocol UDP `
        -RemoteAddress 127.0.0.1 -Program $path -Profile Any | Out-Null

    New-NetFirewallRule -DisplayName "MailVault deny outbound $name" `
        -Direction Outbound -Action Block `
        -Program $path -Profile Any | Out-Null

    Write-Host "WFP rules applied for $path"
}

Write-Host 'Outlook.exe was not modified. Verify llama-server listens on 127.0.0.1 only.'
