#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Uninstalls OutlookAI v2 from this server.

.DESCRIPTION
    Removes:
      - HKLM Outlook add-in registration (64-bit + WOW6432Node)
      - C:\Program Files\OutlookAI install directory
      - C:\ProgramData\OutlookAI\auth.json + sidecar refresh lock
        (clears the shared ChatGPT OAuth credentials)

    Preserves:
      - C:\ProgramData\OutlookAI\Backups (v1 config rollback artifacts)

    Phase 1 does NOT perform server-side OAuth token revocation. To
    fully revoke credentials with OpenAI, sign the account out from
    https://chatgpt.com/#settings or rotate by signing in fresh.

    Run as Administrator.

.EXAMPLE
    .\Uninstall-OutlookAI.ps1
#>

param()

$ErrorActionPreference = "Stop"
$InstallPath     = "C:\Program Files\OutlookAI"
$ProgramDataPath = "C:\ProgramData\OutlookAI"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI v2 Uninstaller" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

# --- 1. Registry ---------------------------------------------------------
Write-Host "[1/3] Removing add-in registry entries..." -ForegroundColor Yellow
foreach ($path in @(
    "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
)) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        Write-Host "  Removed: $path" -ForegroundColor Gray
    }
}
Write-Host "  Done." -ForegroundColor Green

# --- 2. Install directory ------------------------------------------------
Write-Host "[2/3] Removing install files..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
    Write-Host "  Removed: $InstallPath" -ForegroundColor Gray
} else {
    Write-Host "  Already absent: $InstallPath" -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

# --- 3. Local OAuth credentials (preserve Backups) -----------------------
Write-Host "[3/3] Removing local OAuth credentials (preserving Backups)..." -ForegroundColor Yellow
$authFile  = Join-Path $ProgramDataPath "auth.json"
$lockFile  = Join-Path $ProgramDataPath "auth.json.refresh.lock"
$tempFile  = Join-Path $ProgramDataPath "auth.json.tmp"
foreach ($path in @($authFile, $lockFile, $tempFile)) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        Write-Host "  Removed: $path" -ForegroundColor Gray
    }
}
Write-Host "  Preserved backups under $ProgramDataPath\Backups" -ForegroundColor Gray
Write-Host "  Done." -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Uninstall Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Users will need to restart Outlook for changes to take effect." -ForegroundColor White
Write-Host ""
Write-Host "NOTE: Phase 1 does not revoke OAuth tokens server-side." -ForegroundColor Yellow
Write-Host "      To fully revoke, sign the ChatGPT account out at" -ForegroundColor Yellow
Write-Host "      https://chatgpt.com/#settings or rotate the credential" -ForegroundColor Yellow
Write-Host "      by signing in again on a fresh install." -ForegroundColor Yellow
Write-Host ""
