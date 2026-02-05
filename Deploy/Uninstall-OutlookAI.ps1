#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Uninstalls OutlookAI add-in from all users

.DESCRIPTION
    This script removes the OutlookAI Outlook add-in for all users.
    Must be run as Administrator.

.PARAMETER InstallPath
    Where OutlookAI is installed. Defaults to C:\Program Files\OutlookAI

.EXAMPLE
    .\Uninstall-OutlookAI.ps1
#>

param(
    [string]$InstallPath = "C:\Program Files\OutlookAI"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI Uninstaller" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

# Step 1: Remove registry entries
Write-Host "[1/2] Removing registry entries..." -ForegroundColor Yellow

$addinPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI"
if (Test-Path $addinPath) {
    Remove-Item -Path $addinPath -Force
    Write-Host "  Removed: $addinPath" -ForegroundColor Gray
}

$addinPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
if (Test-Path $addinPath32) {
    Remove-Item -Path $addinPath32 -Force
    Write-Host "  Removed: $addinPath32" -ForegroundColor Gray
}

Write-Host "  Done." -ForegroundColor Green

# Step 2: Remove files
Write-Host "[2/2] Removing files..." -ForegroundColor Yellow

if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
    Write-Host "  Removed: $InstallPath" -ForegroundColor Gray
}

Write-Host "  Done." -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Uninstall Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Users will need to restart Outlook for changes to take effect." -ForegroundColor White
Write-Host ""
