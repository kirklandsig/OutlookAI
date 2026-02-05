#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs OutlookAI add-in for all users on RDS/Terminal Server

.DESCRIPTION
    This script installs the OutlookAI Outlook add-in for all users.
    Must be run as Administrator.

.PARAMETER SourcePath
    Path to the published OutlookAI files (containing OutlookAI.vsto and Application Files folder).
    Defaults to C:\OutlookAI

.PARAMETER InstallPath
    Where to install. Defaults to C:\Program Files\OutlookAI

.EXAMPLE
    .\Install-OutlookAI.ps1

.EXAMPLE
    .\Install-OutlookAI.ps1 -SourcePath "\\fileserver\deploy\OutlookAI"
#>

param(
    [string]$SourcePath = "C:\OutlookAI",
    [string]$InstallPath = "C:\Program Files\OutlookAI"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI Installer for RDS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

# Verify source exists
if (!(Test-Path $SourcePath)) {
    Write-Host "ERROR: Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

# Check for required files
$vstoFile = Join-Path $SourcePath "OutlookAI.vsto"
$appFilesDir = Join-Path $SourcePath "Application Files"

if (!(Test-Path $vstoFile)) {
    Write-Host "ERROR: OutlookAI.vsto not found in $SourcePath" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $appFilesDir)) {
    Write-Host "ERROR: 'Application Files' folder not found in $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Source: $SourcePath" -ForegroundColor Gray
Write-Host "Target: $InstallPath" -ForegroundColor Gray
Write-Host ""

# Step 1: Create/clean install directory
Write-Host "[1/6] Preparing install directory..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Write-Host "  Removing existing installation..." -ForegroundColor Gray
    Remove-Item -Path $InstallPath -Recurse -Force
}
New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
Write-Host "  Done." -ForegroundColor Green

# Step 2: Copy files (entire ClickOnce structure)
Write-Host "[2/6] Copying files..." -ForegroundColor Yellow

# Copy the .vsto manifest
Copy-Item -Path $vstoFile -Destination $InstallPath -Force
Write-Host "  Copied: OutlookAI.vsto" -ForegroundColor Gray

# Copy the entire Application Files folder
Copy-Item -Path $appFilesDir -Destination $InstallPath -Recurse -Force
Write-Host "  Copied: Application Files\" -ForegroundColor Gray

# Also copy and rename .deploy files to root as regular DLLs
$latestVersionDir = Get-ChildItem -Path (Join-Path $InstallPath "Application Files") -Directory | Sort-Object Name -Descending | Select-Object -First 1
if ($latestVersionDir) {
    $deployFiles = Get-ChildItem -Path $latestVersionDir.FullName -Filter "*.deploy"
    foreach ($file in $deployFiles) {
        $newName = $file.Name -replace '\.deploy$', ''
        Copy-Item -Path $file.FullName -Destination (Join-Path $InstallPath $newName) -Force
        Write-Host "  Copied: $newName" -ForegroundColor Gray
    }

    # Also copy the manifest file
    $manifestFile = Join-Path $latestVersionDir.FullName "OutlookAI.dll.manifest"
    if (Test-Path $manifestFile) {
        Copy-Item -Path $manifestFile -Destination $InstallPath -Force
        Write-Host "  Copied: OutlookAI.dll.manifest" -ForegroundColor Gray
    }
}

# Count files
$fileCount = (Get-ChildItem -Path $InstallPath -Recurse -File).Count
Write-Host "  Total files: $fileCount" -ForegroundColor Green

# Step 3: Unblock all files
Write-Host "[3/6] Unblocking files..." -ForegroundColor Yellow
Get-ChildItem -Path $InstallPath -Recurse | Unblock-File -ErrorAction SilentlyContinue
Write-Host "  Done." -ForegroundColor Green

# Step 4: Configure VSTO trust
Write-Host "[4/6] Configuring VSTO trust settings..." -ForegroundColor Yellow

# PromptingLevel settings
$trustPath = "HKLM:\SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
if (!(Test-Path $trustPath)) { New-Item -Path $trustPath -Force | Out-Null }
Set-ItemProperty -Path $trustPath -Name "MyComputer" -Value "Enabled"
Set-ItemProperty -Path $trustPath -Name "LocalIntranet" -Value "Enabled"

# WOW64 (32-bit)
$trustPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
if (!(Test-Path $trustPath32)) { New-Item -Path $trustPath32 -Force | Out-Null }
Set-ItemProperty -Path $trustPath32 -Name "MyComputer" -Value "Enabled"
Set-ItemProperty -Path $trustPath32 -Name "LocalIntranet" -Value "Enabled"

# VSTO Inclusion List
$inclusionPath = "HKLM:\SOFTWARE\Microsoft\VSTO\Security\Inclusion"
if (!(Test-Path $inclusionPath)) { New-Item -Path $inclusionPath -Force | Out-Null }
$guid = [System.Guid]::NewGuid().ToString("B")
$addinKey = Join-Path $inclusionPath $guid
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name "Url" -Value "file:///C:/Program Files/OutlookAI/OutlookAI.vsto"
Set-ItemProperty -Path $addinKey -Name "PublicKey" -Value ""

# WOW64 Inclusion List
$inclusionPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO\Security\Inclusion"
if (!(Test-Path $inclusionPath32)) { New-Item -Path $inclusionPath32 -Force | Out-Null }
$addinKey32 = Join-Path $inclusionPath32 $guid
New-Item -Path $addinKey32 -Force | Out-Null
Set-ItemProperty -Path $addinKey32 -Name "Url" -Value "file:///C:/Program Files/OutlookAI/OutlookAI.vsto"
Set-ItemProperty -Path $addinKey32 -Name "PublicKey" -Value ""

Write-Host "  Done." -ForegroundColor Green

# Step 5: Register add-in for all users
Write-Host "[5/6] Registering add-in for all users..." -ForegroundColor Yellow

$manifestPath = "file:///C:/Program Files/OutlookAI/OutlookAI.vsto|vstolocal"

# 64-bit Office
$addinPath = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI"
if (Test-Path $addinPath) { Remove-Item -Path $addinPath -Force }
New-Item -Path $addinPath -Force | Out-Null
Set-ItemProperty -Path $addinPath -Name "Description" -Value "AI Writing Assistant for Outlook"
Set-ItemProperty -Path $addinPath -Name "FriendlyName" -Value "OutlookAI"
Set-ItemProperty -Path $addinPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $addinPath -Name "Manifest" -Value $manifestPath

# 32-bit Office (WOW64)
$addinPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
if (Test-Path $addinPath32) { Remove-Item -Path $addinPath32 -Force }
New-Item -Path $addinPath32 -Force | Out-Null
Set-ItemProperty -Path $addinPath32 -Name "Description" -Value "AI Writing Assistant for Outlook"
Set-ItemProperty -Path $addinPath32 -Name "FriendlyName" -Value "OutlookAI"
Set-ItemProperty -Path $addinPath32 -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $addinPath32 -Name "Manifest" -Value $manifestPath

Write-Host "  Done." -ForegroundColor Green

# Step 6: Clear Outlook Resiliency for Default User profile (new users)
Write-Host "[6/6] Configuring for auto-load..." -ForegroundColor Yellow

# Load Default User registry hive
$defaultUserPath = "C:\Users\Default\NTUSER.DAT"
$tempKey = "HKU\DefaultUser"

try {
    reg load $tempKey $defaultUserPath 2>$null

    # Set add-in to load for new user profiles
    $defaultAddinPath = "Registry::$tempKey\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI"
    if (!(Test-Path $defaultAddinPath)) {
        New-Item -Path $defaultAddinPath -Force | Out-Null
    }
    Set-ItemProperty -Path $defaultAddinPath -Name "LoadBehavior" -Value 3 -Type DWord

    reg unload $tempKey 2>$null
    Write-Host "  Configured default user profile." -ForegroundColor Gray
}
catch {
    Write-Host "  Note: Could not modify default user profile (may need reboot)." -ForegroundColor Gray
}

Write-Host "  Done." -ForegroundColor Green

# Verify
Write-Host ""
Write-Host "Verifying installation..." -ForegroundColor Yellow
$verified = $true

if (!(Test-Path (Join-Path $InstallPath "OutlookAI.vsto"))) {
    Write-Host "  WARNING: OutlookAI.vsto not found!" -ForegroundColor Red
    $verified = $false
}

if (!(Test-Path (Join-Path $InstallPath "Application Files"))) {
    Write-Host "  WARNING: Application Files folder not found!" -ForegroundColor Red
    $verified = $false
}

if ($verified) {
    Write-Host "  All checks passed." -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Users will see the add-in when they open Outlook." -ForegroundColor White
Write-Host "The 'AI Assistant' button appears in the ribbon when composing a new email." -ForegroundColor White
Write-Host ""
Write-Host "To verify: Outlook > File > Options > Add-ins" -ForegroundColor Gray
Write-Host "To uninstall: Run Uninstall-OutlookAI.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "If add-in doesn't auto-load for existing users:" -ForegroundColor Yellow
Write-Host "  - User runs: Enable-OutlookAI-User.ps1" -ForegroundColor Gray
Write-Host "  - Or set as logon script via Group Policy" -ForegroundColor Gray
Write-Host ""
