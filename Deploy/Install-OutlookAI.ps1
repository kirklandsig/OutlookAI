#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs OutlookAI v2 (ChatGPT OAuth) for all users on RDS / Terminal Server.

.DESCRIPTION
    Phase 1 v2 install:
      - Hardcoded install path: C:\Program Files\OutlookAI
      - Backs up any v1 config to C:\ProgramData\OutlookAI\Backups
      - Writes a fresh v2 config when one is missing
      - Creates the shared OAuth auth directory at C:\ProgramData\OutlookAI
        with Authenticated Users: Modify (accepted shared-credential risk)
      - Renames any per-user v1 %APPDATA%\OutlookAI\config.xml to a backup
        so per-user files don't override the new server-authoritative
        Model / MaxTokens / CodexAuthPath settings

    Run as Administrator.

.PARAMETER SourcePath
    Path to the published OutlookAI files (containing OutlookAI.vsto and the
    Application Files folder). Defaults to C:\OutlookAI.

.EXAMPLE
    .\Install-OutlookAI.ps1
#>

param(
    [string]$SourcePath = "C:\OutlookAI"
)

$ErrorActionPreference = "Stop"
$InstallPath        = "C:\Program Files\OutlookAI"
$ProgramDataPath    = "C:\ProgramData\OutlookAI"
$BackupRoot         = Join-Path $ProgramDataPath "Backups"
$ConfigFilePath     = Join-Path $InstallPath "config.xml"
$AuthFilePath       = Join-Path $ProgramDataPath "auth.json"
$Timestamp          = Get-Date -Format "yyyyMMdd-HHmmss"

# Cleans every known OutlookAI registration for one Windows user. Designed
# to be called once per user hive (offline-loaded for non-logged-in users,
# directly under HKEY_USERS\<sid> for logged-on ones). Idempotent.
#
# Why all of these?  VSTO's ClickOnceAddInDeploymentManager throws
# AddInAlreadyInstalledException at Outlook startup if any of the following
# point at an older or differently-located manifest than HKLM does:
#   - HKCU\Software\Microsoft\VSTA\Solutions\<guid>          (ClickOnce subscription)
#   - HKCU\Software\Microsoft\VSTO\SolutionMetadata\<guid>   (SolutionId → addin metadata)
#   - HKCU\Software\Microsoft\VSTO\SolutionMetadata          (URL → SolutionId map values)
#   - HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid> (trust list)
#   - HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\<guid> (Add/Remove Programs)
#   - HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI
#   - HKCU\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI
#   - HKCU\Software\Microsoft\Office\Outlook\AddinsData\OutlookAI
#   - HKCU\Software\Microsoft\Office\16.0\Outlook\AddinsData\OutlookAI
#   - %LOCALAPPDATA%\Apps\2.0\...\outl..vsto_*               (ClickOnce app cache)
# Even with `|vstolocal` in HKLM, Outlook prefers any matching HKCU entry,
# and any of the VSTA/VSTO state above can re-trigger the ClickOnce path.
function Clean-StaleOutlookAIRegistrations {
    param(
        [Parameter(Mandatory=$true)] [string] $UserRegistryRoot,
        [string] $ProfilePath,
        [string] $VstoInstaller
    )

    function _RemoveKey($path) {
        if (Test-Path $path) {
            try { Remove-Item -Path $path -Recurse -Force -ErrorAction Stop; return $true } catch { return $false }
        }
        return $false
    }

    # 1) Add/Remove Programs subscription (Uninstall key). Run its own
    #    UninstallString first when VSTOInstaller is available; then drop
    #    the registry entry whether or not the helper succeeded.
    $uninstallRoot = "$UserRegistryRoot\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    if (Test-Path $uninstallRoot) {
        Get-ChildItem -Path $uninstallRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if (-not $entry) { return }
            if ($entry.DisplayName -notlike "*OutlookAI*") { return }

            if ($VstoInstaller -and $entry.UninstallString -and $entry.UninstallString -match 'file:[^\s"]+OutlookAI\.vsto') {
                $manifest = $Matches[0]
                Write-Host ("  Uninstalling ClickOnce subscription: {0}" -f $manifest) -ForegroundColor Gray
                & $VstoInstaller /Uninstall $manifest /s 2>$null | Out-Null
            }
            try { Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop } catch { }
        }
    }

    # 2) VSTA Solutions subscription (the one VerifySolutionCodebaseIsUnchanged
    #    actually checks). Matched by ProductName/Url containing OutlookAI.
    $vstaRoot = "$UserRegistryRoot\Software\Microsoft\VSTA\Solutions"
    if (Test-Path $vstaRoot) {
        Get-ChildItem -Path $vstaRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and ($entry.ProductName -like "*OutlookAI*" -or $entry.Url -like "*OutlookAI*")) {
                Write-Host ("  Removing VSTA subscription: {0}" -f $entry.Url) -ForegroundColor Gray
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 3) VSTO SolutionMetadata: subkeys (SolutionId → addin name) and
    #    values directly on the parent (URL → SolutionId map).
    $metaRoot = "$UserRegistryRoot\Software\Microsoft\VSTO\SolutionMetadata"
    if (Test-Path $metaRoot) {
        Get-ChildItem -Path $metaRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and ($entry.addInName -like "*OutlookAI*" -or $entry.friendlyName -like "*OutlookAI*")) {
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        $metaKey = Get-Item -Path $metaRoot -ErrorAction SilentlyContinue
        if ($metaKey) {
            foreach ($valName in $metaKey.GetValueNames()) {
                if ($valName -like "*OutlookAI*") {
                    Remove-ItemProperty -Path $metaRoot -Name $valName -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    # 4) VSTO Security Inclusion list (trust). Stale entries point at moved
    #    build paths and confuse the manifest resolver.
    $inclusionRoot = "$UserRegistryRoot\Software\Microsoft\VSTO\Security\Inclusion"
    if (Test-Path $inclusionRoot) {
        Get-ChildItem -Path $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($entry -and $entry.Url -like "*OutlookAI*") {
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 5) HKCU Outlook Addin entries (all known variants).
    foreach ($leaf in @(
        "Software\Microsoft\Office\Outlook\Addins\OutlookAI",
        "Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI",
        "Software\Microsoft\Office\Outlook\AddinsData\OutlookAI",
        "Software\Microsoft\Office\16.0\Outlook\AddinsData\OutlookAI"
    )) {
        _RemoveKey "$UserRegistryRoot\$leaf" | Out-Null
    }

    # 6) Per-user Outlook cache state. These are harmless if left, but Outlook
    #    occasionally surfaces stale ribbon validation results when an addin
    #    is reinstalled with different ribbon XML.
    $loadTimes = "$UserRegistryRoot\Software\Microsoft\Office\16.0\Outlook\AddInLoadTimes"
    if (Test-Path $loadTimes) {
        Remove-ItemProperty -Path $loadTimes -Name "OutlookAI" -Force -ErrorAction SilentlyContinue
    }
    $uiCache = "$UserRegistryRoot\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
    if (Test-Path $uiCache) {
        $uiKey = Get-Item -Path $uiCache -ErrorAction SilentlyContinue
        if ($uiKey) {
            foreach ($valName in $uiKey.GetValueNames()) {
                if ($valName -like "OutlookAI*") {
                    Remove-ItemProperty -Path $uiCache -Name $valName -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    # 7) ClickOnce app cache directories under this profile.
    if ($ProfilePath) {
        $appsRoot = Join-Path $ProfilePath "AppData\Local\Apps\2.0"
        if (Test-Path $appsRoot) {
            Get-ChildItem -Path $appsRoot -Recurse -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "outl*.vsto*" -or $_.Name -like "OutlookAI*" } |
                ForEach-Object {
                    try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
                }
        }
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  OutlookAI v2 Installer (ChatGPT OAuth)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Admin check ---------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

if (!(Test-Path $SourcePath)) {
    Write-Host "ERROR: Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

$vstoFile    = Join-Path $SourcePath "OutlookAI.vsto"
$appFilesDir = Join-Path $SourcePath "Application Files"

if (!(Test-Path $vstoFile)) {
    Write-Host "ERROR: OutlookAI.vsto not found in $SourcePath" -ForegroundColor Red
    exit 1
}
if (!(Test-Path $appFilesDir)) {
    Write-Host "ERROR: 'Application Files' folder not found in $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Source : $SourcePath"     -ForegroundColor Gray
Write-Host "Target : $InstallPath"    -ForegroundColor Gray
Write-Host "Auth   : $ProgramDataPath" -ForegroundColor Gray
Write-Host ""

# --- 0. Clean up any stale OutlookAI registrations ----------------------
# A previous OutlookAI install (Visual Studio publish, ClickOnce setup.exe,
# F5/debug deploy, or v1 installer) may have left behind:
#   - HKCU Outlook add-in entry pointing at a stale build path
#   - HKCU\...\Uninstall ClickOnce subscription pointing at an old manifest
#   - %LOCALAPPDATA%\Apps\2.0 cached ClickOnce manifest
# Even with |vstolocal in our HKLM entry, Outlook prefers HKCU and the
# ClickOnce runtime throws AddInAlreadyInstalledException when the cached
# subscription's manifest URL no longer matches the current install path.
#
# This block cleans up every Windows user profile we can see on this
# machine, not just the admin running the installer. ClickOnce state is
# per-user, so cleaning only the admin profile leaves other users broken.
Write-Host "[0/9] Removing stale OutlookAI registrations for all users..." -ForegroundColor Yellow

# Stop Outlook so the cache files aren't locked.
Get-Process OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  Stopping running Outlook process..." -ForegroundColor Gray
    try { $_ | Stop-Process -Force -ErrorAction Stop } catch { }
}

$vstoInstallerCandidates = @(
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
)
$vstoInstaller = $vstoInstallerCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

# Iterate every loadable user hive. Load offline hives, clean, unload.
$userProfiles = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("Public", "Default", "Default User", "All Users") }

foreach ($profile in $userProfiles) {
    $ntuser = Join-Path $profile.FullName "NTUSER.DAT"
    if (!(Test-Path $ntuser)) { continue }

    $hiveName = "OutlookAICleanup_" + $profile.Name
    $loaded = $false
    try {
        # Try to load the offline hive. Fails if the user is logged in (their
        # hive is already mounted under HKEY_USERS\<sid>); in that case we
        # walk HKEY_USERS directly below.
        & reg.exe load "HKU\$hiveName" "$ntuser" 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) { $loaded = $true }
    } catch { }

    if ($loaded) {
        $userRoot = "Registry::HKEY_USERS\$hiveName"
        try {
            Clean-StaleOutlookAIRegistrations -UserRegistryRoot $userRoot -ProfilePath $profile.FullName -VstoInstaller $vstoInstaller
        } finally {
            & reg.exe unload "HKU\$hiveName" 2>$null | Out-Null
        }
    }
}

# Also clean every currently-loaded HKEY_USERS hive (the admin running this
# script and any other logged-on session).
Get-ChildItem "Registry::HKEY_USERS" -ErrorAction SilentlyContinue | Where-Object {
    $_.PSChildName -match '^S-1-5-21-' -and $_.PSChildName -notmatch '_Classes$'
} | ForEach-Object {
    $sid = $_.PSChildName
    $userRoot = "Registry::HKEY_USERS\$sid"
    $profilePath = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
    Clean-StaleOutlookAIRegistrations -UserRegistryRoot $userRoot -ProfilePath $profilePath -VstoInstaller $vstoInstaller
}

Write-Host "  Done." -ForegroundColor Green

# --- 1. External config backup BEFORE we touch the install dir -----------
Write-Host "[1/9] Backing up any existing v1 config..." -ForegroundColor Yellow
if (!(Test-Path $BackupRoot)) {
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
}
if (Test-Path $ConfigFilePath) {
    $backupTarget = Join-Path $BackupRoot ("config.xml.v1.backup." + $Timestamp)
    Copy-Item -Path $ConfigFilePath -Destination $backupTarget -Force
    Write-Host "  Backed up to $backupTarget" -ForegroundColor Gray
} else {
    Write-Host "  No existing config.xml to back up." -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

# --- 2. Replace install directory ----------------------------------------
Write-Host "[2/9] Preparing install directory..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
}
New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
Write-Host "  Done." -ForegroundColor Green

# --- 3. Copy publish payload ---------------------------------------------
Write-Host "[3/9] Copying files..." -ForegroundColor Yellow
Copy-Item -Path $vstoFile -Destination $InstallPath -Force
Copy-Item -Path $appFilesDir -Destination $InstallPath -Recurse -Force

# Mirror .deploy artifacts as plain DLLs at the root (mirrors v1 layout).
$latestVersionDir = Get-ChildItem -Path (Join-Path $InstallPath "Application Files") -Directory |
    Sort-Object Name -Descending | Select-Object -First 1
if ($latestVersionDir) {
    Get-ChildItem -Path $latestVersionDir.FullName -Filter "*.deploy" | ForEach-Object {
        $newName = $_.Name -replace '\.deploy$', ''
        Copy-Item -Path $_.FullName -Destination (Join-Path $InstallPath $newName) -Force
    }
    $manifestFile = Join-Path $latestVersionDir.FullName "OutlookAI.dll.manifest"
    if (Test-Path $manifestFile) {
        Copy-Item -Path $manifestFile -Destination $InstallPath -Force
    }
}
Write-Host "  Done." -ForegroundColor Green

# --- 4. Unblock files ----------------------------------------------------
Write-Host "[4/9] Unblocking files..." -ForegroundColor Yellow
Get-ChildItem -Path $InstallPath -Recurse | Unblock-File -ErrorAction SilentlyContinue
Write-Host "  Done." -ForegroundColor Green

# --- 5. Write v2 config --------------------------------------------------
Write-Host "[5/9] Writing v2 config.xml..." -ForegroundColor Yellow

# Carry over only the AdminPassword from any v1 file (everything else is
# server-authoritative under v2).
$preservedAdminPassword = "admin"
$latestBackup = Get-ChildItem -Path $BackupRoot -Filter "config.xml.v1.backup.*" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestBackup) {
    try {
        [xml]$oldXml = Get-Content -Path $latestBackup.FullName -Raw
        if ($oldXml.Config -and $oldXml.Config.AdminPassword) {
            $candidate = [string]$oldXml.Config.AdminPassword
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $preservedAdminPassword = $candidate
                Write-Host "  Preserved AdminPassword from previous config." -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  Could not parse previous config; using default AdminPassword." -ForegroundColor Yellow
    }
}

$v2Config = @"
<Config>
  <AdminPassword>$preservedAdminPassword</AdminPassword>
  <CodexAuthPath>C:\ProgramData\OutlookAI\auth.json</CodexAuthPath>
  <Model>gpt-5.5</Model>
  <VoiceModel>gpt-realtime-1.5</VoiceModel>
  <MaxTokens>65536</MaxTokens>
</Config>
"@

Set-Content -Path $ConfigFilePath -Value $v2Config -Encoding UTF8
Write-Host "  Wrote $ConfigFilePath" -ForegroundColor Gray
Write-Host "  Done." -ForegroundColor Green

# --- 6. Provision shared OAuth auth folder + ACL -------------------------
Write-Host "[6/9] Provisioning shared OAuth auth folder..." -ForegroundColor Yellow
if (!(Test-Path $ProgramDataPath)) {
    New-Item -ItemType Directory -Path $ProgramDataPath -Force | Out-Null
}

# Shared per-server OAuth: any signed-in user on this server can read/write
# auth.json. This is the accepted Phase 1 trade-off; if trust changes,
# rotate the credential via Settings -> Sign Out + Sign In.
& icacls.exe $ProgramDataPath /grant "Authenticated Users:(OI)(CI)M" /T | Out-Null
Write-Host "  Granted Authenticated Users: Modify on $ProgramDataPath" -ForegroundColor Gray
Write-Host "  Done." -ForegroundColor Green

# --- 7. Per-user v1 AppData cleanup --------------------------------------
Write-Host "[7/9] Renaming per-user v1 AppData configs..." -ForegroundColor Yellow
$userProfiles = Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("Public", "Default", "Default User", "All Users") }
$renamed = 0
foreach ($profile in $userProfiles) {
    $userConfig = Join-Path $profile.FullName "AppData\Roaming\OutlookAI\config.xml"
    if (Test-Path $userConfig) {
        try {
            [xml]$xml = Get-Content -Path $userConfig -Raw
            if ($xml.Config -and -not $xml.Config.CodexAuthPath) {
                $renamed++
                $renamedTarget = "$userConfig.v1.backup.$Timestamp"
                Move-Item -Path $userConfig -Destination $renamedTarget -Force
                Write-Host "  Renamed $userConfig -> $renamedTarget" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Skipped unreadable per-user config: $userConfig" -ForegroundColor Yellow
        }
    }
}
Write-Host ("  Done ({0} per-user v1 configs renamed)." -f $renamed) -ForegroundColor Green

# --- 8. VSTO trust + Inclusion List + add-in registration ----------------
Write-Host "[8/9] Configuring VSTO trust & registering add-in..." -ForegroundColor Yellow

$trustPath   = "HKLM:\SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
$trustPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
foreach ($path in @($trustPath, $trustPath32)) {
    if (!(Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "MyComputer"    -Value "Enabled"
    Set-ItemProperty -Path $path -Name "LocalIntranet" -Value "Enabled"
}

$inclusionPath   = "HKLM:\SOFTWARE\Microsoft\VSTO\Security\Inclusion"
$inclusionPath32 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO\Security\Inclusion"
foreach ($base in @($inclusionPath, $inclusionPath32)) {
    if (!(Test-Path $base)) { New-Item -Path $base -Force | Out-Null }
    $guid    = [System.Guid]::NewGuid().ToString("B")
    $entry   = Join-Path $base $guid
    New-Item -Path $entry -Force | Out-Null
    Set-ItemProperty -Path $entry -Name "Url"       -Value "file:///C:/Program Files/OutlookAI/OutlookAI.vsto"
    Set-ItemProperty -Path $entry -Name "PublicKey" -Value ""
}

$manifestPath  = "file:///C:/Program Files/OutlookAI/OutlookAI.vsto|vstolocal"
$addinPath     = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI"
$addinPath32   = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI"
foreach ($path in @($addinPath, $addinPath32)) {
    if (Test-Path $path) { Remove-Item -Path $path -Force }
    New-Item -Path $path -Force | Out-Null
    Set-ItemProperty -Path $path -Name "Description"  -Value "AI Writing Assistant for Outlook"
    Set-ItemProperty -Path $path -Name "FriendlyName" -Value "OutlookAI"
    Set-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $path -Name "Manifest"     -Value $manifestPath
}
Write-Host "  Done." -ForegroundColor Green

# --- 9. Default user profile auto-load -----------------------------------
Write-Host "[9/9] Configuring auto-load for new user profiles..." -ForegroundColor Yellow
$defaultUserPath = "C:\Users\Default\NTUSER.DAT"
$tempKey         = "HKU\DefaultUser"
try {
    reg load $tempKey $defaultUserPath 2>$null
    $defaultAddinPath = "Registry::$tempKey\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookAI"
    if (!(Test-Path $defaultAddinPath)) {
        New-Item -Path $defaultAddinPath -Force | Out-Null
    }
    Set-ItemProperty -Path $defaultAddinPath -Name "LoadBehavior" -Value 3 -Type DWord
    reg unload $tempKey 2>$null
    Write-Host "  Configured default user profile." -ForegroundColor Gray
} catch {
    Write-Host "  Note: could not modify default user profile (may need reboot)." -ForegroundColor Gray
}
Write-Host "  Done." -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "ChatGPT OAuth shared on this server:" -ForegroundColor Yellow
Write-Host "  $AuthFilePath" -ForegroundColor Yellow
Write-Host "  Any signed-in user can read/write this file." -ForegroundColor Yellow
Write-Host "  Only deploy to RDS servers where every interactive user is" -ForegroundColor Yellow
Write-Host "  trusted with the OpenAI / ChatGPT account that signs in." -ForegroundColor Yellow
Write-Host "  To rotate: have an admin open Outlook, click Settings, Sign Out," -ForegroundColor Yellow
Write-Host "  then Sign In again. See Deploy\\README.txt." -ForegroundColor Yellow
Write-Host ""
Write-Host "Users will see 'AI Assistant' on the compose-window ribbon." -ForegroundColor White
Write-Host "First action prompts for ChatGPT sign-in via the local browser." -ForegroundColor White
Write-Host ""
Write-Host "If add-in doesn't auto-load for existing users:" -ForegroundColor Yellow
Write-Host "  - Have user run Enable-OutlookAI-User.ps1, OR"             -ForegroundColor Gray
Write-Host "  - Push it via logon script / GPO."                          -ForegroundColor Gray
Write-Host ""
