<#
.SYNOPSIS
    Enables OutlookAI add-in for the current user (clears Outlook's disabled list)

.DESCRIPTION
    Run this if Outlook has disabled the OutlookAI add-in.
    Can be used as a logon script for all users.
#>

# Office version keys to check
$officeVersions = @("16.0", "15.0")

foreach ($version in $officeVersions) {
    $resiliencyPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Resiliency"

    if (Test-Path $resiliencyPath) {
        # Remove disabled items tracking
        Remove-Item "$resiliencyPath\DisabledItems" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\DoNotDisableAddinList" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\NotDisabledAddinList" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\CrashingAddinList" -Force -ErrorAction SilentlyContinue
    }

    # Ensure add-in is enabled at user level
    $addinPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Addins\OutlookAI"
    if (!(Test-Path $addinPath)) {
        New-Item -Path $addinPath -Force | Out-Null
    }
    Set-ItemProperty -Path $addinPath -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue
}

Write-Host "OutlookAI enabled. Please restart Outlook." -ForegroundColor Green
