<#
.SYNOPSIS
    Enables OutlookAI add-in for the current user and prevents Outlook from disabling it.

.DESCRIPTION
    Sets Group Policy registry keys to force-enable OutlookAI and prevent
    Outlook's resiliency feature from ever disabling it.
    Can be used as a logon script for all users.
#>

# Office version keys to check
$officeVersions = @("16.0", "15.0")

foreach ($version in $officeVersions) {
    # Clear any existing resiliency blocks
    $resiliencyPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Resiliency"
    if (Test-Path $resiliencyPath) {
        Remove-Item "$resiliencyPath\DisabledItems" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\DoNotDisableAddinList" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\NotDisabledAddinList" -Force -ErrorAction SilentlyContinue
        Remove-Item "$resiliencyPath\CrashingAddinList" -Force -ErrorAction SilentlyContinue
    }

    # Force add-in to always be enabled via Group Policy (Outlook cannot override this)
    $policyPath = "HKCU:\Software\Policies\Microsoft\Office\$version\Outlook\Resiliency\AddinList"
    if (!(Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
    }
    # Value 1 = Always enabled, Outlook cannot disable it
    Set-ItemProperty -Path $policyPath -Name "OutlookAI" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    # Ensure add-in is enabled at user level
    $addinPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Addins\OutlookAI"
    if (!(Test-Path $addinPath)) {
        New-Item -Path $addinPath -Force | Out-Null
    }
    Set-ItemProperty -Path $addinPath -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue
}

Write-Host "OutlookAI enabled and protected from resiliency disabling. Please restart Outlook." -ForegroundColor Green
