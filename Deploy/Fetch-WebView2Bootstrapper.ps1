#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads the Microsoft Edge WebView2 Evergreen Bootstrapper into the
    Deploy/ folder so Install-OutlookAI.ps1 can install the runtime offline
    on stripped-down server images.

.DESCRIPTION
    The Phase 2 chat surface runs inside a WebView2 control. Modern Windows
    desktops ship the runtime, but some RDS / server images don't. We vendor
    the official 'MicrosoftEdgeWebView2Setup.exe' bootstrapper alongside the
    installer; this script grabs the latest signed copy from Microsoft's
    fwlink redirector and verifies the Authenticode signature before storing
    it.

    Re-run this script whenever you refresh the deployment bundle so admins
    get the current bootstrapper (Microsoft updates the binary roughly
    monthly to track WebView2 runtime versions).

.EXAMPLE
    .\Fetch-WebView2Bootstrapper.ps1
#>

$ErrorActionPreference = "Stop"
$target = Join-Path $PSScriptRoot "MicrosoftEdgeWebView2Setup.exe"
$source = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"

Write-Host "Fetching WebView2 bootstrapper from Microsoft..." -ForegroundColor Cyan
Write-Host "  Source: $source" -ForegroundColor Gray
Write-Host "  Target: $target" -ForegroundColor Gray

# Microsoft serves the bootstrapper unsigned-by-default; modern TLS is fine.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
Invoke-WebRequest -Uri $source -OutFile $target -UseBasicParsing

$info = Get-Item $target
if ($info.Length -lt 100KB) {
    throw "Downloaded file is suspiciously small ($($info.Length) bytes). Aborting."
}
Write-Host ("  Size: {0:N0} bytes" -f $info.Length) -ForegroundColor Gray

$sig = Get-AuthenticodeSignature -FilePath $target
Write-Host "  Signature status: $($sig.Status)" -ForegroundColor Gray
if ($sig.Status -ne 'Valid') {
    throw "Authenticode signature is not Valid (got '$($sig.Status)'). Refusing to vendor an unsigned binary."
}
$subject = $sig.SignerCertificate.Subject
Write-Host "  Signer: $subject" -ForegroundColor Gray
if ($subject -notmatch 'O=Microsoft Corporation') {
    throw "Signer is not Microsoft Corporation ('$subject'). Refusing to vendor."
}

Write-Host "Done. Bootstrapper is staged for Install-OutlookAI.ps1." -ForegroundColor Green
