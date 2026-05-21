<#
.SYNOPSIS
    Builds a Release publish of OutlookAI, generates version.json, packages
    the deploy ZIP, computes the SHA256 sidecar. Output is suitable for
    `gh release create`.

.PARAMETER Tag
    Semver release tag, e.g. v2.1.0. Required.

.PARAMETER OutDir
    Directory where the final .zip and .zip.sha256 should land. Defaults to
    .\out next to the repo root.

.EXAMPLE
    .\Deploy\Make-ReleaseZip.ps1 -Tag v2.1.0
#>
param(
    [Parameter(Mandatory=$true)][string]$Tag,
    [string]$OutDir = "out"
)

$ErrorActionPreference = "Stop"

if ($Tag -notmatch '^v\d+\.\d+\.\d+(-[0-9A-Za-z\.\-]+)?$') {
    throw "Tag '$Tag' is not a valid semver tag (e.g. v2.1.0 or v2.1.0-beta.1)."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $repoRoot

$staging = Join-Path $repoRoot "out\staging-$Tag"
if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null

$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
# Publish only the main project, not the whole solution, so the staging dir
# does not pick up OutlookAI.Tests artifacts (Moq, xunit, Castle.Core,
# Microsoft.CodeCoverage.*, etc.). The csproj's Platform is "AnyCPU"
# (no space) even though the sln uses "Any CPU".
& $msbuild "VSTO2\OutlookAI\OutlookAI.csproj" /target:Publish /p:Configuration=Release /p:Platform="AnyCPU" /p:PublishDir="$staging\" /v:minimal /nologo

# Copy install assets next to OutlookAI.vsto / Application Files\
foreach ($f in @(
    "Install-OutlookAI.ps1",
    "Uninstall-OutlookAI.ps1",
    "Fetch-WebView2Bootstrapper.ps1",
    "MicrosoftEdgeWebView2Setup.exe",
    "README.txt"
)) {
    Copy-Item -LiteralPath (Join-Path "Deploy" $f) -Destination (Join-Path $staging $f) -Force
}

# Convenience copy of the actual DLL at the staging root for hash-verify
Copy-Item -LiteralPath "VSTO2\OutlookAI\bin\Release\OutlookAI.dll" -Destination (Join-Path $staging "OutlookAI.dll") -Force

# version.json
$commit = (git rev-parse --short HEAD).Trim()
$buildDate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$versionJson = [ordered]@{
    tag = $Tag
    commit = $commit
    build_date = $buildDate
    repo = "kirklandsig/OutlookAI"
} | ConvertTo-Json
Set-Content -LiteralPath (Join-Path $staging "version.json") -Encoding UTF8 -Value $versionJson

# Zip
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$zipName = "OutlookAI-$Tag-RDS-Deploy.zip"
$zipPath = Join-Path $OutDir $zipName
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -LiteralPath (Get-ChildItem -LiteralPath $staging | ForEach-Object { $_.FullName }) -DestinationPath $zipPath -CompressionLevel Optimal

# SHA256 sidecar (single line of lowercase hex, no trailing newline beyond what Set-Content adds)
$sha = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath ($zipPath + ".sha256") -Encoding ASCII -Value $sha -NoNewline

Write-Host "Built $zipPath" -ForegroundColor Green
Write-Host "SHA256 $sha" -ForegroundColor Green
