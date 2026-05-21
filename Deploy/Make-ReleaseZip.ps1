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

Push-Location $repoRoot
try {
    $staging = Join-Path $repoRoot "out\staging-$Tag"
    if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
    New-Item -ItemType Directory -Path $staging | Out-Null

    # Locate MSBuild.exe:
    #   1) PATH (works in 'Developer PowerShell for VS' locally, and on CI
    #      after microsoft/setup-msbuild@v2).
    #   2) vswhere fallback for plain PowerShell sessions on dev machines.
    #   3) Last-resort: hard-coded path to the VS18 install on this dev box.
    $msbuild = (Get-Command MSBuild.exe -ErrorAction SilentlyContinue) `
        | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue
    if (-not $msbuild) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
        if (Test-Path -LiteralPath $vswhere) {
            $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild `
                -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        }
    }
    if (-not $msbuild) {
        # Last-resort fallback: dev-machine-specific VS18 install path. Only
        # hit when MSBuild isn't on PATH and vswhere can't find it (e.g.
        # plain PowerShell session on this dev box without a working vswhere).
        $devFallback = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path -LiteralPath $devFallback) { $msbuild = $devFallback }
    }
    if (-not $msbuild) {
        throw "MSBuild.exe not found. Ensure 'Developer PowerShell for VS' is active locally, or rely on 'microsoft/setup-msbuild@v2' in CI."
    }

    # Publish only the main project, not the whole solution, so the staging dir
    # does not pick up OutlookAI.Tests artifacts (Moq, xunit, Castle.Core,
    # Microsoft.CodeCoverage.*, etc.). The csproj's Platform is "AnyCPU"
    # (no space) even though the sln uses "Any CPU".
    & $msbuild "VSTO2\OutlookAI\OutlookAI.csproj" `
        /target:Publish `
        /p:Configuration=Release `
        /p:Platform="AnyCPU" `
        /p:PublishDir="$staging\" `
        /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild publish failed with exit code $LASTEXITCODE."
    }

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
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($commit)) {
        throw "git rev-parse failed (exit=$LASTEXITCODE, output='$commit'). Run from a git working tree."
    }
    $buildDate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $versionJson = [ordered]@{
        tag = $Tag
        commit = $commit
        build_date = $buildDate
        repo = "kirklandsig/OutlookAI"
    } | ConvertTo-Json
    # Write UTF-8 *without* BOM. The in-app updater and any downstream JSON
    # reader should not see a leading 0xEF 0xBB 0xBF that some parsers reject.
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $staging "version.json"), $versionJson, $utf8NoBom)

    # Defensive: fail loudly if something we expect downstream isn't there,
    # rather than shipping a broken ZIP.
    foreach ($required in @("OutlookAI.vsto","Application Files","Install-OutlookAI.ps1","version.json")) {
        if (-not (Test-Path -LiteralPath (Join-Path $staging $required))) {
            throw "Staging is missing required artifact: $required"
        }
    }

    # Zip
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
    $zipName = "OutlookAI-$Tag-RDS-Deploy.zip"
    $zipPath = Join-Path $OutDir $zipName
    if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal

    # SHA256 sidecar (single line of lowercase hex, no trailing newline beyond what Set-Content adds)
    $sha = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -LiteralPath ($zipPath + ".sha256") -Encoding ASCII -Value $sha -NoNewline

    Write-Host "Built $zipPath" -ForegroundColor Green
    Write-Host "SHA256 $sha" -ForegroundColor Green
}
finally {
    Pop-Location
}
