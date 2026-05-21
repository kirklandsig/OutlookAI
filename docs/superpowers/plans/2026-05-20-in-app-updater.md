# In-App Updater Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a Settings UI button that downloads the latest GitHub Release ZIP, verifies its SHA256, and spawns the existing `Install-OutlookAI.ps1` elevated so the admin can update OutlookAI without redeploying manually.

**Architecture:** Five small services under `Services/Updates/` (manifest, version comparator, GitHub client, downloader, installer launcher) + an append-only history log, wired into a new "Updates" section of the existing admin-password-gated `SettingsForm`. A GitHub Actions workflow on `v*` tag push builds the deploy ZIP, computes a SHA256 sidecar, and creates a Release with both assets. The installer is augmented to copy a `version.json` from the staging dir into `C:\Program Files\OutlookAI\` at install time so the running add-in knows its own version.

**Tech Stack:** C# .NET Framework 4.7.2 (VSTO), `System.Net.Http.HttpClient`, `System.IO.Compression.ZipFile`, `Newtonsoft.Json`, xUnit, PowerShell, GitHub Actions on `windows-2025`.

**Spec:** `docs/superpowers/specs/2026-05-20-in-app-updater-design.md`

---

## File Structure

**New files:**

- `.github/workflows/release.yml` — release workflow triggered by `v*` tag push.
- `Deploy/Make-ReleaseZip.ps1` — local helper that does the same thing the CI does (build, package, hash); used for offline rebuilds and as the CI fallback.
- `VSTO2/OutlookAI/Services/Updates/UpdateManifest.cs` — `version.json` DTO + loaders.
- `VSTO2/OutlookAI/Services/Updates/SemVer.cs` — tiny internal semver parser (`major.minor.patch[-prerelease[.N]]`) + `IComparable<SemVer>`.
- `VSTO2/OutlookAI/Services/Updates/UpdateAvailability.cs` — enum.
- `VSTO2/OutlookAI/Services/Updates/VersionComparator.cs` — pure static comparator using `SemVer`.
- `VSTO2/OutlookAI/Services/Updates/ReleaseInfo.cs` — release DTO.
- `VSTO2/OutlookAI/Services/Updates/ReleaseLookupResult.cs` — abstract + `ReleaseFound` / `NoReleasesAvailable` / `RateLimited` / `NetworkError`.
- `VSTO2/OutlookAI/Services/Updates/GitHubReleaseClient.cs` — async wrapper over `HttpClient`.
- `VSTO2/OutlookAI/Services/Updates/DownloadResult.cs` — abstract + `DownloadSuccess` / `HashMismatch` / `DownloadFailed` / `MissingInstallerScript` / `Cancelled`.
- `VSTO2/OutlookAI/Services/Updates/UpdateDownloader.cs` — HTTP + SHA256 verify + unzip.
- `VSTO2/OutlookAI/Services/Updates/LaunchResult.cs` — abstract + `Launched` / `UacDeclined` / `LaunchFailed`.
- `VSTO2/OutlookAI/Services/Updates/UpdateInstaller.cs` — detached elevated `powershell.exe` launch.
- `VSTO2/OutlookAI/Services/Updates/UpdateHistoryLog.cs` — append-only JSON log capped at 50 entries.
- `VSTO2/OutlookAI/Services/Updates/UpdatePaths.cs` — small helper exposing `%LOCALAPPDATA%\OutlookAI\Updates\`, the in-progress sentinel path, and the install-dir `version.json` path. Single source of truth so tests can override the base path.
- Matching tests under `VSTO2/OutlookAI.Tests/Services/Updates/`.

**Modified files:**

- `Deploy/Install-OutlookAI.ps1` — add a step that copies `<SourcePath>\version.json` into `C:\Program Files\OutlookAI\version.json` if present.
- `VSTO2/OutlookAI/SettingsForm.cs` (and `.Designer.cs` if separate; for this project the form is a single hand-written file with no designer split) — add an `Updates` `GroupBox`, fields, click handlers, state machine.
- `VSTO2/OutlookAI/ThisAddIn.cs` — on `Startup`, clear the `.in-progress` sentinel if the install completed or the sentinel is stale.
- `VSTO2/OutlookAI/OutlookAI.csproj` — register all new `.cs` files in the `<Compile>` group.
- `handoff.md` — post-merge update (local, gitignored).

**Verification commands** (run from
`C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-in-app-updater`):

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Baseline before starting: **553 tests passing**. Target at end: **≥ 580** (add roughly 27 new tests across the components below).

---

## Task 1: Release workflow + `version.json` plumbing

Set up the release pipeline before any updater code, so the rest of the work can assume `version.json` exists in installed copies and the GitHub Releases API returns assets in the expected shape.

**Files:**
- Create: `.github/workflows/release.yml`
- Create: `Deploy/Make-ReleaseZip.ps1`
- Modify: `Deploy/Install-OutlookAI.ps1` (one new step that copies `version.json` into install dir if present)

- [ ] **Step 1: Add `version.json` copy step in `Install-OutlookAI.ps1`**

In `Deploy/Install-OutlookAI.ps1`, locate the `Set-Content -Path $ConfigFilePath -Value $v2Config -Encoding UTF8` line (~line 378). Append the following block immediately after the existing `Write-Host "  Wrote $ConfigFilePath" -ForegroundColor Gray` line and before the `Write-Host "  Done." -ForegroundColor Green` line:

```powershell
# v2.1+ release packages ship a version.json alongside Install-OutlookAI.ps1.
# Copy it into the install dir so the in-app updater knows what is installed.
$stagedVersionJson = Join-Path $SourcePath "version.json"
if (Test-Path $stagedVersionJson) {
    $installedVersionJson = Join-Path $InstallPath "version.json"
    Copy-Item -LiteralPath $stagedVersionJson -Destination $installedVersionJson -Force
    Write-Host "  Wrote $installedVersionJson" -ForegroundColor Gray
} else {
    # Backwards-compatible: older deploy ZIPs do not have version.json. The
    # in-app updater shows "Current: (dev build)" in this case and still
    # allows updates.
    Write-Host "  (no version.json in source path; skipping)" -ForegroundColor Gray
}
```

This is idempotent: re-running the installer with a fresh ZIP overwrites the file.

- [ ] **Step 2: Add `Deploy/Make-ReleaseZip.ps1`**

This is the local equivalent of the CI workflow. It is used both for offline rebuilds and as the canonical reference the CI script mirrors. Create with this content:

```powershell
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
& $msbuild "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" /p:PublishDir="$staging\" /v:minimal /nologo

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
```

- [ ] **Step 3: Add `.github/workflows/release.yml`**

```yaml
name: Release

on:
  push:
    tags:
      - "v*"

permissions:
  contents: write

jobs:
  build-and-release:
    runs-on: windows-2025
    steps:
      - name: Checkout
        uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Setup MSBuild
        uses: microsoft/setup-msbuild@v2

      - name: Setup NuGet
        uses: NuGet/setup-nuget@v2

      - name: Restore packages
        shell: pwsh
        run: |
          nuget restore VSTO2\OutlookAI.sln

      - name: Build deploy ZIP
        shell: pwsh
        run: |
          $tag = "${{ github.ref_name }}"
          .\Deploy\Make-ReleaseZip.ps1 -Tag $tag -OutDir out

      - name: Read tag annotation as release body
        id: tagbody
        shell: pwsh
        run: |
          $tag = "${{ github.ref_name }}"
          $body = git tag -l --format='%(contents)' $tag
          if ([string]::IsNullOrWhiteSpace($body)) { $body = "Release $tag" }
          $body | Out-File -FilePath release-body.md -Encoding UTF8

      - name: Create GitHub Release
        env:
          GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        shell: pwsh
        run: |
          $tag = "${{ github.ref_name }}"
          $zip = "out\OutlookAI-$tag-RDS-Deploy.zip"
          $sha = "$zip.sha256"
          gh release create $tag $zip $sha --title $tag --notes-file release-body.md
```

- [ ] **Step 4: Verify the local helper end-to-end**

```powershell
.\Deploy\Make-ReleaseZip.ps1 -Tag v2.1.0-dryrun -OutDir out-dryrun
Get-ChildItem out-dryrun
Get-Content out-dryrun\OutlookAI-v2.1.0-dryrun-RDS-Deploy.zip.sha256
```

Expected: a `.zip` (~5 MB) and a `.zip.sha256` (single line of hex). Open the ZIP with Explorer and confirm it contains `version.json` at the root alongside `OutlookAI.vsto`, `Application Files\`, the install scripts, the WebView2 bootstrapper, and `OutlookAI.dll`.

Clean up the dry-run output:

```powershell
Remove-Item -LiteralPath out-dryrun -Recurse -Force
```

- [ ] **Step 5: Commit**

```powershell
git add Deploy/Install-OutlookAI.ps1 Deploy/Make-ReleaseZip.ps1 .github/workflows/release.yml
git commit -m "build(release): add tag-triggered Release workflow with SHA256 sidecar"
```

---

## Task 2: `UpdateManifest` + `UpdatePaths` (pure parsers, TDD)

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/UpdatePaths.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateManifest.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/UpdateManifestTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj` (register both new compile items)

- [ ] **Step 1: Write `UpdateManifestTests.cs` (failing test)**

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateManifestTests
    {
        [Fact]
        public void LoadFromFile_ValidJson_ReturnsManifest()
        {
            var path = Path.GetTempFileName();
            try
            {
                File.WriteAllText(path,
                    "{\"tag\":\"v2.1.0\",\"commit\":\"abc1234\"," +
                    "\"build_date\":\"2026-06-02T19:14:00Z\"," +
                    "\"repo\":\"kirklandsig/OutlookAI\"}");
                var manifest = UpdateManifest.LoadFromFile(path);
                Assert.Equal("v2.1.0", manifest.Tag);
                Assert.Equal("abc1234", manifest.Commit);
                Assert.Equal("kirklandsig/OutlookAI", manifest.Repo);
                Assert.False(manifest.IsDevBuild);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void LoadFromFile_MissingFile_ReturnsDevSentinel()
        {
            var manifest = UpdateManifest.LoadFromFile(@"C:\definitely\does\not\exist\version.json");
            Assert.True(manifest.IsDevBuild);
            Assert.Equal(UpdateManifest.DevSentinel, manifest.Tag);
        }

        [Fact]
        public void LoadFromFile_MalformedJson_ReturnsDevSentinel()
        {
            var path = Path.GetTempFileName();
            try
            {
                File.WriteAllText(path, "this is not json");
                var manifest = UpdateManifest.LoadFromFile(path);
                Assert.True(manifest.IsDevBuild);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void LoadFromZip_ReadsVersionJsonAtArchiveRoot()
        {
            var zipPath = Path.GetTempFileName();
            File.Delete(zipPath);
            try
            {
                using (var fs = File.Create(zipPath))
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    var entry = archive.CreateEntry("version.json");
                    using (var w = new StreamWriter(entry.Open(), Encoding.UTF8))
                    {
                        w.Write("{\"tag\":\"v2.2.0\",\"commit\":\"def5678\"," +
                                "\"build_date\":\"2026-07-01T00:00:00Z\"," +
                                "\"repo\":\"kirklandsig/OutlookAI\"}");
                    }
                }
                var manifest = UpdateManifest.LoadFromZip(zipPath);
                Assert.Equal("v2.2.0", manifest.Tag);
            }
            finally { if (File.Exists(zipPath)) File.Delete(zipPath); }
        }

        [Fact]
        public void LoadFromZip_NoVersionJson_ReturnsDevSentinel()
        {
            var zipPath = Path.GetTempFileName();
            File.Delete(zipPath);
            try
            {
                using (var fs = File.Create(zipPath))
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    archive.CreateEntry("something-else.txt");
                }
                var manifest = UpdateManifest.LoadFromZip(zipPath);
                Assert.True(manifest.IsDevBuild);
            }
            finally { if (File.Exists(zipPath)) File.Delete(zipPath); }
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246: The type or namespace name 'UpdateManifest' could not be found`.

- [ ] **Step 3: Create `UpdatePaths.cs`**

```csharp
using System;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Single source of truth for the on-disk locations the updater uses.
    /// Tests can override BaseUpdatesDir to point at a temp folder.
    /// </summary>
    public static class UpdatePaths
    {
        /// <summary>
        /// Root for staged downloads, one subdir per release tag.
        /// Defaults to %LOCALAPPDATA%\OutlookAI\Updates.
        /// </summary>
        public static string BaseUpdatesDir { get; set; } =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "Updates");

        /// <summary>
        /// Location of the installed version.json. Read at runtime to tell the
        /// updater what version is live.
        /// </summary>
        public static string InstalledVersionJson { get; set; } =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "OutlookAI",
                "version.json");

        /// <summary>
        /// Sentinel file written when an install is launched and cleared on
        /// next successful Outlook startup.
        /// </summary>
        public static string InProgressSentinel =>
            Path.Combine(BaseUpdatesDir, ".in-progress");

        /// <summary>
        /// Append-only structured log of update activity.
        /// </summary>
        public static string HistoryLog =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "update-history.json");
    }
}
```

- [ ] **Step 4: Create `UpdateManifest.cs`**

```csharp
using System;
using System.IO;
using System.IO.Compression;
using Newtonsoft.Json;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// version.json shipped inside every Release ZIP and copied into
    /// C:\Program Files\OutlookAI\ by the installer.
    /// </summary>
    public sealed class UpdateManifest
    {
        public const string DevSentinel = "(dev build)";

        [JsonProperty("tag")]
        public string Tag { get; set; }

        [JsonProperty("commit")]
        public string Commit { get; set; }

        [JsonProperty("build_date")]
        public DateTimeOffset BuildDate { get; set; }

        [JsonProperty("repo")]
        public string Repo { get; set; }

        [JsonIgnore]
        public bool IsDevBuild => string.IsNullOrEmpty(Tag) || Tag == DevSentinel;

        public static UpdateManifest LoadFromInstallDir()
            => LoadFromFile(UpdatePaths.InstalledVersionJson);

        public static UpdateManifest LoadFromFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return Dev();
                var json = File.ReadAllText(path);
                return Parse(json);
            }
            catch { return Dev(); }
        }

        public static UpdateManifest LoadFromZip(string zipPath)
        {
            try
            {
                using (var archive = ZipFile.OpenRead(zipPath))
                {
                    var entry = archive.GetEntry("version.json");
                    if (entry == null) return Dev();
                    using (var reader = new StreamReader(entry.Open()))
                    {
                        return Parse(reader.ReadToEnd());
                    }
                }
            }
            catch { return Dev(); }
        }

        private static UpdateManifest Parse(string json)
        {
            var parsed = JsonConvert.DeserializeObject<UpdateManifest>(json);
            if (parsed == null || string.IsNullOrWhiteSpace(parsed.Tag)) return Dev();
            return parsed;
        }

        private static UpdateManifest Dev() => new UpdateManifest { Tag = DevSentinel };
    }
}
```

- [ ] **Step 5: Register both files in `OutlookAI.csproj`**

In `VSTO2/OutlookAI/OutlookAI.csproj`, locate the `<ItemGroup>` that contains other `<Compile Include="Services\...\*.cs" />` lines. Add:

```xml
<Compile Include="Services\Updates\UpdatePaths.cs" />
<Compile Include="Services\Updates\UpdateManifest.cs" />
```

(All subsequent Updates files will be added to the same `<ItemGroup>` in their respective tasks.)

- [ ] **Step 6: Build and run the new tests** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~UpdateManifestTests"
```

Expected: 5 tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/UpdatePaths.cs `
        VSTO2/OutlookAI/Services/Updates/UpdateManifest.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/UpdateManifestTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add UpdateManifest + UpdatePaths"
```

---

## Task 3: `SemVer` + `VersionComparator` + `UpdateAvailability` (TDD)

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateAvailability.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/SemVer.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/VersionComparator.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/VersionComparatorTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `VersionComparatorTests.cs` (failing)**

```csharp
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class VersionComparatorTests
    {
        [Theory]
        [InlineData("v2.0.0",  "v2.1.0",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.0.0",  "v2.0.1",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.0.0",  "v3.0.0",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.1.0",  "v2.1.0",  UpdateAvailability.NoUpdate)]
        [InlineData("v2.1.0",  "v2.0.9",  UpdateAvailability.OlderThanInstalled)]
        [InlineData("2.1.0",   "v2.1.0",  UpdateAvailability.NoUpdate)]                // missing v prefix
        [InlineData("v2.1.0",  "2.1.0",   UpdateAvailability.NoUpdate)]                // missing v prefix
        [InlineData("v2.1.0-beta.1", "v2.1.0", UpdateAvailability.NewerAvailable)]     // beta < release
        [InlineData("v2.1.0", "v2.1.0-beta.1", UpdateAvailability.OlderThanInstalled)] // release > beta
        [InlineData("v2.1.0-beta.1", "v2.1.0-beta.2", UpdateAvailability.NewerAvailable)]
        public void Compare_KnownPairs_ReturnsExpected(string installed, string latest, UpdateAvailability expected)
        {
            Assert.Equal(expected, VersionComparator.Compare(installed, latest));
        }

        [Theory]
        [InlineData("(dev build)", "v2.1.0")]
        [InlineData("v2.1.0",      "garbage")]
        [InlineData(null,          "v2.1.0")]
        [InlineData("v2.1.0",      "")]
        public void Compare_UnparseableInputs_ReturnsNotComparable(string installed, string latest)
        {
            Assert.Equal(UpdateAvailability.NotComparable, VersionComparator.Compare(installed, latest));
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246: ... 'VersionComparator' could not be found ...`.

- [ ] **Step 3: Create `UpdateAvailability.cs`**

```csharp
namespace OutlookAI.Services.Updates
{
    public enum UpdateAvailability
    {
        NoUpdate,
        NewerAvailable,
        OlderThanInstalled,
        NotComparable,
        NoReleases,
    }
}
```

- [ ] **Step 4: Create `SemVer.cs`** (tiny internal parser)

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Minimal semver parser for tags shaped like
    /// `[v]MAJOR.MINOR.PATCH[-PRERELEASE]`.
    /// Prerelease suffix sorts LOWER than the same numeric tuple without it,
    /// per SemVer 2.0.0. Prerelease itself is compared dot-separated:
    /// numeric identifiers numerically, others lexicographically.
    /// </summary>
    internal sealed class SemVer : IComparable<SemVer>
    {
        public int Major { get; }
        public int Minor { get; }
        public int Patch { get; }
        public string PreRelease { get; }

        private SemVer(int major, int minor, int patch, string pre)
        {
            Major = major;
            Minor = minor;
            Patch = patch;
            PreRelease = pre ?? string.Empty;
        }

        public static bool TryParse(string raw, out SemVer version)
        {
            version = null;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            var s = raw.Trim();
            if (s.StartsWith("v", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1);

            string pre = "";
            var dash = s.IndexOf('-');
            if (dash >= 0)
            {
                pre = s.Substring(dash + 1);
                s = s.Substring(0, dash);
            }

            var parts = s.Split('.');
            if (parts.Length != 3) return false;

            if (!int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var maj)) return false;
            if (!int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var min)) return false;
            if (!int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var pat)) return false;

            version = new SemVer(maj, min, pat, pre);
            return true;
        }

        public int CompareTo(SemVer other)
        {
            if (other == null) return 1;
            var c = Major.CompareTo(other.Major); if (c != 0) return c;
            c = Minor.CompareTo(other.Minor);     if (c != 0) return c;
            c = Patch.CompareTo(other.Patch);     if (c != 0) return c;

            // Per SemVer 2.0.0, a version with prerelease is LOWER than the
            // same version without one.
            if (PreRelease.Length == 0 && other.PreRelease.Length == 0) return 0;
            if (PreRelease.Length == 0) return 1;
            if (other.PreRelease.Length == 0) return -1;

            return ComparePreRelease(PreRelease, other.PreRelease);
        }

        private static int ComparePreRelease(string a, string b)
        {
            var ap = a.Split('.');
            var bp = b.Split('.');
            var n = Math.Min(ap.Length, bp.Length);
            for (var i = 0; i < n; i++)
            {
                var ai = ap[i];
                var bi = bp[i];
                var aNum = int.TryParse(ai, NumberStyles.None, CultureInfo.InvariantCulture, out var an);
                var bNum = int.TryParse(bi, NumberStyles.None, CultureInfo.InvariantCulture, out var bn);
                if (aNum && bNum) { var c = an.CompareTo(bn); if (c != 0) return c; continue; }
                if (aNum) return -1;   // numeric identifiers sort lower than alphanumeric
                if (bNum) return 1;
                var sc = string.CompareOrdinal(ai, bi); if (sc != 0) return sc;
            }
            return ap.Length.CompareTo(bp.Length);
        }
    }
}
```

- [ ] **Step 5: Create `VersionComparator.cs`**

```csharp
namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Pure static comparator used by the updater to decide whether the
    /// GitHub-reported release tag is newer than what is installed.
    /// </summary>
    public static class VersionComparator
    {
        public static UpdateAvailability Compare(string installedTag, string latestTag)
        {
            if (!SemVer.TryParse(installedTag, out var a)) return UpdateAvailability.NotComparable;
            if (!SemVer.TryParse(latestTag,    out var b)) return UpdateAvailability.NotComparable;

            var c = a.CompareTo(b);
            if (c < 0) return UpdateAvailability.NewerAvailable;
            if (c > 0) return UpdateAvailability.OlderThanInstalled;
            return UpdateAvailability.NoUpdate;
        }
    }
}
```

- [ ] **Step 6: Register the three new files in the csproj**

In `VSTO2/OutlookAI/OutlookAI.csproj` add to the Updates `<ItemGroup>`:

```xml
<Compile Include="Services\Updates\UpdateAvailability.cs" />
<Compile Include="Services\Updates\SemVer.cs" />
<Compile Include="Services\Updates\VersionComparator.cs" />
```

- [ ] **Step 7: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~VersionComparatorTests"
```

Expected: all `VersionComparatorTests` pass (10 + 4 = 14 parameterized cases).

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/UpdateAvailability.cs `
        VSTO2/OutlookAI/Services/Updates/SemVer.cs `
        VSTO2/OutlookAI/Services/Updates/VersionComparator.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/VersionComparatorTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add SemVer parser and VersionComparator"
```

---

## Task 4: `ReleaseInfo` + `ReleaseLookupResult` + `GitHubReleaseClient` (TDD)

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/ReleaseInfo.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/ReleaseLookupResult.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/GitHubReleaseClient.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/GitHubReleaseClientTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `GitHubReleaseClientTests.cs` (failing)**

```csharp
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Updates;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class GitHubReleaseClientTests
    {
        private const string SampleJson = @"{
            ""tag_name"": ""v2.1.0"",
            ""html_url"": ""https://github.com/kirklandsig/OutlookAI/releases/tag/v2.1.0"",
            ""published_at"": ""2026-06-02T19:14:00Z"",
            ""body"": ""Release notes here"",
            ""assets"": [
                { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip"",
                  ""browser_download_url"": ""https://github.com/.../OutlookAI-v2.1.0-RDS-Deploy.zip"" },
                { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip.sha256"",
                  ""browser_download_url"": ""https://github.com/.../OutlookAI-v2.1.0-RDS-Deploy.zip.sha256"" }
            ]
        }";

        private static GitHubReleaseClient NewClient(FakeHttpMessageHandler handler)
        {
            var http = new HttpClient(handler);
            return new GitHubReleaseClient(http, "kirklandsig/OutlookAI", "OutlookAI-Updater/test");
        }

        [Fact]
        public async Task GetLatestStableAsync_HappyPath_ReturnsReleaseFound()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, SampleJson);

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var found = Assert.IsType<ReleaseFound>(result);
            Assert.Equal("v2.1.0", found.Info.Tag);
            Assert.EndsWith("OutlookAI-v2.1.0-RDS-Deploy.zip", found.Info.ZipUrl);
            Assert.EndsWith(".sha256", found.Info.ShaUrl);
            Assert.Equal("OutlookAI-v2.1.0-RDS-Deploy.zip", found.Info.ZipAssetName);
        }

        [Fact]
        public async Task GetLatestStableAsync_404_ReturnsNoReleasesAvailable()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.NotFound, "{\"message\":\"Not Found\"}");

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            Assert.IsType<NoReleasesAvailable>(result);
        }

        [Fact]
        public async Task GetLatestStableAsync_5xx_ReturnsNetworkError()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.InternalServerError, "{}");

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            Assert.IsType<NetworkError>(result);
        }

        [Fact]
        public async Task GetLatestStableAsync_MissingShaAsset_ShaUrlIsNull()
        {
            var noShaJson = @"{
                ""tag_name"": ""v2.1.0"",
                ""html_url"": ""https://github.com/.../releases/tag/v2.1.0"",
                ""published_at"": ""2026-06-02T19:14:00Z"",
                ""body"": """",
                ""assets"": [
                    { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip"",
                      ""browser_download_url"": ""https://github.com/.../zip"" }
                ]
            }";
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, noShaJson);

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var found = Assert.IsType<ReleaseFound>(result);
            Assert.NotNull(found.Info.ZipUrl);
            Assert.Null(found.Info.ShaUrl);
        }

        [Fact]
        public async Task GetLatestStableAsync_SendsUserAgentAndAccept()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, SampleJson);

            await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var req = handler.Requests[0];
            Assert.Equal("OutlookAI-Updater", req.Headers.UserAgent.ToString().Split('/')[0]);
            Assert.Contains("application/vnd.github+json",
                req.Headers.Accept.ToString());
            Assert.Equal("https://api.github.com/repos/kirklandsig/OutlookAI/releases/latest",
                req.RequestUri.ToString());
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246` for `GitHubReleaseClient`, `ReleaseFound`, etc.

- [ ] **Step 3: Create `ReleaseInfo.cs`**

```csharp
using System;

namespace OutlookAI.Services.Updates
{
    public sealed class ReleaseInfo
    {
        public string Tag { get; set; }
        public string ReleasePageUrl { get; set; }
        public DateTimeOffset PublishedAt { get; set; }
        public string Body { get; set; }
        public string ZipAssetName { get; set; }
        public string ZipUrl { get; set; }
        public string ShaUrl { get; set; }
    }
}
```

- [ ] **Step 4: Create `ReleaseLookupResult.cs`**

```csharp
using System;

namespace OutlookAI.Services.Updates
{
    public abstract class ReleaseLookupResult { }

    public sealed class ReleaseFound : ReleaseLookupResult
    {
        public ReleaseInfo Info { get; set; }
    }

    public sealed class NoReleasesAvailable : ReleaseLookupResult { }

    public sealed class RateLimited : ReleaseLookupResult
    {
        public DateTimeOffset ResetAt { get; set; }
    }

    public sealed class NetworkError : ReleaseLookupResult
    {
        public string Detail { get; set; }
    }
}
```

- [ ] **Step 5: Create `GitHubReleaseClient.cs`**

```csharp
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Thin async wrapper over HttpClient. Calls api.github.com and parses
    /// the latest-release JSON shape into a ReleaseInfo. Honors the system
    /// proxy via the default HttpClient configuration on .NET Framework.
    /// </summary>
    public sealed class GitHubReleaseClient
    {
        private readonly HttpClient _http;
        private readonly string _repo;
        private readonly string _userAgent;

        public GitHubReleaseClient(HttpClient http, string repo, string userAgent)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
            _userAgent = string.IsNullOrWhiteSpace(userAgent) ? "OutlookAI-Updater/dev" : userAgent;
        }

        public async Task<ReleaseLookupResult> GetLatestStableAsync(CancellationToken ct)
        {
            var url = "https://api.github.com/repos/" + _repo + "/releases/latest";
            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.UserAgent.ParseAdd(_userAgent);
            req.Headers.Accept.ParseAdd("application/vnd.github+json");

            HttpResponseMessage resp;
            try
            {
                resp = await _http.SendAsync(req, ct).ConfigureAwait(false);
            }
            catch (HttpRequestException ex)
            {
                return new NetworkError { Detail = ex.Message };
            }
            catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
            {
                return new NetworkError { Detail = "Request timed out: " + ex.Message };
            }

            if (resp.StatusCode == HttpStatusCode.NotFound)
            {
                return new NoReleasesAvailable();
            }

            if (resp.StatusCode == HttpStatusCode.Forbidden &&
                resp.Headers.TryGetValues("X-RateLimit-Remaining", out var rem) &&
                rem.FirstOrDefault() == "0")
            {
                var reset = DateTimeOffset.UtcNow.AddHours(1);
                if (resp.Headers.TryGetValues("X-RateLimit-Reset", out var resetVals) &&
                    long.TryParse(resetVals.FirstOrDefault(), out var unixReset))
                {
                    reset = DateTimeOffset.FromUnixTimeSeconds(unixReset);
                }
                return new RateLimited { ResetAt = reset };
            }

            if (!resp.IsSuccessStatusCode)
            {
                return new NetworkError { Detail = "HTTP " + (int)resp.StatusCode };
            }

            var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            try
            {
                return new ReleaseFound { Info = ParseRelease(body) };
            }
            catch (Exception ex)
            {
                return new NetworkError { Detail = "Malformed release JSON: " + ex.Message };
            }
        }

        internal static ReleaseInfo ParseRelease(string json)
        {
            var o = JObject.Parse(json);
            var info = new ReleaseInfo
            {
                Tag = (string)o["tag_name"],
                ReleasePageUrl = (string)o["html_url"],
                Body = (string)o["body"] ?? string.Empty,
                PublishedAt = ParseDate((string)o["published_at"]),
            };

            var assets = o["assets"] as JArray;
            if (assets != null)
            {
                foreach (var a in assets.OfType<JObject>())
                {
                    var name = (string)a["name"] ?? string.Empty;
                    var url = (string)a["browser_download_url"];
                    if (string.IsNullOrEmpty(url)) continue;

                    if (name.EndsWith(".zip.sha256", StringComparison.OrdinalIgnoreCase))
                    {
                        info.ShaUrl = url;
                    }
                    else if (name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        info.ZipAssetName = name;
                        info.ZipUrl = url;
                    }
                }
            }
            return info;
        }

        private static DateTimeOffset ParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return DateTimeOffset.MinValue;
            if (DateTimeOffset.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal,
                out var dt)) return dt;
            return DateTimeOffset.MinValue;
        }
    }
}
```

- [ ] **Step 6: Register the three new files in the csproj**

```xml
<Compile Include="Services\Updates\ReleaseInfo.cs" />
<Compile Include="Services\Updates\ReleaseLookupResult.cs" />
<Compile Include="Services\Updates\GitHubReleaseClient.cs" />
```

- [ ] **Step 7: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~GitHubReleaseClientTests"
```

Expected: 5 tests pass.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/ReleaseInfo.cs `
        VSTO2/OutlookAI/Services/Updates/ReleaseLookupResult.cs `
        VSTO2/OutlookAI/Services/Updates/GitHubReleaseClient.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/GitHubReleaseClientTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add GitHubReleaseClient + ReleaseInfo DTOs"
```

---

## Task 5: `DownloadResult` + `UpdateDownloader` (TDD)

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/DownloadResult.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateDownloader.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/UpdateDownloaderTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `UpdateDownloaderTests.cs` (failing)**

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Updates;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateDownloaderTests : IDisposable
    {
        private readonly string _tempBase;
        private readonly string _originalBaseDir;

        public UpdateDownloaderTests()
        {
            _tempBase = Path.Combine(Path.GetTempPath(), "updater-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(_tempBase);
            _originalBaseDir = UpdatePaths.BaseUpdatesDir;
            UpdatePaths.BaseUpdatesDir = Path.Combine(_tempBase, "Updates");
        }

        public void Dispose()
        {
            UpdatePaths.BaseUpdatesDir = _originalBaseDir;
            try { Directory.Delete(_tempBase, true); } catch { }
        }

        private static byte[] BuildZip(params (string name, string content)[] entries)
        {
            using (var ms = new MemoryStream())
            {
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
                {
                    foreach (var (name, content) in entries)
                    {
                        var e = archive.CreateEntry(name);
                        using (var w = new StreamWriter(e.Open(), Encoding.UTF8)) w.Write(content);
                    }
                }
                return ms.ToArray();
            }
        }

        private static string Sha256(byte[] data)
        {
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(data);
                var sb = new StringBuilder();
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        private static ReleaseInfo Info(string tag = "v2.1.0") => new ReleaseInfo
        {
            Tag = tag,
            ZipAssetName = "OutlookAI-" + tag + "-RDS-Deploy.zip",
            ZipUrl = "https://github.test/zip",
            ShaUrl = "https://github.test/sha",
        };

        [Fact]
        public async Task DownloadAsync_HashMatches_ReturnsSuccessWithInstallerPath()
        {
            var zip = BuildZip(
                ("Install-OutlookAI.ps1", "Write-Host hi"),
                ("version.json", "{\"tag\":\"v2.1.0\"}"));
            var sha = Sha256(zip);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, sha);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            var ok = Assert.IsType<DownloadSuccess>(result);
            Assert.True(File.Exists(ok.InstallerScriptPath));
            Assert.EndsWith(@"\v2.1.0\extracted\Install-OutlookAI.ps1", ok.InstallerScriptPath);
            Assert.Equal(sha, ok.ExpectedSha256);
        }

        [Fact]
        public async Task DownloadAsync_HashMismatch_DeletesStagingAndReturnsHashMismatch()
        {
            var zip = BuildZip(("Install-OutlookAI.ps1", "Write-Host hi"));
            var wrong = new string('0', 64);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, wrong);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<HashMismatch>(result);
            var stagingDir = Path.Combine(UpdatePaths.BaseUpdatesDir, "v2.1.0");
            Assert.False(Directory.Exists(stagingDir));
        }

        [Fact]
        public async Task DownloadAsync_ZipMissingInstaller_ReturnsMissingInstallerScript()
        {
            var zip = BuildZip(("readme.txt", "no installer here"));
            var sha = Sha256(zip);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, sha);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<MissingInstallerScript>(result);
        }

        [Fact]
        public async Task DownloadAsync_HttpError_ReturnsDownloadFailed()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueText(HttpStatusCode.InternalServerError, "boom");

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<DownloadFailed>(result);
        }

        [Fact]
        public async Task DownloadAsync_NullShaUrl_ReturnsDownloadFailed()
        {
            var handler = new FakeHttpMessageHandler();
            var downloader = new UpdateDownloader(new HttpClient(handler));
            var info = Info();
            info.ShaUrl = null;

            var result = await downloader.DownloadAsync(info, null, CancellationToken.None);

            var failed = Assert.IsType<DownloadFailed>(result);
            Assert.Contains("sha256", failed.Detail, StringComparison.OrdinalIgnoreCase);
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246: ... 'UpdateDownloader' ... 'DownloadSuccess' ...`.

- [ ] **Step 3: Create `DownloadResult.cs`**

```csharp
namespace OutlookAI.Services.Updates
{
    public abstract class DownloadResult { }

    public sealed class DownloadSuccess : DownloadResult
    {
        public string StagingDir { get; set; }
        public string ExtractedDir { get; set; }
        public string InstallerScriptPath { get; set; }
        public string ExpectedSha256 { get; set; }
    }

    public sealed class HashMismatch : DownloadResult
    {
        public string Expected { get; set; }
        public string Actual { get; set; }
    }

    public sealed class DownloadFailed : DownloadResult
    {
        public string Detail { get; set; }
    }

    public sealed class MissingInstallerScript : DownloadResult { }

    public sealed class Cancelled : DownloadResult { }
}
```

- [ ] **Step 4: Create `UpdateDownloader.cs`**

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Downloads a release ZIP, verifies the SHA256 against a sidecar file,
    /// extracts the archive into the per-tag staging dir, and returns either
    /// a ready-to-launch DownloadSuccess or a typed failure.
    /// </summary>
    public sealed class UpdateDownloader
    {
        private readonly HttpClient _http;

        public UpdateDownloader(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<DownloadResult> DownloadAsync(
            ReleaseInfo info,
            IProgress<int> progress,
            CancellationToken ct)
        {
            if (info == null) throw new ArgumentNullException(nameof(info));
            if (string.IsNullOrWhiteSpace(info.ShaUrl))
            {
                return new DownloadFailed { Detail = "Release is missing the .sha256 sidecar; refusing to download." };
            }

            var stagingDir = Path.Combine(UpdatePaths.BaseUpdatesDir, info.Tag);
            try { if (Directory.Exists(stagingDir)) Directory.Delete(stagingDir, recursive: true); } catch { }
            Directory.CreateDirectory(stagingDir);

            var zipPath = Path.Combine(stagingDir, info.ZipAssetName ?? "release.zip");
            var shaPath = zipPath + ".sha256";

            try
            {
                await DownloadToFileAsync(info.ZipUrl, zipPath, progress, ct).ConfigureAwait(false);
                await DownloadToFileAsync(info.ShaUrl, shaPath, null, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new Cancelled();
            }
            catch (Exception ex)
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new DownloadFailed { Detail = ex.Message };
            }

            var expected = (File.ReadAllText(shaPath) ?? string.Empty).Trim().ToLowerInvariant();
            var actual = ComputeSha256(zipPath);
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new HashMismatch { Expected = expected, Actual = actual };
            }

            var extracted = Path.Combine(stagingDir, "extracted");
            try
            {
                if (Directory.Exists(extracted)) Directory.Delete(extracted, true);
                ZipFile.ExtractToDirectory(zipPath, extracted);
            }
            catch (Exception ex)
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new DownloadFailed { Detail = "Extract failed: " + ex.Message };
            }

            var installer = Path.Combine(extracted, "Install-OutlookAI.ps1");
            if (!File.Exists(installer))
            {
                return new MissingInstallerScript();
            }

            return new DownloadSuccess
            {
                StagingDir = stagingDir,
                ExtractedDir = extracted,
                InstallerScriptPath = installer,
                ExpectedSha256 = expected,
            };
        }

        private async Task DownloadToFileAsync(string url, string destPath, IProgress<int> progress, CancellationToken ct)
        {
            using (var resp = await _http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
            {
                if (!resp.IsSuccessStatusCode)
                {
                    throw new HttpRequestException("HTTP " + (int)resp.StatusCode + " from " + url);
                }

                var total = resp.Content.Headers.ContentLength ?? -1L;
                using (var src = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var dst = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024, useAsync: true))
                {
                    var buffer = new byte[64 * 1024];
                    long written = 0;
                    int n;
                    while ((n = await src.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(false)) > 0)
                    {
                        await dst.WriteAsync(buffer, 0, n, ct).ConfigureAwait(false);
                        written += n;
                        if (progress != null && total > 0)
                        {
                            progress.Report((int)(written * 100 / total));
                        }
                    }
                }
            }
        }

        private static string ComputeSha256(string path)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(path))
            {
                var hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }
    }
}
```

- [ ] **Step 5: Register the two new files in the csproj**

```xml
<Compile Include="Services\Updates\DownloadResult.cs" />
<Compile Include="Services\Updates\UpdateDownloader.cs" />
```

If `System.IO.Compression.FileSystem` is not already referenced, add it to the `<ItemGroup>` containing other `<Reference Include="System.*">` entries:

```xml
<Reference Include="System.IO.Compression" />
<Reference Include="System.IO.Compression.FileSystem" />
```

- [ ] **Step 6: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~UpdateDownloaderTests"
```

Expected: 5 tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/DownloadResult.cs `
        VSTO2/OutlookAI/Services/Updates/UpdateDownloader.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/UpdateDownloaderTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add UpdateDownloader with SHA256 verify"
```

---

## Task 6: `LaunchResult` + `UpdateInstaller` (source-level test)

Cannot unit-test UAC, but we can pin the elevated launch invocation via a source-level assertion (same pattern as `ExportBridgeTests.OpenWithDefaultApp_SetsLocalWorkingDirectory_ToAvoidUncFailure`).

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/LaunchResult.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateInstaller.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/UpdateInstallerSourceTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `UpdateInstallerSourceTests.cs` (failing)**

```csharp
using System;
using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateInstallerSourceTests
    {
        private static string FindSourceFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate)) return candidate;
                current = current.Parent;
            }
            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }

        [Fact]
        public void LaunchElevatedInstall_UsesRunasVerb_WithNoExitAndExecutionPolicyBypass()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Updates", "UpdateInstaller.cs"));
            var methodStart = source.IndexOf("public LaunchResult LaunchElevatedInstall", StringComparison.Ordinal);
            Assert.True(methodStart >= 0, "LaunchElevatedInstall should exist.");
            var method = source.Substring(methodStart);

            Assert.Contains("Verb = \"runas\"", method);
            Assert.Contains("UseShellExecute = true", method);
            Assert.Contains("-NoExit", method);
            Assert.Contains("-NoProfile", method);
            Assert.Contains("-ExecutionPolicy Bypass", method);
            Assert.Contains("-File", method);
            Assert.Contains("-SourcePath", method);
            Assert.Contains("WorkingDirectory = Path.GetTempPath()", method);
        }

        [Fact]
        public void LaunchElevatedInstall_HandlesUacDeclinedExitCode1223()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Updates", "UpdateInstaller.cs"));
            Assert.Contains("NativeErrorCode == 1223", source);
            Assert.Contains("UacDeclined", source);
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: csproj does not yet have the file, or the file does not yet exist. Either failure is acceptable as RED.

- [ ] **Step 3: Create `LaunchResult.cs`**

```csharp
namespace OutlookAI.Services.Updates
{
    public abstract class LaunchResult { }

    public sealed class Launched : LaunchResult
    {
        public int Pid { get; set; }
    }

    public sealed class UacDeclined : LaunchResult { }

    public sealed class LaunchFailed : LaunchResult
    {
        public string Detail { get; set; }
    }
}
```

- [ ] **Step 4: Create `UpdateInstaller.cs`**

```csharp
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Launches the extracted Install-OutlookAI.ps1 elevated via UAC. The
    /// process is detached: this call returns once the user accepts UAC, and
    /// the elevated PowerShell window will outlive Outlook getting killed by
    /// the installer.
    /// </summary>
    public sealed class UpdateInstaller
    {
        public LaunchResult LaunchElevatedInstall(DownloadSuccess update)
        {
            if (update == null) throw new ArgumentNullException(nameof(update));
            if (string.IsNullOrWhiteSpace(update.InstallerScriptPath)) throw new ArgumentException("InstallerScriptPath required.");
            if (string.IsNullOrWhiteSpace(update.ExtractedDir)) throw new ArgumentException("ExtractedDir required.");

            var psi = new ProcessStartInfo("powershell.exe")
            {
                UseShellExecute = true,
                Verb = "runas",
                // Force a local working directory so the elevated process does
                // not inherit a UNC CWD (Folder-Redirected Documents).
                WorkingDirectory = Path.GetTempPath(),
                Arguments = string.Format(
                    "-NoExit -NoProfile -ExecutionPolicy Bypass -File \"{0}\" -SourcePath \"{1}\"",
                    update.InstallerScriptPath,
                    update.ExtractedDir),
            };

            try
            {
                var p = Process.Start(psi);
                return new Launched { Pid = p?.Id ?? 0 };
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                // ERROR_CANCELLED — user clicked No on the UAC prompt.
                return new UacDeclined();
            }
            catch (Exception ex)
            {
                return new LaunchFailed { Detail = ex.Message };
            }
        }
    }
}
```

- [ ] **Step 5: Register the two new files**

```xml
<Compile Include="Services\Updates\LaunchResult.cs" />
<Compile Include="Services\Updates\UpdateInstaller.cs" />
```

- [ ] **Step 6: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~UpdateInstallerSourceTests"
```

Expected: 2 tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/LaunchResult.cs `
        VSTO2/OutlookAI/Services/Updates/UpdateInstaller.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/UpdateInstallerSourceTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add UpdateInstaller with detached elevated launch"
```

---

## Task 7: `UpdateHistoryLog` (TDD)

Append-only JSON log capped at 50 entries.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateHistoryLog.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/UpdateHistoryLogTests.cs`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `UpdateHistoryLogTests.cs` (failing)**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateHistoryLogTests : IDisposable
    {
        private readonly string _path;

        public UpdateHistoryLogTests()
        {
            var dir = Path.Combine(Path.GetTempPath(), "updater-history-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            _path = Path.Combine(dir, "update-history.json");
        }

        public void Dispose()
        {
            try { var d = Path.GetDirectoryName(_path); if (d != null) Directory.Delete(d, true); } catch { }
        }

        [Fact]
        public void Append_WritesEntryAndReadAllReturnsIt()
        {
            var log = new UpdateHistoryLog(_path);
            log.Append("check", "newer_available", "v2.1.0", "");

            var entries = log.ReadAll();
            Assert.Single(entries);
            Assert.Equal("check", entries[0].Action);
            Assert.Equal("newer_available", entries[0].Result);
            Assert.Equal("v2.1.0", entries[0].Tag);
        }

        [Fact]
        public void Append_KeepsOnlyLast50Entries()
        {
            var log = new UpdateHistoryLog(_path);
            for (var i = 0; i < 60; i++) log.Append("check", "noop", "v" + i, "");

            var entries = log.ReadAll();
            Assert.Equal(50, entries.Count);
            Assert.Equal("v10", entries[0].Tag);  // oldest 10 dropped
            Assert.Equal("v59", entries[49].Tag);
        }

        [Fact]
        public void ReadAll_MissingFile_ReturnsEmpty()
        {
            var log = new UpdateHistoryLog(_path);
            Assert.Empty(log.ReadAll());
        }

        [Fact]
        public void ReadAll_MalformedFile_ReturnsEmptyAndDoesNotThrow()
        {
            File.WriteAllText(_path, "not json");
            var log = new UpdateHistoryLog(_path);
            Assert.Empty(log.ReadAll());
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246: ... 'UpdateHistoryLog' ...`.

- [ ] **Step 3: Create `UpdateHistoryLog.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Append-only JSON log of update activity. Caps at 50 entries; oldest
    /// dropped on overflow. Tolerant of missing or malformed files (returns
    /// empty list).
    /// </summary>
    public sealed class UpdateHistoryLog
    {
        public sealed class Entry
        {
            [JsonProperty("ts")]      public DateTimeOffset Ts { get; set; }
            [JsonProperty("action")]  public string Action { get; set; }
            [JsonProperty("result")]  public string Result { get; set; }
            [JsonProperty("tag")]     public string Tag { get; set; }
            [JsonProperty("details")] public string Details { get; set; }
        }

        public const int MaxEntries = 50;

        private readonly string _path;
        private readonly object _gate = new object();

        public UpdateHistoryLog() : this(UpdatePaths.HistoryLog) { }

        public UpdateHistoryLog(string path)
        {
            _path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public void Append(string action, string result, string tag, string details)
        {
            lock (_gate)
            {
                var entries = ReadAll();
                entries.Add(new Entry
                {
                    Ts = DateTimeOffset.UtcNow,
                    Action = action ?? "",
                    Result = result ?? "",
                    Tag = tag ?? "",
                    Details = details ?? "",
                });
                while (entries.Count > MaxEntries) entries.RemoveAt(0);

                var dir = Path.GetDirectoryName(_path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var json = JsonConvert.SerializeObject(entries, Formatting.Indented);
                File.WriteAllText(_path, json);
            }
        }

        public List<Entry> ReadAll()
        {
            lock (_gate)
            {
                try
                {
                    if (!File.Exists(_path)) return new List<Entry>();
                    var json = File.ReadAllText(_path);
                    var parsed = JsonConvert.DeserializeObject<List<Entry>>(json);
                    return parsed ?? new List<Entry>();
                }
                catch { return new List<Entry>(); }
            }
        }
    }
}
```

- [ ] **Step 4: Register**

```xml
<Compile Include="Services\Updates\UpdateHistoryLog.cs" />
```

- [ ] **Step 5: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~UpdateHistoryLogTests"
```

Expected: 4 tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/UpdateHistoryLog.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/UpdateHistoryLogTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): add UpdateHistoryLog append-only log"
```

---

## Task 8: Settings UI — Updates section

Add the Updates `GroupBox` to `SettingsForm`. State machine: `Idle → Checking → {UpdateAvailable | UpToDate | NoReleases | CheckFailed} → Downloading → Launching`. Two buttons (Check Now, Install Update). Source-level tests pin the RDS-warning copy and the button-state invariant.

**Files:**
- Modify: `VSTO2/OutlookAI/SettingsForm.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/SettingsForm_UpdatesSection_SourceTests.cs`

- [ ] **Step 1: Write source-level tests (failing)**

```csharp
using System;
using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class SettingsForm_UpdatesSection_SourceTests
    {
        private static string FindSourceFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate)) return candidate;
                current = current.Parent;
            }
            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }

        private static string SettingsFormSource =>
            File.ReadAllText(FindSourceFile("OutlookAI", "SettingsForm.cs"));

        [Fact]
        public void SettingsForm_HasUpdatesGroupBoxWithExpectedControls()
        {
            var src = SettingsFormSource;
            Assert.Contains("\"Updates\"", src);                       // GroupBox text
            Assert.Contains("_lblCurrentVersion", src);
            Assert.Contains("_lblLatestVersion", src);
            Assert.Contains("_lblLastChecked", src);
            Assert.Contains("_btnCheckNow", src);
            Assert.Contains("_btnInstallUpdate", src);
            Assert.Contains("_lblUpdateStatus", src);
        }

        [Fact]
        public void SettingsForm_InstallUpdateButton_DisabledUntilNewerAvailable()
        {
            var src = SettingsFormSource;
            // Button starts disabled
            Assert.Contains("_btnInstallUpdate.Enabled = false", src);
            // ... and is only enabled when the comparator says newer
            Assert.Contains("UpdateAvailability.NewerAvailable", src);
        }

        [Fact]
        public void SettingsForm_InstallClick_ShowsRdsWarningBeforeInstalling()
        {
            var src = SettingsFormSource;
            // Pin the operative copy so a future edit can't remove the
            // "all users" warning without breaking this test.
            Assert.Contains("close Outlook for ALL users", src);
            Assert.Contains("MessageBoxButtons.OKCancel", src);
        }

        [Fact]
        public void SettingsForm_UsesGitHubReleaseClient_AndPassesUserAgentFromInstalledTag()
        {
            var src = SettingsFormSource;
            Assert.Contains("new GitHubReleaseClient(", src);
            Assert.Contains("OutlookAI-Updater/", src);
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile or fail to find substrings** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SettingsForm_UpdatesSection" 2>&1 |
  Select-String -Pattern "Failed|Passed" | Select-Object -Last 3
```

Expected: 4 failures with `Assert.Contains() Failure: ... Not found: ...`.

- [ ] **Step 3: Add controls to `SettingsForm`**

In `VSTO2/OutlookAI/SettingsForm.cs`, alongside the existing private fields for `_cmbModel` / `_clbWriteTools` / etc., add:

```csharp
// Updates group
private GroupBox _grpUpdates;
private Label _lblCurrentVersionCaption;
private Label _lblCurrentVersion;
private Label _lblLatestVersionCaption;
private Label _lblLatestVersion;
private Label _lblLastCheckedCaption;
private Label _lblLastChecked;
private Button _btnCheckNow;
private Button _btnInstallUpdate;
private Label _lblUpdateStatus;

// Updater state
private ReleaseInfo _latestRelease;
private UpdateAvailability _availability = UpdateAvailability.NoUpdate;
private DateTimeOffset? _lastCheckedAt;
```

Add `using OutlookAI.Services.Updates;` to the file header alongside the existing usings.

- [ ] **Step 4: Build the Updates GroupBox**

In the form's existing layout method (the same one that builds `_panelSettings` and the AI Behavior group), append a new GroupBox after AI Behavior. Place it at consistent x/y offsets with the existing groups so the form stays readable. Suggested layout (vertical inside the group, 6 rows × ~24 px tall, 8 px padding):

```csharp
_grpUpdates = new GroupBox
{
    Text = "Updates",
    Location = new System.Drawing.Point(12, /* y just below AI Behavior */),
    Size = new System.Drawing.Size(420, 180),
    ForeColor = System.Drawing.Color.Black,
};

_lblCurrentVersionCaption = new Label { Text = "Current:", Location = new System.Drawing.Point(12, 24), AutoSize = true, ForeColor = System.Drawing.Color.Black };
_lblCurrentVersion        = new Label { Text = "—",        Location = new System.Drawing.Point(96, 24), AutoSize = true, ForeColor = System.Drawing.Color.Black };

_lblLatestVersionCaption  = new Label { Text = "Latest:",  Location = new System.Drawing.Point(12, 48), AutoSize = true, ForeColor = System.Drawing.Color.Black };
_lblLatestVersion         = new Label { Text = "—",        Location = new System.Drawing.Point(96, 48), AutoSize = true, ForeColor = System.Drawing.Color.Black };

_lblLastCheckedCaption    = new Label { Text = "Last checked:", Location = new System.Drawing.Point(12, 72), AutoSize = true, ForeColor = System.Drawing.Color.Black };
_lblLastChecked           = new Label { Text = "—",            Location = new System.Drawing.Point(96, 72), AutoSize = true, ForeColor = System.Drawing.Color.Black };

_btnCheckNow = new Button
{
    Text = "Check Now",
    Location = new System.Drawing.Point(12, 100),
    Size = new System.Drawing.Size(110, 28),
    ForeColor = System.Drawing.Color.Black,
    BackColor = System.Drawing.SystemColors.ButtonFace,
    UseVisualStyleBackColor = false,
};
_btnCheckNow.Click += BtnCheckNow_Click;

_btnInstallUpdate = new Button
{
    Text = "Install Update",
    Location = new System.Drawing.Point(130, 100),
    Size = new System.Drawing.Size(130, 28),
    ForeColor = System.Drawing.Color.Black,
    BackColor = System.Drawing.SystemColors.ButtonFace,
    UseVisualStyleBackColor = false,
    Enabled = false,
};
_btnInstallUpdate.Click += BtnInstallUpdate_Click;

_lblUpdateStatus = new Label
{
    Text = "",
    Location = new System.Drawing.Point(12, 138),
    AutoSize = false,
    Size = new System.Drawing.Size(395, 32),
    ForeColor = System.Drawing.Color.Black,
};

_grpUpdates.Controls.AddRange(new System.Windows.Forms.Control[] {
    _lblCurrentVersionCaption, _lblCurrentVersion,
    _lblLatestVersionCaption,  _lblLatestVersion,
    _lblLastCheckedCaption,    _lblLastChecked,
    _btnCheckNow, _btnInstallUpdate, _lblUpdateStatus,
});
_panelSettings.Controls.Add(_grpUpdates);
```

In the same method, also populate the current version from disk:

```csharp
_lblCurrentVersion.Text = UpdateManifest.LoadFromInstallDir().Tag;
```

- [ ] **Step 5: Add the click handlers**

Append these methods to `SettingsForm`:

```csharp
private static readonly System.Net.Http.HttpClient _updaterHttp = new System.Net.Http.HttpClient();

private async void BtnCheckNow_Click(object sender, EventArgs e)
{
    _btnCheckNow.Enabled = false;
    _btnInstallUpdate.Enabled = false;
    _lblUpdateStatus.Text = "Checking…";

    var installed = UpdateManifest.LoadFromInstallDir();
    var ua = "OutlookAI-Updater/" + (installed.IsDevBuild ? "dev" : installed.Tag);
    var client = new GitHubReleaseClient(_updaterHttp, "kirklandsig/OutlookAI", ua);

    var result = await client.GetLatestStableAsync(System.Threading.CancellationToken.None);
    _lastCheckedAt = DateTimeOffset.Now;
    _lblLastChecked.Text = _lastCheckedAt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");

    switch (result)
    {
        case ReleaseFound found:
            _latestRelease = found.Info;
            _lblLatestVersion.Text = found.Info.Tag;
            _availability = VersionComparator.Compare(installed.Tag, found.Info.Tag);
            _btnInstallUpdate.Enabled = _availability == UpdateAvailability.NewerAvailable;
            _lblUpdateStatus.Text = _availability == UpdateAvailability.NewerAvailable
                ? ("Update available: " + found.Info.Tag)
                : (_availability == UpdateAvailability.NoUpdate ? "Already up to date." :
                   _availability == UpdateAvailability.OlderThanInstalled ? "Latest is older than installed (unusual)." :
                   "Latest tag could not be compared to installed version.");
            break;
        case NoReleasesAvailable _:
            _lblLatestVersion.Text = "—";
            _lblUpdateStatus.Text = "No releases published yet on GitHub.";
            break;
        case RateLimited rl:
            _lblLatestVersion.Text = "—";
            _lblUpdateStatus.Text = "GitHub rate limit hit. Try again after " + rl.ResetAt.ToLocalTime().ToString("HH:mm") + ".";
            break;
        case NetworkError ne:
            _lblLatestVersion.Text = "—";
            _lblUpdateStatus.Text = "Could not reach GitHub: " + ne.Detail;
            break;
    }

    new UpdateHistoryLog().Append("check",
        result.GetType().Name.ToLowerInvariant(),
        (_latestRelease != null ? _latestRelease.Tag : ""),
        _lblUpdateStatus.Text);

    _btnCheckNow.Enabled = true;
}

private async void BtnInstallUpdate_Click(object sender, EventArgs e)
{
    if (_latestRelease == null || _availability != UpdateAvailability.NewerAvailable) return;

    var confirm = System.Windows.Forms.MessageBox.Show(
        text:
            "Install OutlookAI " + _latestRelease.Tag + ".\n\n" +
            "This will:\n" +
            "  • close Outlook for ALL users currently on this server\n" +
            "  • run the OutlookAI installer with administrator privileges\n" +
            "  • leave Outlook closed when finished — everyone reopens manually\n\n" +
            "Have you given users a heads-up?",
        caption: "Install Update",
        buttons: System.Windows.Forms.MessageBoxButtons.OKCancel,
        icon: System.Windows.Forms.MessageBoxIcon.Warning,
        defaultButton: System.Windows.Forms.MessageBoxDefaultButton.Button2);
    if (confirm != System.Windows.Forms.DialogResult.OK) return;

    _btnInstallUpdate.Enabled = false;
    _btnCheckNow.Enabled = false;
    _lblUpdateStatus.Text = "Downloading…";

    var downloader = new UpdateDownloader(_updaterHttp);
    var dl = await downloader.DownloadAsync(_latestRelease, null, System.Threading.CancellationToken.None);

    if (!(dl is DownloadSuccess success))
    {
        _lblUpdateStatus.Text = dl switch
        {
            HashMismatch _ => "Downloaded file failed integrity check. Aborting.",
            MissingInstallerScript _ => "Update package is malformed (no installer). Please file a bug.",
            DownloadFailed df => "Download failed: " + df.Detail,
            Cancelled _ => "Cancelled.",
            _ => "Unknown download result.",
        };
        new UpdateHistoryLog().Append("download", "failed", _latestRelease.Tag, _lblUpdateStatus.Text);
        _btnInstallUpdate.Enabled = true;
        _btnCheckNow.Enabled = true;
        return;
    }
    new UpdateHistoryLog().Append("download", "ok", _latestRelease.Tag, "sha256_ok");

    // Write sentinel; cleared by ThisAddIn.Startup on next Outlook start.
    try
    {
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(UpdatePaths.InProgressSentinel));
        System.IO.File.WriteAllText(UpdatePaths.InProgressSentinel, _latestRelease.Tag);
    } catch { }

    var installer = new UpdateInstaller();
    var launch = installer.LaunchElevatedInstall(success);

    switch (launch)
    {
        case Launched l:
            _lblUpdateStatus.Text = "Installer launched (PID " + l.Pid + "). Outlook will close shortly to apply the update.";
            new UpdateHistoryLog().Append("launch", "launched", _latestRelease.Tag, "pid=" + l.Pid);
            break;
        case UacDeclined _:
            _lblUpdateStatus.Text = "Update cancelled — administrator privileges required.";
            try { System.IO.File.Delete(UpdatePaths.InProgressSentinel); } catch { }
            _btnInstallUpdate.Enabled = true;
            _btnCheckNow.Enabled = true;
            new UpdateHistoryLog().Append("launch", "uac_declined", _latestRelease.Tag, "");
            break;
        case LaunchFailed lf:
            _lblUpdateStatus.Text = "Failed to launch installer: " + lf.Detail;
            try { System.IO.File.Delete(UpdatePaths.InProgressSentinel); } catch { }
            _btnInstallUpdate.Enabled = true;
            _btnCheckNow.Enabled = true;
            new UpdateHistoryLog().Append("launch", "failed", _latestRelease.Tag, lf.Detail);
            break;
    }
}
```

Notes:
- The `switch` expression syntax (`dl switch { … }`) requires C# 8. The project's csproj already targets `<LangVersion>8</LangVersion>` (verify; if not, set it to `8.0` or expand the expression to a classic switch statement).
- Existing buttons in this form all use `ForeColor = Color.Black; BackColor = SystemColors.ButtonFace; UseVisualStyleBackColor = false;` to dodge the Server 2025 white-on-white issue. The new buttons follow the same pattern.

- [ ] **Step 6: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SettingsForm_UpdatesSection"
```

Expected: 4 tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/SettingsForm.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/SettingsForm_UpdatesSection_SourceTests.cs
git commit -m "feat(updater): add Updates section to Settings form"
```

---

## Task 9: `ThisAddIn.Startup` sentinel handling

Clear `%LOCALAPPDATA%\OutlookAI\Updates\.in-progress` on Outlook startup if either (a) the install succeeded (installed `version.json` tag now matches the sentinel's tag) or (b) the sentinel is older than 30 minutes (assume aborted).

**Files:**
- Modify: `VSTO2/OutlookAI/ThisAddIn.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Updates/UpdateStartupReconcilerTests.cs`
- Create: `VSTO2/OutlookAI/Services/Updates/UpdateStartupReconciler.cs` (factor the logic out of `ThisAddIn` so it can be unit-tested without spinning up Outlook)
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Write `UpdateStartupReconcilerTests.cs` (failing)**

```csharp
using System;
using System.IO;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateStartupReconcilerTests : IDisposable
    {
        private readonly string _tempBase;
        private readonly string _originalBaseDir;
        private readonly string _originalInstalledVersionJson;

        public UpdateStartupReconcilerTests()
        {
            _tempBase = Path.Combine(Path.GetTempPath(), "updater-startup-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(_tempBase);

            _originalBaseDir = UpdatePaths.BaseUpdatesDir;
            _originalInstalledVersionJson = UpdatePaths.InstalledVersionJson;

            UpdatePaths.BaseUpdatesDir = Path.Combine(_tempBase, "Updates");
            UpdatePaths.InstalledVersionJson = Path.Combine(_tempBase, "version.json");
            Directory.CreateDirectory(UpdatePaths.BaseUpdatesDir);
        }

        public void Dispose()
        {
            UpdatePaths.BaseUpdatesDir = _originalBaseDir;
            UpdatePaths.InstalledVersionJson = _originalInstalledVersionJson;
            try { Directory.Delete(_tempBase, true); } catch { }
        }

        private void WriteSentinel(string tag, DateTime? lastWriteUtc = null)
        {
            File.WriteAllText(UpdatePaths.InProgressSentinel, tag);
            if (lastWriteUtc.HasValue) File.SetLastWriteTimeUtc(UpdatePaths.InProgressSentinel, lastWriteUtc.Value);
        }

        private void WriteInstalledTag(string tag)
        {
            File.WriteAllText(UpdatePaths.InstalledVersionJson,
                "{\"tag\":\"" + tag + "\",\"commit\":\"-\",\"build_date\":\"2026-01-01T00:00:00Z\",\"repo\":\"x/y\"}");
        }

        [Fact]
        public void Reconcile_InstalledMatchesSentinel_ClearsAndLogsSuccess()
        {
            WriteSentinel("v2.1.0");
            WriteInstalledTag("v2.1.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            var entries = log.ReadAll();
            Assert.Contains(entries, e => e.Action == "install" && e.Result == "succeeded" && e.Tag == "v2.1.0");
        }

        [Fact]
        public void Reconcile_InstalledStillOld_SentinelStaleMoreThan30Min_ClearsAndLogsAborted()
        {
            WriteSentinel("v2.1.0", DateTime.UtcNow.AddMinutes(-31));
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            var entries = log.ReadAll();
            Assert.Contains(entries, e => e.Action == "install" && e.Result == "aborted" && e.Tag == "v2.1.0");
        }

        [Fact]
        public void Reconcile_InstalledStillOld_SentinelFresh_LeavesSentinelAlone()
        {
            WriteSentinel("v2.1.0", DateTime.UtcNow.AddMinutes(-5));
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.True(File.Exists(UpdatePaths.InProgressSentinel));
            Assert.Empty(log.ReadAll());
        }

        [Fact]
        public void Reconcile_NoSentinel_NoOp()
        {
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            Assert.Empty(log.ReadAll());
        }
    }
}
```

- [ ] **Step 2: Verify the tests fail to compile** (RED)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo 2>&1 |
  Select-String -Pattern "error CS" | Select-Object -First 3
```

Expected: `CS0246: ... 'UpdateStartupReconciler' ...`.

- [ ] **Step 3: Create `UpdateStartupReconciler.cs`**

```csharp
using System;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Runs on every Outlook startup. Reconciles the in-progress sentinel
    /// with the actually-installed version.json:
    ///  - if installed tag matches sentinel tag -> install succeeded, clear + log
    ///  - if installed tag differs and sentinel is older than 30 min -> assume aborted, clear + log
    ///  - otherwise leave the sentinel alone
    /// </summary>
    public static class UpdateStartupReconciler
    {
        public static readonly TimeSpan StaleAfter = TimeSpan.FromMinutes(30);

        public static void Reconcile(UpdateHistoryLog history)
        {
            try
            {
                if (!File.Exists(UpdatePaths.InProgressSentinel)) return;
                var sentinelTag = (File.ReadAllText(UpdatePaths.InProgressSentinel) ?? "").Trim();
                var installed = UpdateManifest.LoadFromInstallDir();

                if (!installed.IsDevBuild && string.Equals(installed.Tag, sentinelTag, StringComparison.Ordinal))
                {
                    File.Delete(UpdatePaths.InProgressSentinel);
                    history?.Append("install", "succeeded", sentinelTag, "");
                    return;
                }

                var age = DateTime.UtcNow - File.GetLastWriteTimeUtc(UpdatePaths.InProgressSentinel);
                if (age > StaleAfter)
                {
                    File.Delete(UpdatePaths.InProgressSentinel);
                    history?.Append("install", "aborted", sentinelTag, "sentinel stale > 30 min");
                }
            }
            catch
            {
                // Best-effort; never break Outlook startup over a stale sentinel.
            }
        }
    }
}
```

- [ ] **Step 4: Wire into `ThisAddIn.Startup`**

In `VSTO2/OutlookAI/ThisAddIn.cs`, locate the `Startup` event handler. Near the top, after config load and before any heavy work, add:

```csharp
try { OutlookAI.Services.Updates.UpdateStartupReconciler.Reconcile(new OutlookAI.Services.Updates.UpdateHistoryLog()); }
catch { /* never block startup */ }
```

(Use a fully-qualified namespace to avoid disturbing existing `using` ordering.)

- [ ] **Step 5: Register the new file**

```xml
<Compile Include="Services\Updates\UpdateStartupReconciler.cs" />
```

- [ ] **Step 6: Build and run** (GREEN)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~UpdateStartupReconcilerTests"
```

Expected: 4 tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Updates/UpdateStartupReconciler.cs `
        VSTO2/OutlookAI/ThisAddIn.cs `
        VSTO2/OutlookAI.Tests/Services/Updates/UpdateStartupReconcilerTests.cs `
        VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(updater): reconcile in-progress sentinel on Outlook startup"
```

---

## Task 10: Verification gate

Confirm the cumulative change builds, all tests pass, and no stray TODOs / dead code landed.

**Files:** none (verification only).

- [ ] **Step 1: WebUI syntax check**

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
```

Expected: both commands exit 0 with no output.

- [ ] **Step 2: Full Debug build**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds. Only the pre-existing `MSB3277` warning emits.

- [ ] **Step 3: Full VSTest suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Total tests: ≥ 580   Passed: same   Failed: 0`.

Component coverage tally (approximate):

| Component | Tests added |
|---|---|
| `UpdateManifestTests` | 5 |
| `VersionComparatorTests` | 14 (10 Theory + 4 unparseable) |
| `GitHubReleaseClientTests` | 5 |
| `UpdateDownloaderTests` | 5 |
| `UpdateInstallerSourceTests` | 2 |
| `UpdateHistoryLogTests` | 4 |
| `UpdateStartupReconcilerTests` | 4 |
| `SettingsForm_UpdatesSection_SourceTests` | 4 |
| **New total** | **43** → baseline `553` → expected `≥ 596` |

- [ ] **Step 4: Working tree clean**

```powershell
git status --short
```

Expected: empty output.

- [ ] **Step 5: No stray `TODO` / `FIXME` in newly added files**

```powershell
Select-String -Path "VSTO2\OutlookAI\Services\Updates\*.cs",
                    "VSTO2\OutlookAI.Tests\Services\Updates\*.cs" `
              -Pattern "TODO|FIXME"
```

Expected: empty output.

---

## Task 11: PR + merge + first release `v2.1.0`

The updater needs at least one release on GitHub to be meaningful, and the release workflow needs to have been merged to `master` before tagging will trigger it. Order:

1. Open PR with the branch.
2. Merge to `master`.
3. Tag `v2.1.0` from `master`; the workflow builds + creates the release.
4. Install `v2.1.0` manually via `Install-OutlookAI.ps1`.
5. Confirm Settings → Updates shows `Current: v2.1.0`.

**Files:** none (deployment + GitHub only).

- [ ] **Step 1: Push the feature branch**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git push -u origin feature/in-app-updater
```

- [ ] **Step 2: Open the PR via `--body-file`** (PowerShell + `gh` gotcha — see `handoff.md` §11)

```powershell
$tmp = New-TemporaryFile
Set-Content -LiteralPath $tmp.FullName -Encoding UTF8 -Value @'
## Summary
Adds an in-app "Check for updates" / "Install Update" workflow to the admin-password-gated Settings dialog. Reuses `Install-OutlookAI.ps1` verbatim for the elevated install step.

- New `Services/Updates/` namespace: `UpdateManifest`, `SemVer`, `VersionComparator`, `GitHubReleaseClient`, `UpdateDownloader`, `UpdateInstaller`, `UpdateHistoryLog`, `UpdateStartupReconciler`, `UpdatePaths`, and their typed result hierarchies.
- New `version.json` baked into Release ZIPs at build time; installer copies it into `C:\Program Files\OutlookAI\` so the running add-in knows its own tag.
- New `Settings → Updates` section: current version, latest, last-checked, Check Now, Install Update, status line. Install button gated by `VersionComparator.Compare` returning `NewerAvailable`. Loud RDS confirmation dialog before launching.
- New `.github/workflows/release.yml`: triggers on `v*` tag push, builds Release publish, generates `version.json`, packages the deploy ZIP, computes SHA256 sidecar, creates a GitHub Release with both assets.
- New `Deploy/Make-ReleaseZip.ps1`: local equivalent of the CI workflow; also serves as the canonical build script the CI mirrors.
- `ThisAddIn.Startup` reconciles the `.in-progress` sentinel against the installed tag on every startup (clears on success or after 30 min staleness).

## Test Plan
- `node --check` on `chat.js` and `markdown.js`: both pass.
- VS MSBuild Debug Any CPU: succeeds with only the existing `MSB3277` warning.
- VSTest: `~596` passing (baseline 553 + 43 new across the eight new test classes).
- Local dry run of `Deploy\Make-ReleaseZip.ps1 -Tag v2.1.0-dryrun -OutDir out-dryrun`: produces a valid ZIP + sidecar; ZIP contains `version.json`.

Spec: `docs/superpowers/specs/2026-05-20-in-app-updater-design.md`
Plan: `docs/superpowers/plans/2026-05-20-in-app-updater.md`

## Smoke Plan (post-merge)
1. `git tag -a v2.1.0 -m "OutlookAI v2.1.0 — in-app updater"` on `master`; `git push --tags`.
2. Wait for `.github/workflows/release.yml` to finish. Confirm release page shows `OutlookAI-v2.1.0-RDS-Deploy.zip` + `.sha256`.
3. Download the ZIP locally; install via `Install-OutlookAI.ps1 -SourcePath <unzipped>` elevated.
4. Open Outlook → Settings → Updates. Expect `Current: v2.1.0`.
5. Future trivial fix → tag `v2.1.1` → in the running v2.1.0 install, click `Check Now`, expect `Update available: v2.1.1`, click `Install Update`, accept UAC, confirm install transcript, reopen Outlook, expect `Current: v2.1.1`.
'@
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr create --repo kirklandsig/OutlookAI --base master --head feature/in-app-updater `
    --title "Add in-app updater (Settings → Updates + GitHub Releases workflow)" `
    --body-file $tmp.FullName
Remove-Item -LiteralPath $tmp.FullName -Force
```

Expected: the command prints the PR URL.

- [ ] **Step 3: Merge the PR**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr view --json mergeStateStatus,mergeable
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr merge --merge --subject "Merge: in-app updater"
```

Expected: `mergeable: MERGEABLE`, then `Merged pull request`.

- [ ] **Step 4: Fast-forward local `master`**

```powershell
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" fetch origin master
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" status -sb
# If the main worktree has an auto-bumped ApplicationVersion in OutlookAI.csproj,
# discard it before pulling (it's a build artifact that re-bumps on next build):
$dirty = git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" status -s | Where-Object { $_ -match "OutlookAI\.csproj" }
if ($dirty) {
    git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" restore VSTO2/OutlookAI/OutlookAI.csproj
}
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" pull --ff-only origin master
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" log -1 --oneline
```

Expected: latest commit on `master` is the merge.

- [ ] **Step 5: Tag `v2.1.0` and push**

```powershell
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" tag -a v2.1.0 -m "OutlookAI v2.1.0 — in-app updater"
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" push origin v2.1.0
```

- [ ] **Step 6: Wait for the release workflow to finish and verify**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh run list --repo kirklandsig/OutlookAI --workflow release.yml --limit 1
# Watch the latest run:
gh run watch --repo kirklandsig/OutlookAI
```

Expected: workflow succeeds. Then:

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh release view v2.1.0 --repo kirklandsig/OutlookAI --json assets,tagName --jq '{tag: .tagName, assets: [.assets[] | .name]}'
```

Expected: `tag = "v2.1.0"` with both `OutlookAI-v2.1.0-RDS-Deploy.zip` and `…zip.sha256` in the assets array.

- [ ] **Step 7: Download and install `v2.1.0`**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh release download v2.1.0 --repo kirklandsig/OutlookAI --pattern "OutlookAI-v2.1.0-RDS-Deploy.zip" --dir $staging
Expand-Archive -LiteralPath (Join-Path $staging "OutlookAI-v2.1.0-RDS-Deploy.zip") -DestinationPath $staging -Force
# Close Outlook if running.
$procs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) { $procs | Stop-Process -Force }
# Elevated install.
$script = Join-Path $staging "Install-OutlookAI.ps1"
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -SourcePath `"$staging`""
$proc = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -Wait -PassThru
"installer exit=$($proc.ExitCode)"
```

Expected: installer exits 0.

- [ ] **Step 8: Verify `version.json` landed in install dir**

```powershell
Get-Content -LiteralPath "C:\Program Files\OutlookAI\version.json"
```

Expected: JSON with `"tag": "v2.1.0"`.

- [ ] **Step 9: Smoke the UI**

Open Outlook → AI Assistant → gear icon → Settings → enter admin password → look at the new Updates section. Expect:
- `Current: v2.1.0`
- `Latest: —`
- `Last checked: —`
- `Install Update` button disabled.

Click `Check Now`. Expect:
- `Latest: v2.1.0`
- `Already up to date.`
- `Install Update` still disabled.

This proves the wiring works end-to-end against a real GitHub release. The actual install-flow smoke happens after tagging `v2.1.1` in a follow-up commit / session (any trivial change, then `git tag v2.1.1`, then watch the updater find it). This plan does not require that follow-up to be done in the same session.

---

## Task 12: Cleanup

Same shape as the previous PR's cleanup (after the `issues-2-5-cleanup` merge).

**Files:** none.

- [ ] **Step 1: Remove the feature worktree and local branch**

From the main worktree root:

```powershell
$top = "C:\Users\MDASR\Desktop\Projects\OutlookAI"
$wt  = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-in-app-updater"
git -C $top worktree remove $wt
git -C $top worktree prune
git -C $top branch -d feature/in-app-updater
```

- [ ] **Step 2: Delete the remote feature branch**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" push origin --delete feature/in-app-updater
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" branch -a
```

Expected: only `master` and `remotes/origin/master` remain. `v2.1.0` tag stays.

- [ ] **Step 3: Update `handoff.md`** (local, gitignored)

In `C:\Users\MDASR\Desktop\Projects\OutlookAI\handoff.md`:

- §1 Snapshot: bump the latest commit hash (use `git log -1 --oneline` on `master`), bump the installed `OutlookAI.dll` SHA256 (use `Get-FileHash`), update test count to the new total.
- §12 Recent history: append a one-line entry: "Shipped v2.1.0 with in-app updater (Settings → Updates section, GitHub Releases workflow, `version.json` baked into deploy ZIPs)."
- §13 Open follow-ups: optionally add "v2.1.1 trivial-change smoke" if not yet done.

This is the only post-merge maintenance step. `handoff.md` is not committed.

---

## Self-Review

**1. Spec coverage:**

| Spec requirement | Task |
|---|---|
| Release workflow + tag-triggered Action | Task 1 |
| `version.json` baked + installer copies to install dir | Task 1 |
| `UpdateManifest` + `UpdatePaths` | Task 2 |
| `SemVer` + `VersionComparator` + `UpdateAvailability` | Task 3 |
| `ReleaseInfo` + `ReleaseLookupResult` hierarchy + `GitHubReleaseClient` | Task 4 |
| `DownloadResult` hierarchy + `UpdateDownloader` | Task 5 |
| `LaunchResult` hierarchy + `UpdateInstaller` (detached elevated launch) | Task 6 |
| `UpdateHistoryLog` (50-entry cap, malformed-tolerant) | Task 7 |
| Settings UI Updates section + RDS-warning copy + button gating | Task 8 |
| `ThisAddIn` sentinel reconcile on startup | Task 9 |
| Failure modes (network down, 404, hash mismatch, UAC declined, malformed ZIP) | covered across tasks 4-9 |
| Verification (build, test, no TODOs) | Task 10 |
| First release + PR + merge + install + smoke | Task 11 |
| Cleanup + handoff | Task 12 |

**2. Placeholder scan:** every code-changing step shows actual code. No "TBD", "TODO", "similar to Task N", "handle errors", etc.

**3. Type / name consistency:**

- `UpdateAvailability` enum referenced in Tasks 3 and 8 ✓
- `ReleaseFound` / `NoReleasesAvailable` / `RateLimited` / `NetworkError` referenced in Tasks 4 and 8 ✓
- `DownloadSuccess` / `HashMismatch` / `DownloadFailed` / `MissingInstallerScript` / `Cancelled` referenced in Tasks 5 and 8 ✓
- `Launched` / `UacDeclined` / `LaunchFailed` referenced in Tasks 6 and 8 ✓
- `UpdatePaths.BaseUpdatesDir`, `InstalledVersionJson`, `InProgressSentinel`, `HistoryLog` all defined in Task 2 and used in Tasks 5, 7, 8, 9 ✓
- `UpdateHistoryLog.Entry` shape: `Ts`, `Action`, `Result`, `Tag`, `Details` consistent across Tasks 7, 8, 9 ✓
- `version.json` schema (`tag`, `commit`, `build_date`, `repo`) consistent across Task 1 (CI) and Task 2 (parser) ✓

---

## Final State After This Plan

- `master` advances by one merge commit. New tag `v2.1.0` exists on the remote.
- A GitHub Release at `v2.1.0` carries the deploy ZIP + SHA256 sidecar.
- Installed copies of OutlookAI carry `version.json` and can self-check against GitHub.
- The admin can update OutlookAI from inside Outlook (Settings → Updates → Check Now → Install Update → UAC) without re-running scripts manually.
- New test surface: 8 test classes, ~43 tests. Total suite ≈ 596 passing.
- `feature/in-app-updater` branch and worktree are gone.
- `handoff.md` reflects the new state.

