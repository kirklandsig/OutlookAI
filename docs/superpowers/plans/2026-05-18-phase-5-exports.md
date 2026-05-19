# Phase 5: Excel & PDF Exports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship two model-callable export tools (`outlook_export_excel`, `outlook_export_pdf`) and a per-message "Save as PDF" button so any chat output can be exported. Both tools save to `~\Documents\OutlookAI\Reports\` with auto-generated timestamped filenames and surface inline FileCards with `Open` / `Show in folder` actions.

**Architecture:** Two single-purpose `IOutlookTool`s added to the existing catalog. `LiveOutlookSurface` grows `ExportExcel` (via ClosedXML) and `ExportPdf` (via an off-screen `WebView2` instance + `PrintToPdfAsync`). Existing markdown renderer in `chat.js` is factored out to a shared `markdown.js` so the PDF print template renders identically to chat. A new `ExportBridge` helper handles three new `WebMessageReceived` types (`export_pdf`, `open_file`, `reveal_in_explorer`) and is wired into all three chat controllers (Chat, Inbox Copilot, Inbox Reports). File paths handed to `Open` / `Reveal` are validated against the canonical Reports dir via `IExportPathPolicy` to block traversal.

**Tech Stack:** C# .NET Framework 4.7.2 (VSTO), xUnit, ClosedXML 0.102.x, WebView2, `Microsoft.Office.Interop.Outlook`.

**Spec:** `docs/superpowers/specs/2026-05-18-phase-5-exports-design.md`

**Current status (Task 25):** Tasks 1-24 are implemented and reviewed. Debug verification before this docs pass was `499/499` tests passing. Tasks 26-27 (publish/install and live smoke) and Task 28 (push) remain pending.

---

## File Structure

**New files:**

- `VSTO2/OutlookAI/Services/Export/ExportPathResolver.cs` — resolves `~\Documents\OutlookAI\Reports\`, idempotent create.
- `VSTO2/OutlookAI/Services/Export/ExportFilenameSanitizer.cs` — pure helper, sanitize + timestamp + collision suffix.
- `VSTO2/OutlookAI/Services/Export/IExportPathPolicy.cs` — bridge security gate interface.
- `VSTO2/OutlookAI/Services/Export/ExportPathPolicy.cs` — impl, validates path under Reports dir.
- `VSTO2/OutlookAI/Services/Export/UnauthorizedExportPathException.cs` — thrown by policy.
- `VSTO2/OutlookAI/Services/Export/ExportException.cs` — wraps runtime export failures with a `Code` string.
- `VSTO2/OutlookAI/Services/Export/ExcelColumnType.cs` — enum + parser.
- `VSTO2/OutlookAI/Services/Export/ExcelCellCoercer.cs` — pure helper, coerces cell values per column type.
- `VSTO2/OutlookAI/Services/Export/ExcelWorkbookBuilder.cs` — pure builder, returns `XLWorkbook` in memory.
- `VSTO2/OutlookAI/Services/Export/PrintTemplateRenderer.cs` — pure renderer for the PDF print HTML.
- `VSTO2/OutlookAI/Services/Export/PdfRenderer.cs` — owns the off-screen WebView2 + `PrintToPdfAsync`.
- `VSTO2/OutlookAI/Services/Export/FileSavedResult.cs` — DTO returned by both export surface methods.
- `VSTO2/OutlookAI/Services/Tools/ExportExcelArgs.cs` — tool args DTO.
- `VSTO2/OutlookAI/Services/Tools/ExportExcelArgsParser.cs` — pure JSON parser.
- `VSTO2/OutlookAI/Services/Tools/OutlookExportExcelTool.cs` — `IOutlookTool` impl.
- `VSTO2/OutlookAI/Services/Tools/ExportPdfArgs.cs` — tool args DTO.
- `VSTO2/OutlookAI/Services/Tools/ExportPdfArgsParser.cs` — pure JSON parser.
- `VSTO2/OutlookAI/Services/Tools/OutlookExportPdfTool.cs` — `IOutlookTool` impl.
- `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs` — handles `export_pdf` / `open_file` / `reveal_in_explorer` web messages.
- `VSTO2/OutlookAI/WebUI/markdown.js` — extracted markdown renderer, shared by chat and PDF.
- `VSTO2/OutlookAI/WebUI/print-template.html` — PDF page skeleton.
- `VSTO2/OutlookAI/WebUI/print-styles.css` — print-tuned CSS (A4, frozen table headers, etc.).
- Matching test files under `VSTO2/OutlookAI.Tests/Services/Export/` and `Services/Tools/`.

**Modified files:**

- `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` — add `ExportExcel` + `ExportPdf` signatures.
- `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` — implement both new surface methods.
- `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` — schemas for the two new tools.
- `VSTO2/OutlookAI/Services/OutlookToolHost.cs` — register both tools.
- `VSTO2/OutlookAI/TaskPane/Chat/ChatController.cs` — wire `ExportBridge` into `HandleHostMessage`.
- `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotController.cs` — wire `ExportBridge`.
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs` — wire `ExportBridge`.
- `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` — `NullSurface` (or `MinimalSurface`) overrides for the new signatures.
- `VSTO2/OutlookAI/ThisAddIn.cs` — construct singleton `PdfRenderer` + `IExportPathPolicy`; dispose on shutdown.
- `VSTO2/OutlookAI/OutlookAI.csproj` — add ClosedXML PackageReference, register all new compile items and WebUI embedded resources.
- `VSTO2/OutlookAI/WebUI/chat.js` — call into `markdown.js` instead of inline renderer; add file-card rendering, per-message PDF button, export-error rendering; post `export_pdf` / `open_file` / `reveal_in_explorer` messages.
- `VSTO2/OutlookAI/WebUI/styles.css` — `.msg-action`, `.file-card`, `.error-card` styles.
- `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs` — virtual overrides for `ExportExcel` / `ExportPdf`.

**Build & test commands** (use throughout this plan):

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Run from `C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration`.

Baseline before starting: **317 tests passing**. Target at end: **~413 passing**.

---

## Task 1: Add ClosedXML NuGet dependency

ClosedXML is the Excel-writing library. Pin to 0.102.x (stable, net472-compatible, MIT-licensed). Verifies the baseline build + test suite still passes with the new dependency before touching any new code.

**Files:**
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`
- Modify: `VSTO2/OutlookAI/packages.config` (if present — the project uses old-style packages.config for net472).

- [ ] **Step 1: Identify how NuGet packages are currently added**

Look for `<PackageReference>` entries in `OutlookAI.csproj`:

```powershell
Select-String -Path "VSTO2\OutlookAI\OutlookAI.csproj" -Pattern "PackageReference|HintPath" | Select-Object -First 20
```

If `PackageReference` style is in use, follow Step 2A. If `packages.config` style is in use (HintPath references under `..\packages\`), follow Step 2B.

- [ ] **Step 2A: Add PackageReference (if project uses PackageReference style)**

In `VSTO2/OutlookAI/OutlookAI.csproj`, find the `<ItemGroup>` that contains existing `PackageReference` entries (e.g. Newtonsoft.Json or Microsoft.Web.WebView2). Add:

```xml
<PackageReference Include="ClosedXML" Version="0.102.3" />
```

- [ ] **Step 2B: Add via packages.config (if project uses packages.config style)**

Edit `VSTO2/OutlookAI/packages.config`. Inside `<packages>` add:

```xml
<package id="ClosedXML" version="0.102.3" targetFramework="net472" />
<package id="DocumentFormat.OpenXml" version="2.20.0" targetFramework="net472" />
<package id="ExcelNumberFormat" version="1.1.0" targetFramework="net472" />
<package id="FastMember.Signed" version="1.5.0" targetFramework="net472" />
<package id="SixLabors.Fonts" version="1.0.0" targetFramework="net472" />
<package id="System.IO.Packaging" version="6.0.0" targetFramework="net472" />
```

Then in `OutlookAI.csproj`, add references in the `<ItemGroup>` with other `<Reference>` entries:

```xml
<Reference Include="ClosedXML">
  <HintPath>..\packages\ClosedXML.0.102.3\lib\net46\ClosedXML.dll</HintPath>
  <Private>True</Private>
</Reference>
<Reference Include="DocumentFormat.OpenXml">
  <HintPath>..\packages\DocumentFormat.OpenXml.2.20.0\lib\net46\DocumentFormat.OpenXml.dll</HintPath>
  <Private>True</Private>
</Reference>
```

(Other transitive deps will be added automatically when you run `nuget restore`.)

- [ ] **Step 3: Restore packages**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /t:Restore /v:minimal /nologo
```

Expected: `Restore completed`. If errors mention missing transitive deps, run:

```powershell
& "C:\ProgramData\chocolatey\bin\nuget.exe" restore "VSTO2\OutlookAI.sln"
```

(Alternative: download `nuget.exe` to a known location and use that path.)

- [ ] **Step 4: Build + run tests to verify the dependency does not break baseline**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: build succeeds; `Passed: 317`.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI/packages.config
git commit -m "chore(deps): add ClosedXML for Excel export"
```

---

## Task 2: `ExportFilenameSanitizer` — pure helper

Generates safe Windows filenames from a user/model-supplied hint. Strips invalid chars (`\ / : * ? " < > |`), trims dots/spaces, truncates to 80 chars, falls back to a default, appends a timestamp suffix, and handles collisions with a numeric suffix.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/ExportFilenameSanitizer.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExportFilenameSanitizerTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Export/ExportFilenameSanitizerTests.cs`:

```csharp
using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportFilenameSanitizerTests
    {
        private static DateTimeOffset FixedNow => new DateTimeOffset(2026, 5, 18, 19, 47, 0, TimeSpan.Zero);

        [Theory]
        [InlineData("IT Creations Quotes", "IT-Creations-Quotes")]
        [InlineData("hello world",         "hello-world")]
        [InlineData("Report: Q1",          "Report-Q1")]
        [InlineData("a/b\\c:d*e?f\"g<h>i|j", "a-b-c-d-e-f-g-h-i-j")]
        public void Sanitize_StripsInvalidCharsAndSpaces(string input, string expectedStem)
        {
            var result = ExportFilenameSanitizer.Build(input, ".xlsx", FixedNow, _ => false);
            Assert.StartsWith(expectedStem + "-2026-05-18-1947", result);
            Assert.EndsWith(".xlsx", result);
        }

        [Fact]
        public void Sanitize_TrailingDotsAndSpacesRemoved()
        {
            var result = ExportFilenameSanitizer.Build("Report...   ", ".xlsx", FixedNow, _ => false);
            Assert.StartsWith("Report-2026-05-18-1947", result);
        }

        [Fact]
        public void Sanitize_EmptyHintFallsBackToDefault()
        {
            var result = ExportFilenameSanitizer.Build("   ", ".xlsx", FixedNow, _ => false);
            Assert.StartsWith("OutlookAI-Report-2026-05-18-1947", result);
        }

        [Fact]
        public void Sanitize_NullHintFallsBackToDefault()
        {
            var result = ExportFilenameSanitizer.Build(null, ".pdf", FixedNow, _ => false);
            Assert.StartsWith("OutlookAI-Report-2026-05-18-1947", result);
            Assert.EndsWith(".pdf", result);
        }

        [Fact]
        public void Sanitize_TruncatesHintTo80Chars()
        {
            var hint = new string('A', 200);
            var result = ExportFilenameSanitizer.Build(hint, ".xlsx", FixedNow, _ => false);
            var stem = Path.GetFileNameWithoutExtension(result);
            // stem = "AAAA...(80 A's)-2026-05-18-1947"; check the A-count ≤ 80
            var aCount = 0;
            foreach (var c in stem) { if (c == 'A') aCount++; else break; }
            Assert.True(aCount <= 80, $"hint truncated to ≤80 chars, got {aCount}");
        }

        [Fact]
        public void Sanitize_AppendsCollisionSuffixWhenExists()
        {
            int callCount = 0;
            bool Exists(string path)
            {
                callCount++;
                // First two probes "exist", third does not.
                return callCount <= 2;
            }
            var result = ExportFilenameSanitizer.Build("Quotes", ".xlsx", FixedNow, Exists);
            Assert.EndsWith("Quotes-2026-05-18-1947-3.xlsx", result);
        }

        [Fact]
        public void Sanitize_DoesNotAddSuffixWhenNoCollision()
        {
            var result = ExportFilenameSanitizer.Build("Quotes", ".xlsx", FixedNow, _ => false);
            Assert.Equal("Quotes-2026-05-18-1947.xlsx", result);
        }

        [Fact]
        public void Sanitize_PreservesExtensionDot()
        {
            var resultPdf = ExportFilenameSanitizer.Build("x", ".pdf", FixedNow, _ => false);
            var resultXlsx = ExportFilenameSanitizer.Build("x", ".xlsx", FixedNow, _ => false);
            Assert.EndsWith(".pdf", resultPdf);
            Assert.EndsWith(".xlsx", resultXlsx);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExportFilenameSanitizerTests"
```

Expected: Build fails because `ExportFilenameSanitizer` doesn't exist yet.

- [ ] **Step 3: Implement `ExportFilenameSanitizer`**

Create `VSTO2/OutlookAI/Services/Export/ExportFilenameSanitizer.cs`:

```csharp
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Pure helper that converts a free-form filename hint into a safe Windows
    /// filename with a timestamp suffix and an optional collision suffix.
    /// </summary>
    public static class ExportFilenameSanitizer
    {
        private const string DefaultStem = "OutlookAI-Report";
        private const int MaxStemLength = 80;
        private static readonly char[] _invalid = Path.GetInvalidFileNameChars();

        /// <summary>
        /// Build a filename. <paramref name="exists"/> is invoked with each
        /// candidate filename (no path) and should return true if a file of
        /// that name already exists in the target directory.
        /// </summary>
        public static string Build(string hint, string extension, DateTimeOffset now, Func<string, bool> exists)
        {
            if (extension == null) throw new ArgumentNullException(nameof(extension));
            if (exists == null) throw new ArgumentNullException(nameof(exists));
            if (!extension.StartsWith(".", StringComparison.Ordinal)) extension = "." + extension;

            var stem = SanitizeStem(hint);
            var timestamp = now.ToString("yyyy-MM-dd-HHmm");
            var baseName = $"{stem}-{timestamp}";

            var candidate = baseName + extension;
            if (!exists(candidate)) return candidate;

            for (int n = 2; n < 1000; n++)
            {
                candidate = $"{baseName}-{n}{extension}";
                if (!exists(candidate)) return candidate;
            }
            // Pathological: 1000 collisions in one minute. Append epoch ticks.
            return $"{baseName}-{now.UtcTicks}{extension}";
        }

        private static string SanitizeStem(string hint)
        {
            if (string.IsNullOrWhiteSpace(hint)) return DefaultStem;

            var sb = new StringBuilder(hint.Length);
            foreach (var c in hint)
            {
                if (_invalid.Contains(c) || char.IsControl(c) || c == ' ')
                {
                    sb.Append('-');
                }
                else
                {
                    sb.Append(c);
                }
            }

            // Collapse runs of '-' to a single dash.
            var collapsed = new StringBuilder(sb.Length);
            bool lastWasDash = false;
            foreach (var c in sb.ToString())
            {
                if (c == '-')
                {
                    if (!lastWasDash) collapsed.Append('-');
                    lastWasDash = true;
                }
                else
                {
                    collapsed.Append(c);
                    lastWasDash = false;
                }
            }

            var result = collapsed.ToString().Trim('-', '.', ' ');
            if (string.IsNullOrEmpty(result)) return DefaultStem;
            if (result.Length > MaxStemLength) result = result.Substring(0, MaxStemLength).TrimEnd('-', '.', ' ');
            return result;
        }
    }
}
```

Register in `VSTO2/OutlookAI/OutlookAI.csproj` — find the `<ItemGroup>` containing `Services\Tools\*.cs` Compile entries and add:

```xml
<Compile Include="Services\Export\ExportFilenameSanitizer.cs" />
```

Also register the test file under `VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj`:

```xml
<Compile Include="Services\Export\ExportFilenameSanitizerTests.cs" />
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExportFilenameSanitizerTests"
```

Expected: all 8 sanitizer tests pass.

- [ ] **Step 5: Run the full suite to verify no regressions**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 325` (317 + 8).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/ExportFilenameSanitizer.cs VSTO2/OutlookAI.Tests/Services/Export/ExportFilenameSanitizerTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add filename sanitizer with collision handling"
```

---

## Task 3: `ExportPathResolver` + `IExportPathPolicy` + `ExportPathPolicy` + `UnauthorizedExportPathException`

Path resolution and security gate. `ExportPathResolver` returns the canonical Reports dir, creates it if missing. `ExportPathPolicy.RequireInsideReportsDir(path)` normalizes a path and throws `UnauthorizedExportPathException` if it escapes the Reports dir or is a UNC path. Used by the bridge's `OpenFile` / `RevealInExplorer` handlers to prevent path-traversal attacks via malicious tool output.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/ExportPathResolver.cs`
- Create: `VSTO2/OutlookAI/Services/Export/IExportPathPolicy.cs`
- Create: `VSTO2/OutlookAI/Services/Export/ExportPathPolicy.cs`
- Create: `VSTO2/OutlookAI/Services/Export/UnauthorizedExportPathException.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExportPathResolverTests.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExportPathPolicyTests.cs`

- [ ] **Step 1: Write the failing tests for the resolver**

Create `VSTO2/OutlookAI.Tests/Services/Export/ExportPathResolverTests.cs`:

```csharp
using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportPathResolverTests
    {
        [Fact]
        public void ResolveBaseDir_ReturnsDocumentsOutlookAIReports()
        {
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var resolver = new ExportPathResolver();
            var path = resolver.ResolveBaseDir();
            Assert.Equal(Path.Combine(docs, "OutlookAI", "Reports"), path);
        }

        [Fact]
        public void EnsureExists_CreatesDirectoryIfMissing()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            try
            {
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);
                Assert.False(Directory.Exists(sandbox));
                resolver.EnsureExists();
                Assert.True(Directory.Exists(sandbox));
            }
            finally
            {
                if (Directory.Exists(sandbox)) Directory.Delete(sandbox, true);
            }
        }

        [Fact]
        public void EnsureExists_IsIdempotent()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(sandbox);
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);
                resolver.EnsureExists();
                resolver.EnsureExists();
                Assert.True(Directory.Exists(sandbox));
            }
            finally
            {
                if (Directory.Exists(sandbox)) Directory.Delete(sandbox, true);
            }
        }

        [Fact]
        public void EnsureExists_ThrowsWhenPathIsAFile()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            File.WriteAllText(sandbox, "not a directory");
            try
            {
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);
                Assert.Throws<IOException>(() => resolver.EnsureExists());
            }
            finally
            {
                if (File.Exists(sandbox)) File.Delete(sandbox);
            }
        }
    }
}
```

- [ ] **Step 2: Write the failing tests for the policy**

Create `VSTO2/OutlookAI.Tests/Services/Export/ExportPathPolicyTests.cs`:

```csharp
using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportPathPolicyTests
    {
        private static (ExportPathPolicy policy, string baseDir) CreateInSandbox()
        {
            var baseDir = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(baseDir);
            return (new ExportPathPolicy(new ExportPathResolver(baseDirOverride: baseDir)), baseDir);
        }

        [Fact]
        public void AcceptsCanonicalPathInsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var ok = Path.Combine(baseDir, "Report.xlsx");
                File.WriteAllText(ok, "");
                policy.RequireInsideReportsDir(ok); // does not throw
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void AcceptsPathCaseInsensitive()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var weirdCase = baseDir.ToUpperInvariant() + Path.DirectorySeparatorChar + "Report.xlsx";
                File.WriteAllText(Path.Combine(baseDir, "Report.xlsx"), "");
                policy.RequireInsideReportsDir(weirdCase); // does not throw
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsTraversalAttempt()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var evil = Path.Combine(baseDir, "..", "..", "Windows", "System32", "cmd.exe");
                Assert.Throws<UnauthorizedExportPathException>(() => policy.RequireInsideReportsDir(evil));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsPathOutsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"C:\Windows\System32\cmd.exe"));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsUncPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"\\server\share\evil.exe"));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsNullPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(null));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsEmptyPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(""));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }

        [Fact]
        public void RejectsRelativePath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir("evil.exe"));
            }
            finally { if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true); }
        }
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExportPath"
```

Expected: build fails (types not defined).

- [ ] **Step 4: Implement `ExportPathResolver`**

Create `VSTO2/OutlookAI/Services/Export/ExportPathResolver.cs`:

```csharp
using System;
using System.IO;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Resolves the canonical base directory for exports
    /// (<c>~\Documents\OutlookAI\Reports\</c>) and ensures it exists.
    /// Caches the resolved path for the process lifetime.
    /// </summary>
    public sealed class ExportPathResolver
    {
        private readonly string _baseDir;

        public ExportPathResolver() : this(baseDirOverride: null) { }

        /// <summary>Test-only ctor to override the base directory.</summary>
        public ExportPathResolver(string baseDirOverride)
        {
            if (baseDirOverride != null)
            {
                _baseDir = Path.GetFullPath(baseDirOverride);
                return;
            }
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _baseDir = Path.Combine(docs, "OutlookAI", "Reports");
        }

        public string ResolveBaseDir() => _baseDir;

        /// <summary>
        /// Create the Reports dir if missing. Throws <see cref="IOException"/>
        /// if the path exists as a file or cannot be created.
        /// </summary>
        public void EnsureExists()
        {
            if (File.Exists(_baseDir))
            {
                throw new IOException(
                    $"Reports path '{_baseDir}' exists as a file, not a directory.");
            }
            Directory.CreateDirectory(_baseDir);
        }
    }
}
```

- [ ] **Step 5: Implement the policy types**

Create `VSTO2/OutlookAI/Services/Export/UnauthorizedExportPathException.cs`:

```csharp
using System;

namespace OutlookAI.Services.Export
{
    public sealed class UnauthorizedExportPathException : Exception
    {
        public UnauthorizedExportPathException(string message) : base(message) { }
    }
}
```

Create `VSTO2/OutlookAI/Services/Export/IExportPathPolicy.cs`:

```csharp
namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Security gate for bridge file-action methods. Implementations must
    /// throw <see cref="UnauthorizedExportPathException"/> for any path
    /// outside the canonical Reports directory.
    /// </summary>
    public interface IExportPathPolicy
    {
        void RequireInsideReportsDir(string path);
    }
}
```

Create `VSTO2/OutlookAI/Services/Export/ExportPathPolicy.cs`:

```csharp
using System;
using System.IO;

namespace OutlookAI.Services.Export
{
    public sealed class ExportPathPolicy : IExportPathPolicy
    {
        private readonly ExportPathResolver _resolver;

        public ExportPathPolicy(ExportPathResolver resolver)
        {
            _resolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
        }

        public void RequireInsideReportsDir(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new UnauthorizedExportPathException("Path is null or empty.");

            // Reject UNC outright.
            if (path.StartsWith(@"\\", StringComparison.Ordinal))
                throw new UnauthorizedExportPathException("UNC paths are not permitted.");

            string fullPath;
            try { fullPath = Path.GetFullPath(path); }
            catch (Exception ex)
            {
                throw new UnauthorizedExportPathException("Path could not be normalized: " + ex.Message);
            }

            // Reject path that didn't resolve absolute (relative paths).
            if (!Path.IsPathRooted(fullPath))
                throw new UnauthorizedExportPathException("Path is not absolute.");

            var baseDir = _resolver.ResolveBaseDir();
            var baseFull = Path.GetFullPath(baseDir);
            if (!baseFull.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                baseFull += Path.DirectorySeparatorChar;

            if (!fullPath.StartsWith(baseFull, StringComparison.OrdinalIgnoreCase))
                throw new UnauthorizedExportPathException(
                    $"Path '{fullPath}' is not inside the Reports directory '{baseFull}'.");
        }
    }
}
```

Register all four new files in `VSTO2/OutlookAI/OutlookAI.csproj` `<ItemGroup>`:

```xml
<Compile Include="Services\Export\ExportPathResolver.cs" />
<Compile Include="Services\Export\IExportPathPolicy.cs" />
<Compile Include="Services\Export\ExportPathPolicy.cs" />
<Compile Include="Services\Export\UnauthorizedExportPathException.cs" />
```

And tests in `VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj`:

```xml
<Compile Include="Services\Export\ExportPathResolverTests.cs" />
<Compile Include="Services\Export\ExportPathPolicyTests.cs" />
```

- [ ] **Step 6: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExportPath"
```

Expected: 12 tests pass (4 resolver + 8 policy).

- [ ] **Step 7: Run the full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 337` (325 + 12).

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/ VSTO2/OutlookAI.Tests/Services/Export/ExportPathResolverTests.cs VSTO2/OutlookAI.Tests/Services/Export/ExportPathPolicyTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add reports directory resolver and path policy"
```

---

## Task 4: `ExcelColumnType` enum + `ExcelCellCoercer`

The model sends columns with a `type` tag (`text|date|datetime|number|currency|boolean`). `ExcelCellCoercer.Coerce(value, type)` converts the cell value (which arrives as `JToken` — string, number, or null) into the strongly-typed object ClosedXML expects, with graceful fallback to text when the value is unparseable.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/ExcelColumnType.cs`
- Create: `VSTO2/OutlookAI/Services/Export/ExcelCellCoercer.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExcelColumnTypeTests.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExcelCellCoercerTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Export/ExcelColumnTypeTests.cs`:

```csharp
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelColumnTypeTests
    {
        [Theory]
        [InlineData("text",     ExcelColumnType.Text)]
        [InlineData("date",     ExcelColumnType.Date)]
        [InlineData("datetime", ExcelColumnType.DateTime)]
        [InlineData("number",   ExcelColumnType.Number)]
        [InlineData("currency", ExcelColumnType.Currency)]
        [InlineData("boolean",  ExcelColumnType.Boolean)]
        [InlineData("TEXT",     ExcelColumnType.Text)]    // case-insensitive
        [InlineData(" date ",   ExcelColumnType.Date)]    // trims
        public void Parse_RecognizesSupportedTypes(string input, ExcelColumnType expected)
        {
            Assert.True(ExcelColumnTypeParser.TryParse(input, out var t));
            Assert.Equal(expected, t);
        }

        [Theory]
        [InlineData("money")]
        [InlineData("integer")]
        [InlineData("string")]
        [InlineData(null)]
        [InlineData("")]
        public void Parse_RejectsUnsupportedTypes(string input)
        {
            Assert.False(ExcelColumnTypeParser.TryParse(input, out _));
        }
    }
}
```

Create `VSTO2/OutlookAI.Tests/Services/Export/ExcelCellCoercerTests.cs`:

```csharp
using System;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelCellCoercerTests
    {
        [Fact]
        public void Coerce_Text_PassesStringThrough()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("hello"), ExcelColumnType.Text);
            Assert.Equal("hello", v);
        }

        [Fact]
        public void Coerce_Text_StringifiesNumbers()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject(42), ExcelColumnType.Text);
            Assert.Equal("42", v);
        }

        [Fact]
        public void Coerce_Date_ParsesIso8601()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("2026-05-18"), ExcelColumnType.Date);
            Assert.IsType<DateTime>(v);
            var d = (DateTime)v;
            Assert.Equal(new DateTime(2026, 5, 18), d.Date);
        }

        [Fact]
        public void Coerce_DateTime_ParsesIsoWithTime()
        {
            var v = ExcelCellCoercer.Coerce(
                JToken.FromObject("2026-05-18T14:21:00Z"),
                ExcelColumnType.DateTime);
            Assert.IsType<DateTime>(v);
        }

        [Fact]
        public void Coerce_Date_FallsBackToTextOnUnparseable()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("yesterday"), ExcelColumnType.Date);
            Assert.Equal("yesterday", v);
        }

        [Fact]
        public void Coerce_Number_ParsesNumeric()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject(12500.5), ExcelColumnType.Number);
            Assert.IsType<double>(v);
            Assert.Equal(12500.5, (double)v);
        }

        [Fact]
        public void Coerce_Number_ParsesNumericString()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("12500"), ExcelColumnType.Number);
            Assert.IsType<double>(v);
            Assert.Equal(12500.0, (double)v);
        }

        [Fact]
        public void Coerce_Number_FallsBackToTextOnUnparseable()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("twelve"), ExcelColumnType.Number);
            Assert.Equal("twelve", v);
        }

        [Fact]
        public void Coerce_Currency_StripsDollarSignAndCommas()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("$12,500.00"), ExcelColumnType.Currency);
            Assert.IsType<double>(v);
            Assert.Equal(12500.0, (double)v);
        }

        [Fact]
        public void Coerce_Currency_AcceptsRawNumber()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject(12500), ExcelColumnType.Currency);
            Assert.Equal(12500.0, (double)v);
        }

        [Fact]
        public void Coerce_Currency_FallsBackToTextOnExoticSymbol()
        {
            var v = ExcelCellCoercer.Coerce(JToken.FromObject("\u20B91,200"), ExcelColumnType.Currency); // rupee
            Assert.Equal("\u20B91,200", v);
        }

        [Fact]
        public void Coerce_Boolean_ParsesTrueFalse()
        {
            Assert.Equal(true,  ExcelCellCoercer.Coerce(JToken.FromObject(true), ExcelColumnType.Boolean));
            Assert.Equal(false, ExcelCellCoercer.Coerce(JToken.FromObject(false), ExcelColumnType.Boolean));
            Assert.Equal(true,  ExcelCellCoercer.Coerce(JToken.FromObject("true"), ExcelColumnType.Boolean));
            Assert.Equal(false, ExcelCellCoercer.Coerce(JToken.FromObject("FALSE"), ExcelColumnType.Boolean));
        }

        [Fact]
        public void Coerce_NullValue_ReturnsNull()
        {
            Assert.Null(ExcelCellCoercer.Coerce(JValue.CreateNull(), ExcelColumnType.Text));
            Assert.Null(ExcelCellCoercer.Coerce(JValue.CreateNull(), ExcelColumnType.Currency));
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExcelColumnType|FullyQualifiedName~ExcelCellCoercer"
```

Expected: build fails (types not defined).

- [ ] **Step 3: Implement `ExcelColumnType` + parser**

Create `VSTO2/OutlookAI/Services/Export/ExcelColumnType.cs`:

```csharp
using System;

namespace OutlookAI.Services.Export
{
    public enum ExcelColumnType
    {
        Text,
        Date,
        DateTime,
        Number,
        Currency,
        Boolean
    }

    public static class ExcelColumnTypeParser
    {
        public static bool TryParse(string raw, out ExcelColumnType type)
        {
            type = ExcelColumnType.Text;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            switch (raw.Trim().ToLowerInvariant())
            {
                case "text":     type = ExcelColumnType.Text;     return true;
                case "date":     type = ExcelColumnType.Date;     return true;
                case "datetime": type = ExcelColumnType.DateTime; return true;
                case "number":   type = ExcelColumnType.Number;   return true;
                case "currency": type = ExcelColumnType.Currency; return true;
                case "boolean":  type = ExcelColumnType.Boolean;  return true;
                default: return false;
            }
        }
    }
}
```

- [ ] **Step 4: Implement `ExcelCellCoercer`**

Create `VSTO2/OutlookAI/Services/Export/ExcelCellCoercer.cs`:

```csharp
using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Coerces a JSON cell value to a strongly-typed object appropriate for
    /// the column's <see cref="ExcelColumnType"/>. Falls back to text when
    /// the value can't be parsed - never throws.
    /// </summary>
    public static class ExcelCellCoercer
    {
        public static object Coerce(JToken value, ExcelColumnType type)
        {
            if (value == null || value.Type == JTokenType.Null) return null;

            switch (type)
            {
                case ExcelColumnType.Text:     return AsString(value);
                case ExcelColumnType.Date:
                case ExcelColumnType.DateTime: return AsDateTime(value);
                case ExcelColumnType.Number:   return AsNumber(value);
                case ExcelColumnType.Currency: return AsCurrency(value);
                case ExcelColumnType.Boolean:  return AsBoolean(value);
                default:                       return AsString(value);
            }
        }

        private static string AsString(JToken v)
            => v.Type == JTokenType.String ? (string)v : v.ToString(Newtonsoft.Json.Formatting.None);

        private static object AsDateTime(JToken v)
        {
            var s = AsString(v);
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var d))
            {
                return d;
            }
            return s;
        }

        private static object AsNumber(JToken v)
        {
            if (v.Type == JTokenType.Integer || v.Type == JTokenType.Float)
                return (double)v;
            var s = AsString(v);
            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture, out var n))
            {
                return n;
            }
            return s;
        }

        private static object AsCurrency(JToken v)
        {
            if (v.Type == JTokenType.Integer || v.Type == JTokenType.Float)
                return (double)v;
            var s = AsString(v);
            // Strip $ and commas (US-locale-ish; other locales fall back to text).
            var stripped = s.Replace("$", "").Replace(",", "").Trim();
            if (double.TryParse(stripped, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var n))
            {
                return n;
            }
            return s;
        }

        private static object AsBoolean(JToken v)
        {
            if (v.Type == JTokenType.Boolean) return (bool)v;
            var s = AsString(v).Trim().ToLowerInvariant();
            if (s == "true")  return true;
            if (s == "false") return false;
            return s;
        }
    }
}
```

Register in `OutlookAI.csproj`:

```xml
<Compile Include="Services\Export\ExcelColumnType.cs" />
<Compile Include="Services\Export\ExcelCellCoercer.cs" />
```

Register tests in `OutlookAI.Tests.csproj`:

```xml
<Compile Include="Services\Export\ExcelColumnTypeTests.cs" />
<Compile Include="Services\Export\ExcelCellCoercerTests.cs" />
```

- [ ] **Step 5: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExcelColumnType|FullyQualifiedName~ExcelCellCoercer"
```

Expected: ~22 tests pass.

- [ ] **Step 6: Full suite**

Expected total: 359 (337 + 22).

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/ExcelColumnType.cs VSTO2/OutlookAI/Services/Export/ExcelCellCoercer.cs VSTO2/OutlookAI.Tests/Services/Export/ExcelColumnTypeTests.cs VSTO2/OutlookAI.Tests/Services/Export/ExcelCellCoercerTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add Excel column type coercion"
```

---

## Task 5: `ExcelWorkbookBuilder` — pure builder

Given a sheet name + columns + rows, produce an `XLWorkbook` with: bold + gray-filled + frozen header row, autofilter over the header range, columns auto-sized to content, currency cells `$#,##0.00`, date cells `yyyy-mm-dd`, datetime cells `yyyy-mm-dd hh:mm`. Pure builder — no I/O.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/ExcelColumnSpec.cs` (small DTO used by builder + args).
- Create: `VSTO2/OutlookAI/Services/Export/ExcelWorkbookBuilder.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/ExcelWorkbookBuilderTests.cs`

- [ ] **Step 1: Create the `ExcelColumnSpec` DTO**

Create `VSTO2/OutlookAI/Services/Export/ExcelColumnSpec.cs`:

```csharp
namespace OutlookAI.Services.Export
{
    public sealed class ExcelColumnSpec
    {
        public string Name { get; set; }
        public ExcelColumnType Type { get; set; }
    }
}
```

Register in `OutlookAI.csproj`:

```xml
<Compile Include="Services\Export\ExcelColumnSpec.cs" />
```

- [ ] **Step 2: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Export/ExcelWorkbookBuilderTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelWorkbookBuilderTests
    {
        private static IList<ExcelColumnSpec> SampleColumns() => new[]
        {
            new ExcelColumnSpec { Name = "Date",         Type = ExcelColumnType.Date     },
            new ExcelColumnSpec { Name = "Subject",      Type = ExcelColumnType.Text     },
            new ExcelColumnSpec { Name = "Sender",       Type = ExcelColumnType.Text     },
            new ExcelColumnSpec { Name = "Quoted Total", Type = ExcelColumnType.Currency },
        };

        private static IList<object[]> SampleRows() => new[]
        {
            new object[] { new DateTime(2026, 5, 12), "RE: Quote for Dell R750", "Murad Lalaiev", 12500.00 },
            new object[] { new DateTime(2026, 5, 14), "RE: HPE DL380 quote",      "Murad Lalaiev",  8200.50 },
        };

        [Fact]
        public void Build_SheetNameApplied()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Quotes", SampleColumns(), SampleRows()))
            {
                Assert.Equal("Quotes", wb.Worksheets.First().Name);
            }
        }

        [Fact]
        public void Build_HeaderRowMatchesColumnNames()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                Assert.Equal("Date",         ws.Cell(1, 1).GetString());
                Assert.Equal("Subject",      ws.Cell(1, 2).GetString());
                Assert.Equal("Sender",       ws.Cell(1, 3).GetString());
                Assert.Equal("Quoted Total", ws.Cell(1, 4).GetString());
            }
        }

        [Fact]
        public void Build_HeaderRowIsBoldAndFilled()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                var headerStyle = ws.Cell(1, 1).Style;
                Assert.True(headerStyle.Font.Bold);
                Assert.NotEqual(XLColor.NoColor, headerStyle.Fill.BackgroundColor);
            }
        }

        [Fact]
        public void Build_FreezesTopRow()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                Assert.Equal(1, ws.SheetView.SplitRow);
            }
        }

        [Fact]
        public void Build_AppliesAutoFilterOverHeaderRange()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                Assert.True(ws.AutoFilter.IsEnabled);
            }
        }

        [Fact]
        public void Build_DateColumnUsesIsoFormat()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                Assert.Equal("yyyy-mm-dd", ws.Cell(2, 1).Style.DateFormat.Format);
            }
        }

        [Fact]
        public void Build_CurrencyColumnUsesDollarFormat()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                Assert.Contains("$", ws.Cell(2, 4).Style.NumberFormat.Format);
            }
        }

        [Fact]
        public void Build_RowCount_MatchesInput()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), SampleRows()))
            {
                var ws = wb.Worksheets.First();
                // header + 2 data rows
                Assert.Equal(3, ws.LastRowUsed().RowNumber());
            }
        }

        [Fact]
        public void Build_EmptyRows_HeaderOnly()
        {
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), new List<object[]>()))
            {
                var ws = wb.Worksheets.First();
                Assert.Equal(1, ws.LastRowUsed().RowNumber());
            }
        }

        [Fact]
        public void Build_NullCellRendersAsEmpty()
        {
            var rows = new List<object[]>
            {
                new object[] { new DateTime(2026, 5, 12), null, "Sender", 100.0 }
            };
            using (var wb = ExcelWorkbookBuilder.Build("Sheet1", SampleColumns(), rows))
            {
                var ws = wb.Worksheets.First();
                Assert.True(ws.Cell(2, 2).IsEmpty());
            }
        }
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExcelWorkbookBuilderTests"
```

Expected: build fails (`ExcelWorkbookBuilder` not defined).

- [ ] **Step 4: Implement `ExcelWorkbookBuilder`**

Create `VSTO2/OutlookAI/Services/Export/ExcelWorkbookBuilder.cs`:

```csharp
using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Pure builder. Returns an in-memory <see cref="XLWorkbook"/> with the
    /// supplied data laid out as a single sheet with bold/frozen/filtered
    /// header row, type-aware cell formatting, and content-fit columns.
    /// </summary>
    public static class ExcelWorkbookBuilder
    {
        public static XLWorkbook Build(string sheetName, IList<ExcelColumnSpec> columns, IList<object[]> rows)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet(sheetName);

            // Header row.
            for (int c = 0; c < columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = columns[c].Name ?? "";
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F4F6");
            }
            ws.SheetView.FreezeRows(1);

            // Data rows.
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (int c = 0; c < columns.Count; c++)
                {
                    if (c >= row.Length) continue;
                    var v = row[c];
                    var cell = ws.Cell(r + 2, c + 1);

                    if (v == null) continue;

                    var type = columns[c].Type;
                    switch (type)
                    {
                        case ExcelColumnType.Date:
                            cell.Value = v;
                            cell.Style.DateFormat.Format = "yyyy-mm-dd";
                            break;
                        case ExcelColumnType.DateTime:
                            cell.Value = v;
                            cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm";
                            break;
                        case ExcelColumnType.Number:
                            cell.Value = v;
                            break;
                        case ExcelColumnType.Currency:
                            cell.Value = v;
                            cell.Style.NumberFormat.Format = "$#,##0.00";
                            break;
                        case ExcelColumnType.Boolean:
                            cell.Value = v;
                            break;
                        case ExcelColumnType.Text:
                        default:
                            cell.Value = v;
                            break;
                    }
                }
            }

            // Autofilter over the header range only when we have header columns.
            if (columns.Count > 0)
            {
                int lastRow = Math.Max(1, rows.Count + 1);
                ws.Range(1, 1, lastRow, columns.Count).SetAutoFilter();
            }

            ws.Columns().AdjustToContents();
            return wb;
        }
    }
}
```

Register in `OutlookAI.csproj`:

```xml
<Compile Include="Services\Export\ExcelWorkbookBuilder.cs" />
```

Register tests in `OutlookAI.Tests.csproj`:

```xml
<Compile Include="Services\Export\ExcelWorkbookBuilderTests.cs" />
```

- [ ] **Step 5: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExcelWorkbookBuilderTests"
```

Expected: 10 tests pass.

- [ ] **Step 6: Full suite**

Expected total: 369 (359 + 10).

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/ExcelColumnSpec.cs VSTO2/OutlookAI/Services/Export/ExcelWorkbookBuilder.cs VSTO2/OutlookAI.Tests/Services/Export/ExcelWorkbookBuilderTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add ExcelWorkbookBuilder with frozen header and autofilter"
```

---

## Task 6: `ExportExcelArgs` + `ExportExcelArgsParser`

JSON parser for the tool's args envelope. Validates that columns is non-empty, types are valid, rows ≤ 10 000, row shapes match column count.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/ExportExcelArgs.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/ExportExcelArgsParser.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/ToolArgValidationException.cs` (if doesn't exist; check first)
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/ExportExcelArgsParserTests.cs`

- [ ] **Step 1: Check whether `ToolArgValidationException` exists**

```powershell
Get-ChildItem -Path "VSTO2\OutlookAI\Services\Tools" -Filter "*Validation*.cs"
```

If it exists, reuse it. If not, create:

```csharp
// VSTO2/OutlookAI/Services/Tools/ToolArgValidationException.cs
using System;

namespace OutlookAI.Services.Tools
{
    public sealed class ToolArgValidationException : Exception
    {
        public string Code { get; }
        public ToolArgValidationException(string code, string detail) : base(detail)
        {
            Code = code;
        }
    }
}
```

(Other tools may use a different error pattern; check `OutlookSearchMessagesTool` and `SearchMessagesArgsParser` for the existing convention before deviating. If they use return-value tuples or a different exception, follow their lead.)

- [ ] **Step 2: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/ExportExcelArgsParserTests.cs`:

```csharp
using System.Collections.Generic;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportExcelArgsParserTests
    {
        private const string ValidArgs = @"{
            ""filename_hint"":""IT Creations Quotes"",
            ""sheet_name"":""Quotes"",
            ""columns"":[
                {""name"":""Date"",""type"":""date""},
                {""name"":""Subject"",""type"":""text""},
                {""name"":""Sender"",""type"":""text""},
                {""name"":""Quoted Total"",""type"":""currency""}
            ],
            ""rows"":[
                [""2026-05-12"",""Quote 1"",""Murad"",12500],
                [""2026-05-14"",""Quote 2"",""Murad"",8200.50]
            ]
        }";

        [Fact]
        public void Parses_ValidArgs()
        {
            var args = ExportExcelArgsParser.Parse(ValidArgs);
            Assert.Equal("IT Creations Quotes", args.FilenameHint);
            Assert.Equal("Quotes", args.SheetName);
            Assert.Equal(4, args.Columns.Count);
            Assert.Equal("Date", args.Columns[0].Name);
            Assert.Equal(ExcelColumnType.Date, args.Columns[0].Type);
            Assert.Equal(2, args.Rows.Count);
        }

        [Fact]
        public void SheetName_FallsBackToFilenameHintWhenMissing()
        {
            var json = @"{""filename_hint"":""Quotes"",""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[[""x""]]}";
            var args = ExportExcelArgsParser.Parse(json);
            Assert.Equal("Quotes", args.SheetName);
        }

        [Fact]
        public void Throws_WhenColumnsMissing()
        {
            var json = @"{""filename_hint"":""x"",""rows"":[[""a""]]}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));
            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("columns", ex.Message);
        }

        [Fact]
        public void Throws_WhenColumnsEmpty()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[],""rows"":[]}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));
            Assert.Equal("invalid_args", ex.Code);
        }

        [Fact]
        public void Throws_WhenColumnTypeUnknown()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[{""name"":""A"",""type"":""money""}],""rows"":[]}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));
            Assert.Contains("type 'money'", ex.Message);
        }

        [Fact]
        public void Throws_WhenColumnMissingName()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[{""type"":""text""}],""rows"":[]}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));
            Assert.Contains("name", ex.Message);
        }

        [Fact]
        public void Throws_WhenRowsExceedHardCap()
        {
            // build a JSON with 10001 rows
            var sb = new System.Text.StringBuilder(@"{""filename_hint"":""x"",""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[");
            for (int i = 0; i < 10001; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(@"[""x""]");
            }
            sb.Append("]}");
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(sb.ToString()));
            Assert.Equal("too_many_rows", ex.Code);
        }

        [Fact]
        public void Throws_WhenRowShapeMismatches()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[{""name"":""A"",""type"":""text""},{""name"":""B"",""type"":""text""}],""rows"":[[""only-one-cell""]]}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));
            Assert.Equal("row_shape_mismatch", ex.Code);
            Assert.Contains("row 0", ex.Message);
        }

        [Fact]
        public void EmptyFilenameHint_FallsBackToDefault()
        {
            var json = @"{""filename_hint"":"""",""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[]}";
            var args = ExportExcelArgsParser.Parse(json);
            Assert.Equal("OutlookAI-Report", args.FilenameHint);
        }

        [Fact]
        public void MissingFilenameHint_FallsBackToDefault()
        {
            var json = @"{""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[]}";
            var args = ExportExcelArgsParser.Parse(json);
            Assert.Equal("OutlookAI-Report", args.FilenameHint);
        }

        [Fact]
        public void RowsCanBeEmpty()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[]}";
            var args = ExportExcelArgsParser.Parse(json);
            Assert.Empty(args.Rows);
        }

        [Fact]
        public void RowsCellsArePreservedAsJTokens()
        {
            var json = @"{""filename_hint"":""x"",""columns"":[{""name"":""N"",""type"":""number""}],""rows"":[[42]]}";
            var args = ExportExcelArgsParser.Parse(json);
            // JToken with Integer type
            Assert.Equal(Newtonsoft.Json.Linq.JTokenType.Integer, args.Rows[0][0].Type);
        }
    }
}
```

- [ ] **Step 3: Implement `ExportExcelArgs`**

Create `VSTO2/OutlookAI/Services/Tools/ExportExcelArgs.cs`:

```csharp
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public sealed class ExportExcelArgs
    {
        public string FilenameHint { get; set; }
        public string SheetName { get; set; }
        public IList<ExcelColumnSpec> Columns { get; set; }
        /// <summary>Rows of JToken cells; types preserved for coercion.</summary>
        public IList<JToken[]> Rows { get; set; }
    }
}
```

- [ ] **Step 4: Implement `ExportExcelArgsParser`**

Create `VSTO2/OutlookAI/Services/Tools/ExportExcelArgsParser.cs`:

```csharp
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public static class ExportExcelArgsParser
    {
        private const int MaxRows = 10000;
        private const string DefaultFilenameHint = "OutlookAI-Report";

        public static ExportExcelArgs Parse(string json)
        {
            JObject root;
            try { root = JObject.Parse(json ?? ""); }
            catch (System.Exception ex)
            {
                throw new ToolArgValidationException("invalid_args", "could not parse JSON: " + ex.Message);
            }

            var hintRaw = (string)root["filename_hint"];
            var filenameHint = string.IsNullOrWhiteSpace(hintRaw) ? DefaultFilenameHint : hintRaw.Trim();

            var sheetName = (string)root["sheet_name"];
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = filenameHint;

            var colsToken = root["columns"] as JArray;
            if (colsToken == null || colsToken.Count == 0)
                throw new ToolArgValidationException("invalid_args", "columns must contain at least one column");

            var columns = new List<ExcelColumnSpec>(colsToken.Count);
            for (int i = 0; i < colsToken.Count; i++)
            {
                var col = colsToken[i] as JObject;
                if (col == null)
                    throw new ToolArgValidationException("invalid_args", $"column {i}: must be an object");

                var name = (string)col["name"];
                if (string.IsNullOrWhiteSpace(name))
                    throw new ToolArgValidationException("invalid_args", $"column {i}: 'name' is required");

                var typeRaw = (string)col["type"];
                if (!ExcelColumnTypeParser.TryParse(typeRaw, out var type))
                    throw new ToolArgValidationException("invalid_args",
                        $"column '{name}': type '{typeRaw}' is not supported. Use one of: text, date, datetime, number, currency, boolean");

                columns.Add(new ExcelColumnSpec { Name = name.Trim(), Type = type });
            }

            var rowsToken = root["rows"] as JArray;
            var rows = new List<JToken[]>(rowsToken?.Count ?? 0);
            if (rowsToken != null)
            {
                if (rowsToken.Count > MaxRows)
                    throw new ToolArgValidationException("too_many_rows",
                        $"{MaxRows} row limit; got {rowsToken.Count}. Aggregate or filter before exporting.");

                for (int r = 0; r < rowsToken.Count; r++)
                {
                    var rowArr = rowsToken[r] as JArray;
                    if (rowArr == null)
                        throw new ToolArgValidationException("invalid_args", $"row {r}: must be an array");
                    if (rowArr.Count != columns.Count)
                        throw new ToolArgValidationException("row_shape_mismatch",
                            $"row {r} has {rowArr.Count} cells, expected {columns.Count}");

                    var cells = new JToken[rowArr.Count];
                    for (int c = 0; c < rowArr.Count; c++) cells[c] = rowArr[c];
                    rows.Add(cells);
                }
            }

            return new ExportExcelArgs
            {
                FilenameHint = filenameHint,
                SheetName = sheetName,
                Columns = columns,
                Rows = rows
            };
        }
    }
}
```

Register both in `OutlookAI.csproj`:

```xml
<Compile Include="Services\Tools\ExportExcelArgs.cs" />
<Compile Include="Services\Tools\ExportExcelArgsParser.cs" />
<Compile Include="Services\Tools\ToolArgValidationException.cs" />
```

Register test in `OutlookAI.Tests.csproj`:

```xml
<Compile Include="Services\Tools\ExportExcelArgsParserTests.cs" />
```

- [ ] **Step 5: Run tests + full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 381` (369 + 12).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ExportExcelArgs.cs VSTO2/OutlookAI/Services/Tools/ExportExcelArgsParser.cs VSTO2/OutlookAI/Services/Tools/ToolArgValidationException.cs VSTO2/OutlookAI.Tests/Services/Tools/ExportExcelArgsParserTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add Excel args parser with strict validation"
```

---

## Task 7: `FileSavedResult` + `IOutlookSurface.ExportExcel` + `LiveOutlookSurface.ExportExcel`

The surface method that actually writes the workbook to disk. Validated via manual smoke (we don't unit-test real I/O against the live Outlook surface). Adds the `FileSavedResult` DTO returned by both export methods.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/FileSavedResult.cs`
- Create: `VSTO2/OutlookAI/Services/Export/ExportException.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` (NullSurface override)

- [ ] **Step 1: Create `FileSavedResult` DTO**

Create `VSTO2/OutlookAI/Services/Export/FileSavedResult.cs`:

```csharp
namespace OutlookAI.Services.Export
{
    public sealed class FileSavedResult
    {
        public string Path { get; set; }
        public string FileUrl { get; set; }
        /// <summary>"xlsx" or "pdf".</summary>
        public string Format { get; set; }
        public long Bytes { get; set; }
        public string Filename { get; set; }
    }
}
```

- [ ] **Step 2: Create `ExportException`**

Create `VSTO2/OutlookAI/Services/Export/ExportException.cs`:

```csharp
using System;

namespace OutlookAI.Services.Export
{
    public sealed class ExportException : Exception
    {
        public string Code { get; }
        public ExportException(string code, string detail, Exception inner = null) : base(detail, inner)
        {
            Code = code;
        }
    }
}
```

- [ ] **Step 3: Add the surface signature**

Modify `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` — inside the `IOutlookSurface` interface, after the existing `AggregateMessages` line, add:

```csharp
FileSavedResult ExportExcel(OutlookAI.Services.Tools.ExportExcelArgs args, CancellationToken ct = default(CancellationToken));
FileSavedResult ExportPdf(OutlookAI.Services.Tools.ExportPdfArgs args, CancellationToken ct = default(CancellationToken));
```

(Add `using OutlookAI.Services.Export;` at the top of the file.)

Note: `ExportPdfArgs` is created in a later task. To keep this task compilable, you can stub it briefly or skip the PDF signature until Task 15. Recommended: add a placeholder `ExportPdfArgs` now (empty class) so all signatures exist; flesh it out in Task 15.

Create the stub `VSTO2/OutlookAI/Services/Tools/ExportPdfArgs.cs`:

```csharp
namespace OutlookAI.Services.Tools
{
    public sealed class ExportPdfArgs
    {
        public string FilenameHint { get; set; }
        public string ContentMarkdown { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
    }
}
```

Register in csproj:

```xml
<Compile Include="Services\Export\FileSavedResult.cs" />
<Compile Include="Services\Export\ExportException.cs" />
<Compile Include="Services\Tools\ExportPdfArgs.cs" />
```

- [ ] **Step 4: Implement `LiveOutlookSurface.ExportExcel`**

At the bottom of `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` (inside the class), add (and add `using OutlookAI.Services.Export;` and `using ClosedXML.Excel;` at top):

```csharp
public FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
{
    if (args == null) throw new System.ArgumentNullException(nameof(args));

    ct.ThrowIfCancellationRequested();

    var resolver = new ExportPathResolver();
    string baseDir;
    try
    {
        resolver.EnsureExists();
        baseDir = resolver.ResolveBaseDir();
    }
    catch (System.IO.IOException ex)
    {
        throw new ExportException("path_unavailable", ex.Message, ex);
    }

    var filename = ExportFilenameSanitizer.Build(
        args.FilenameHint, ".xlsx", System.DateTimeOffset.Now,
        candidate => System.IO.File.Exists(System.IO.Path.Combine(baseDir, candidate)));

    var fullPath = System.IO.Path.Combine(baseDir, filename);

    // Materialize rows into typed cells.
    var typedRows = new System.Collections.Generic.List<object[]>(args.Rows.Count);
    for (int r = 0; r < args.Rows.Count; r++)
    {
        ct.ThrowIfCancellationRequested();
        var row = new object[args.Columns.Count];
        for (int c = 0; c < args.Columns.Count; c++)
        {
            row[c] = ExcelCellCoercer.Coerce(args.Rows[r][c], args.Columns[c].Type);
        }
        typedRows.Add(row);
    }

    try
    {
        using (var wb = ExcelWorkbookBuilder.Build(args.SheetName, args.Columns, typedRows))
        {
            try { wb.SaveAs(fullPath); }
            catch (System.IO.IOException ioex) when ((ioex.HResult & 0xFFFF) == 32 /* ERROR_SHARING_VIOLATION */)
            {
                // Retry once with collision suffix.
                var retryName = ExportFilenameSanitizer.Build(
                    args.FilenameHint + "-2", ".xlsx", System.DateTimeOffset.Now,
                    candidate => System.IO.File.Exists(System.IO.Path.Combine(baseDir, candidate)));
                fullPath = System.IO.Path.Combine(baseDir, retryName);
                filename = retryName;
                try { wb.SaveAs(fullPath); }
                catch (System.IO.IOException retryEx)
                {
                    throw new ExportException("file_locked", retryEx.Message, retryEx);
                }
            }
            catch (System.IO.IOException diskEx) when (diskEx.Message.Contains("space"))
            {
                throw new ExportException("disk_full", diskEx.Message, diskEx);
            }
        }
    }
    catch (ExportException) { throw; }
    catch (System.Exception ex)
    {
        throw new ExportException("excel_build_failed", ex.Message, ex);
    }

    var bytes = new System.IO.FileInfo(fullPath).Length;
    return new FileSavedResult
    {
        Path = fullPath,
        FileUrl = new System.Uri(fullPath).AbsoluteUri,
        Format = "xlsx",
        Bytes = bytes,
        Filename = filename
    };
}

// PDF surface method - implemented in Task 16. Stub so interface compiles.
public FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
{
    throw new System.NotImplementedException("ExportPdf is wired in Task 16.");
}
```

- [ ] **Step 5: Update `MinimalSurface`**

Modify `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs` — add (after the existing virtual methods):

```csharp
public virtual FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
    => throw new System.NotImplementedException();

public virtual FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
    => throw new System.NotImplementedException();
```

(Add `using OutlookAI.Services.Export;` if not already there.)

- [ ] **Step 6: Update `NullSurface` in `AITaskPane.cs`**

Modify `VSTO2/OutlookAI/TaskPane/AITaskPane.cs` — find the `NullSurface` (or equivalent stub) and add the same two methods returning `null` or throwing `NotImplementedException` per the existing pattern.

- [ ] **Step 7: Build to verify compile**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds with no errors. The methods are reachable through the interface.

- [ ] **Step 8: Run full suite to confirm no regressions**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 381` (no new tests in this task; surface is integration-only).

- [ ] **Step 9: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/FileSavedResult.cs VSTO2/OutlookAI/Services/Export/ExportException.cs VSTO2/OutlookAI/Services/Tools/ExportPdfArgs.cs VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs VSTO2/OutlookAI/TaskPane/AITaskPane.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(export): wire LiveOutlookSurface.ExportExcel"
```

---

## Task 8: `OutlookExportExcelTool`

`IOutlookTool` wrapper that parses args, dispatches to surface, and returns the file-saved envelope or an error envelope.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookExportExcelTool.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportExcelToolTests.cs`

- [ ] **Step 1: Inspect an existing tool to match the convention**

Look at `OutlookSearchMessagesTool.cs`:

```powershell
Get-Content -LiteralPath "VSTO2\OutlookAI\Services\Tools\OutlookSearchMessagesTool.cs" | Select-Object -First 60
```

The pattern likely has `Name`, `DispatchAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)`, and returns a JSON string envelope. Match that.

- [ ] **Step 2: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportExcelToolTests.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportExcelToolTests
    {
        private sealed class FakeSurface : MinimalSurface
        {
            public ExportExcelArgs Captured;
            public FileSavedResult ResultToReturn;
            public ExportException ToThrow;
            public bool CancelObserved;

            public override FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
            {
                Captured = args;
                if (ct.IsCancellationRequested) { CancelObserved = true; throw new OperationCanceledException(ct); }
                if (ToThrow != null) throw ToThrow;
                return ResultToReturn;
            }
        }

        private const string ValidArgs = @"{
            ""filename_hint"":""Quotes"",
            ""columns"":[{""name"":""A"",""type"":""text""}],
            ""rows"":[[""x""]]
        }";

        [Fact]
        public async Task ReturnsFileSavedEnvelopeOnSuccess()
        {
            var surface = new FakeSurface
            {
                ResultToReturn = new FileSavedResult
                {
                    Path = @"C:\Reports\Quotes-2026-05-18-1947.xlsx",
                    FileUrl = "file:///C:/Reports/Quotes-2026-05-18-1947.xlsx",
                    Format = "xlsx",
                    Bytes = 12345,
                    Filename = "Quotes-2026-05-18-1947.xlsx"
                }
            };
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync(ValidArgs, surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.Equal("xlsx", (string)obj["format"]);
            Assert.Equal("Quotes-2026-05-18-1947.xlsx", (string)obj["filename"]);
            Assert.Equal(12345L, (long)obj["bytes"]);
        }

        [Fact]
        public async Task ReturnsErrorEnvelopeOnValidationFailure()
        {
            var surface = new FakeSurface();
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync(@"{""rows"":[]}", surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("invalid_args", (string)obj["error"]);
        }

        [Fact]
        public async Task ReturnsErrorEnvelopeOnExportException()
        {
            var surface = new FakeSurface
            {
                ToThrow = new ExportException("file_locked", "in use by Excel")
            };
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync(ValidArgs, surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("file_locked", (string)obj["error"]);
            Assert.Contains("Excel", (string)obj["detail"]);
        }

        [Fact]
        public async Task PropagatesCancellation()
        {
            var surface = new FakeSurface();
            var cts = new CancellationTokenSource();
            cts.Cancel();
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync(ValidArgs, surface, cts.Token);
            var obj = JObject.Parse(json);
            Assert.Equal("cancelled", (string)obj["error"]);
        }

        [Fact]
        public async Task NameIs_outlook_export_excel()
        {
            Assert.Equal("outlook_export_excel", new OutlookExportExcelTool().Name);
            await Task.CompletedTask;
        }

        [Fact]
        public async Task CapturesParsedArgsOnSurface()
        {
            var surface = new FakeSurface
            {
                ResultToReturn = new FileSavedResult { Path = "x", Filename = "x", Format = "xlsx", Bytes = 0, FileUrl = "file:///x" }
            };
            var tool = new OutlookExportExcelTool();
            await tool.DispatchAsync(ValidArgs, surface, CancellationToken.None);
            Assert.NotNull(surface.Captured);
            Assert.Equal("Quotes", surface.Captured.FilenameHint);
            Assert.Single(surface.Captured.Columns);
        }

        [Fact]
        public async Task TooManyRows_ReturnsTooManyRowsCode()
        {
            var surface = new FakeSurface();
            // Build 10001-row payload programmatically
            var sb = new System.Text.StringBuilder(@"{""filename_hint"":""x"",""columns"":[{""name"":""A"",""type"":""text""}],""rows"":[");
            for (int i = 0; i < 10001; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(@"[""x""]");
            }
            sb.Append("]}");
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync(sb.ToString(), surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("too_many_rows", (string)obj["error"]);
        }

        [Fact]
        public async Task ParseFailure_ReturnsInvalidArgs()
        {
            var surface = new FakeSurface();
            var tool = new OutlookExportExcelTool();
            var json = await tool.DispatchAsync("not-json", surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("invalid_args", (string)obj["error"]);
        }
    }
}
```

- [ ] **Step 3: Implement `OutlookExportExcelTool`**

Create `VSTO2/OutlookAI/Services/Tools/OutlookExportExcelTool.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public sealed class OutlookExportExcelTool : IOutlookTool
    {
        public string Name => "outlook_export_excel";

        public Task<string> DispatchAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ExportExcelArgs args;
            try
            {
                args = ExportExcelArgsParser.Parse(argsJson);
            }
            catch (ToolArgValidationException vex)
            {
                return Task.FromResult(ErrorEnvelope(vex.Code, vex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(ErrorEnvelope("invalid_args", ex.Message));
            }

            try
            {
                if (ct.IsCancellationRequested)
                    return Task.FromResult(ErrorEnvelope("cancelled", "tool cancelled before dispatch"));

                var result = surface.ExportExcel(args, ct);
                return Task.FromResult(SuccessEnvelope(result));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(ErrorEnvelope("cancelled", "tool cancelled"));
            }
            catch (ExportException eex)
            {
                return Task.FromResult(ErrorEnvelope(eex.Code, eex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(ErrorEnvelope("excel_build_failed", ex.Message));
            }
        }

        private static string SuccessEnvelope(FileSavedResult r)
        {
            var o = new JObject(
                new JProperty("result_type", "file_saved"),
                new JProperty("path", r.Path),
                new JProperty("file_url", r.FileUrl),
                new JProperty("format", r.Format),
                new JProperty("bytes", r.Bytes),
                new JProperty("filename", r.Filename));
            return o.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string ErrorEnvelope(string code, string detail)
        {
            var o = new JObject(
                new JProperty("error", code),
                new JProperty("detail", detail ?? ""));
            return o.ToString(Newtonsoft.Json.Formatting.None);
        }
    }
}
```

(If `IOutlookTool` is shaped differently — e.g., a synchronous method or different signature — adjust to match `OutlookSearchMessagesTool.cs`. The pattern above is the most common shape.)

Register in csproj:

```xml
<Compile Include="Services\Tools\OutlookExportExcelTool.cs" />
```

Test:

```xml
<Compile Include="Services\Tools\OutlookExportExcelToolTests.cs" />
```

- [ ] **Step 4: Run tests + full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 389` (381 + 8).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookExportExcelTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportExcelToolTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(tools): add outlook_export_excel tool"
```

---

## Task 9: Add `outlook_export_excel` schema

Adds the tool to `ToolCatalogSchema` so the model knows it exists, what to pass, and when to choose it. Strong steering description matters — the model decides when "make a spreadsheet" vs "make a PDF" applies.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`

- [ ] **Step 1: Write the failing tests**

Append to `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void OutlookExportExcel_IsRegistered()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    Assert.Contains(arr, t => (string)t["name"] == "outlook_export_excel");
}

[Fact]
public void OutlookExportExcel_Description_TeachesWhenToUseExcelVsPdf()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_excel");
    var desc = ((string)entry["description"]) ?? "";
    Assert.Contains("spreadsheet", desc, System.StringComparison.OrdinalIgnoreCase);
    Assert.Contains("Excel", desc);
    Assert.Contains("columns", desc);
    Assert.Contains("rows", desc);
}

[Fact]
public void OutlookExportExcel_Schema_HasColumnsAndRows()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_excel");
    var props = entry["parameters"]?["properties"];
    Assert.NotNull(props["filename_hint"]);
    Assert.NotNull(props["columns"]);
    Assert.NotNull(props["rows"]);
}

[Fact]
public void OutlookExportExcel_TeachesWhenNotToUse()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_excel");
    var desc = ((string)entry["description"]) ?? "";
    // Should warn against using for prose/narrative.
    Assert.Contains("prose", desc, System.StringComparison.OrdinalIgnoreCase);
}
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookExportExcel"
```

Expected: 4 tests fail (entry not in catalog).

- [ ] **Step 3: Add the schema entry**

In `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`, locate the `arr` initializer inside `BuildResponsesToolsArray`. Add this entry (after `outlook_aggregate_messages` for grouping with related tools):

```csharp
BuildToolEntry("outlook_export_excel",
    "Save a structured table to an Excel (.xlsx) file. Use when the user asks for a spreadsheet, Excel, xlsx, or any tabular export. "
    + "Pass typed columns and rows; the tool produces a styled workbook with bold/frozen header, autofilter, and per-column formatting (dates, currency, etc.). "
    + "Best for: vendor lists, message tables (date/subject/sender), aggregations broken down by sender or day, structured exports of search results. "
    + "Do NOT use for prose, narrative reports, or arbitrary text - choose outlook_export_pdf for those. "
    + "Maximum 10000 rows. If you have more data, aggregate or filter before calling. "
    + "Example: user says 'give me an Excel of quotes from IT Creations' -> first call outlook_search_messages, then call this tool with columns=[{name:'Date',type:'date'},{name:'Subject',type:'text'},{name:'Sender',type:'text'},{name:'Quoted Total',type:'currency'}] and rows projected from the search hits. "
    + "File is saved to ~\\Documents\\OutlookAI\\Reports\\ with an auto-generated timestamped filename. After success, mention the filename briefly; the UI surfaces an Open / Show-in-folder card automatically.",
    new JObject(
        new JProperty("type", "object"),
        new JProperty("properties", new JObject(
            new JProperty("filename_hint", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Short human-readable hint for the filename (e.g. 'IT Creations Quotes'). Tool sanitizes invalid characters and appends a timestamp. Optional - defaults to 'OutlookAI-Report'."))),
            new JProperty("sheet_name", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Worksheet tab name. Optional - defaults to filename_hint."))),
            new JProperty("columns", new JObject(
                new JProperty("type", "array"),
                new JProperty("description", "Column definitions. Order matches rows."),
                new JProperty("items", new JObject(
                    new JProperty("type", "object"),
                    new JProperty("properties", new JObject(
                        new JProperty("name", new JObject(
                            new JProperty("type", "string"),
                            new JProperty("description", "Column header text shown in row 1."))),
                        new JProperty("type", new JObject(
                            new JProperty("type", "string"),
                            new JProperty("enum", new JArray("text", "date", "datetime", "number", "currency", "boolean")),
                            new JProperty("description", "Cell type for formatting. 'date' uses yyyy-mm-dd format; 'currency' formats as $#,##0.00."))))),
                    new JProperty("required", new JArray("name", "type")))))),
            new JProperty("rows", new JObject(
                new JProperty("type", "array"),
                new JProperty("description", "Array of row arrays. Each row must have exactly columns.length cells. Cells may be strings, numbers, ISO-8601 dates, or booleans matching the column type."),
                new JProperty("items", new JObject(
                    new JProperty("type", "array"))))))),
        new JProperty("required", new JArray("columns", "rows")),
        new JProperty("additionalProperties", false))),
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookExportExcel"
```

Expected: 4 schema tests pass.

- [ ] **Step 5: Full suite**

Expected: 393 (389 + 4).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(schema): teach model when to use Excel export"
```

---

## Task 10: Register `OutlookExportExcelTool` in `OutlookToolHost`

Adds the tool instance to the read-tools list so the dispatcher actually routes to it.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/OutlookToolHost.cs`

- [ ] **Step 1: Add the tool to the list**

In `VSTO2/OutlookAI/Services/OutlookToolHost.cs`, find the `tools` initializer in the constructor. Append after `OutlookListRecentThreadsWithTool()`:

```csharp
new OutlookExportExcelTool(),           // Phase 5: Excel export
```

- [ ] **Step 2: Build**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds.

- [ ] **Step 3: Manual sanity check**

The tool is now reachable from `IToolHost.DispatchAsync("outlook_export_excel", argsJson, ct)`. We'll exercise it end-to-end during the smoke phase. For now, ensure the catalog and host are aligned.

- [ ] **Step 4: Full suite**

Expected: 393 (no test count change).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs
git commit -m "feat(host): register Excel export tool"
```

---

## Task 11: Factor markdown renderer out of `chat.js` into `WebUI/markdown.js`

The PDF print template needs the same markdown-to-HTML converter the chat uses (ATX headings, GFM tables, blockquotes, `<br>`, code blocks). Extract the renderer into a standalone module loaded by both `chat.html` and `print-template.html`.

**Files:**
- Read existing: `VSTO2/OutlookAI/WebUI/chat.js`
- Create: `VSTO2/OutlookAI/WebUI/markdown.js`
- Modify: `VSTO2/OutlookAI/WebUI/chat.js` (delegate to markdown.js)
- Modify: `VSTO2/OutlookAI/WebUI/index.html` (load markdown.js before chat.js)
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj` (register markdown.js as embedded resource)

- [ ] **Step 1: Locate the existing markdown renderer in chat.js**

```powershell
Select-String -Path "VSTO2\OutlookAI\WebUI\chat.js" -Pattern "renderMarkdown|markdown|h1|h2|tableRegex|gfm" | Select-Object -First 30
```

Expected: finds an inline `renderMarkdown(src)` function (probably named slightly differently). Note its current location and behavior.

- [ ] **Step 2: Create `WebUI/markdown.js` with the extracted renderer**

Read the current `chat.js` markdown function, copy the exact implementation, paste it into a new file `VSTO2/OutlookAI/WebUI/markdown.js`. Wrap as a namespaced global so both consumers can call it:

```javascript
// VSTO2/OutlookAI/WebUI/markdown.js
// Markdown -> HTML renderer shared by chat.js (Inbox Copilot / Reports panes)
// and print-template.html (PDF export). Pure function, no DOM dependencies
// in the parsing path - only the consumer touches DOM with the returned string.
(function (root) {
    'use strict';

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function renderInline(src) {
        // bold ** **, italic * *, code ``, links [text](url)
        var out = escapeHtml(src);
        out = out.replace(/`([^`]+)`/g, function (_, c) { return '<code>' + c + '</code>'; });
        out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        out = out.replace(/\*([^*]+)\*/g, '<em>$1</em>');
        out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, function (_, t, u) {
            return '<a href="' + u + '">' + t + '</a>';
        });
        // Single newlines -> <br>; double newlines split paragraphs (handled upstream)
        out = out.replace(/\n/g, '<br>');
        return out;
    }

    function renderTable(lines) {
        // lines[0] = header, lines[1] = ---|---, lines[2..] = body
        var header = lines[0].split('|').slice(1, -1).map(function (c) { return c.trim(); });
        var body = lines.slice(2).map(function (l) {
            return l.split('|').slice(1, -1).map(function (c) { return c.trim(); });
        });
        var html = '<table><thead><tr>';
        header.forEach(function (h) { html += '<th>' + renderInline(h) + '</th>'; });
        html += '</tr></thead><tbody>';
        body.forEach(function (row) {
            html += '<tr>';
            row.forEach(function (c) { html += '<td>' + renderInline(c) + '</td>'; });
            html += '</tr>';
        });
        html += '</tbody></table>';
        return html;
    }

    function renderMarkdown(src) {
        if (src == null) return '';
        var lines = String(src).replace(/\r\n/g, '\n').split('\n');
        var out = [];
        var i = 0;
        while (i < lines.length) {
            var line = lines[i];

            // ATX headings
            var h = /^(#{1,6})\s+(.*)$/.exec(line);
            if (h) {
                var level = h[1].length;
                out.push('<h' + level + '>' + renderInline(h[2]) + '</h' + level + '>');
                i++; continue;
            }

            // Horizontal rule
            if (/^-{3,}\s*$/.test(line)) {
                out.push('<hr>');
                i++; continue;
            }

            // Blockquote
            if (/^>\s?/.test(line)) {
                var bq = [];
                while (i < lines.length && /^>\s?/.test(lines[i])) {
                    bq.push(lines[i].replace(/^>\s?/, ''));
                    i++;
                }
                out.push('<blockquote>' + renderInline(bq.join('\n')) + '</blockquote>');
                continue;
            }

            // Fenced code block
            if (/^```/.test(line)) {
                var code = [];
                i++;
                while (i < lines.length && !/^```/.test(lines[i])) { code.push(lines[i]); i++; }
                if (i < lines.length) i++; // skip closing fence
                out.push('<pre><code>' + escapeHtml(code.join('\n')) + '</code></pre>');
                continue;
            }

            // GFM table
            if (i + 1 < lines.length && /^\|.+\|$/.test(line) && /^\|[\s\-:|]+\|$/.test(lines[i + 1])) {
                var t = [line, lines[i + 1]];
                i += 2;
                while (i < lines.length && /^\|.+\|$/.test(lines[i])) { t.push(lines[i]); i++; }
                out.push(renderTable(t));
                continue;
            }

            // Blank line -> end of paragraph
            if (line.trim() === '') { i++; continue; }

            // Paragraph: collect contiguous non-empty lines
            var para = [line]; i++;
            while (i < lines.length
                && lines[i].trim() !== ''
                && !/^#{1,6}\s/.test(lines[i])
                && !/^-{3,}\s*$/.test(lines[i])
                && !/^>\s?/.test(lines[i])
                && !/^```/.test(lines[i])
                && !(/^\|.+\|$/.test(lines[i]) && i + 1 < lines.length && /^\|[\s\-:|]+\|$/.test(lines[i + 1]))) {
                para.push(lines[i]); i++;
            }
            out.push('<p>' + renderInline(para.join('\n')) + '</p>');
        }

        return out.join('\n');
    }

    root.markdown = root.markdown || {};
    root.markdown.render = renderMarkdown;
    root.markdown.escapeHtml = escapeHtml;
})(typeof window !== 'undefined' ? window : this);
```

(If the existing chat.js renderer has additional features — e.g., task lists, ordered lists, image stripping — port them too. Use the exact functions from the source as a starting point; the above is a baseline.)

- [ ] **Step 3: Update `chat.js` to delegate to `markdown.js`**

Replace the inline `renderMarkdown` definition in `chat.js` with:

```javascript
function renderMarkdown(src) {
    return window.markdown && window.markdown.render ? window.markdown.render(src) : (src || '');
}
```

(Keep all call sites the same; only the implementation moves.)

- [ ] **Step 4: Load `markdown.js` before `chat.js` in `index.html`**

In `VSTO2/OutlookAI/WebUI/index.html`, find the `<script src="chat.js">` line and add the markdown script tag **before** it:

```html
<script src="markdown.js"></script>
<script src="chat.js"></script>
```

- [ ] **Step 5: Register `markdown.js` as an embedded resource**

In `VSTO2/OutlookAI/OutlookAI.csproj`, find the WebUI EmbeddedResource section and add:

```xml
<EmbeddedResource Include="WebUI\markdown.js">
  <LogicalName>OutlookAI.WebUI.markdown.js</LogicalName>
</EmbeddedResource>
```

- [ ] **Step 6: Build to verify everything compiles**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds, `markdown.js` is included in the assembly's embedded resources.

- [ ] **Step 7: Manual smoke — verify chat still renders markdown**

Build + install + open Outlook + open chat pane + paste a markdown message. Headings, bold, tables should render exactly as before. This is a refactor — no visible change.

(If installer is not yet available in this dev cycle, defer the smoke check until Task 25's batch install.)

- [ ] **Step 8: Run full suite**

Expected: `Passed: 393` (no new tests; this is a refactor).

- [ ] **Step 9: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/markdown.js VSTO2/OutlookAI/WebUI/chat.js VSTO2/OutlookAI/WebUI/index.html VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "refactor(webui): extract markdown renderer to shared module"
```

---

## Task 12: `PrintTemplateRenderer` — pure HTML composition

Composes the PDF print-template HTML by substituting tokens (`__TITLE_TEXT__`, `__SUBTITLE_TEXT__`, `__GENERATED_AT__`) and injecting the markdown content via a JSON-encoded `window.__OUTLOOKAI_MD__` literal. Strips inline images. HTML-encodes title/subtitle to prevent XSS via injected attributes.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/PrintTemplateRenderer.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Export/PrintTemplateRendererTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Export/PrintTemplateRendererTests.cs`:

```csharp
using System;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class PrintTemplateRendererTests
    {
        private static string Template = @"<!doctype html>
<html><head><meta charset=""utf-8""><title>__TITLE_TEXT__</title>
<link rel=""stylesheet"" href=""print-styles.css""><script src=""markdown.js""></script></head>
<body><header class=""doc-header"">
<h1 class=""doc-title"">__TITLE_TEXT__</h1>
<p class=""doc-subtitle"">__SUBTITLE_TEXT__</p>
<p class=""doc-meta"">Generated by OutlookAI &middot; __GENERATED_AT__</p>
</header><main id=""content""></main>
<script>__MD_INJECT__
const html = window.markdown.render(window.__OUTLOOKAI_MD__ || '');
document.getElementById('content').innerHTML = html;
document.body.dataset.renderState = 'ready';
</script></body></html>";

        private static DateTimeOffset FixedNow => new DateTimeOffset(2026, 5, 18, 19, 47, 0, TimeSpan.Zero);

        [Fact]
        public void Render_SubstitutesAllTokens()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("My Title", "My Subtitle", "## Heading\n\ncontent", FixedNow);
            Assert.Contains("<title>My Title</title>", html);
            Assert.Contains(">My Title</h1>", html);
            Assert.Contains("My Subtitle</p>", html);
            Assert.Contains("2026", html); // some date format
        }

        [Fact]
        public void Render_HtmlEncodesTitleToBlockXss()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("<script>alert(1)</script>", "", "x", FixedNow);
            Assert.DoesNotContain("<script>alert(1)</script>", html);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html);
        }

        [Fact]
        public void Render_JsonEncodesMarkdownInjection()
        {
            var renderer = new PrintTemplateRenderer(Template);
            // Markdown containing characters that would break a naive injection
            var markdown = "Hello \"world\" \n\\ </script>";
            var html = renderer.Render("t", "", markdown, FixedNow);
            // The literal characters should be JSON-escaped inside __OUTLOOKAI_MD__
            Assert.Contains("window.__OUTLOOKAI_MD__", html);
            Assert.DoesNotContain("</script>\n", html); // raw close-tag would break parsing
        }

        [Fact]
        public void Render_StripsInlineImages()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("t", "",
                "Look: ![alt text](https://example.com/img.png) and ![another](./local.jpg) end.",
                FixedNow);
            Assert.Contains("[image: alt text]", html);
            Assert.Contains("[image: another]", html);
            Assert.DoesNotContain("![alt", html);
            Assert.DoesNotContain("https://example.com/img.png", html);
        }

        [Fact]
        public void Render_EmptySubtitleProducesEmptyParagraphContent()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("Title", null, "x", FixedNow);
            Assert.Contains("doc-subtitle", html);
            // No replacement of __SUBTITLE_TEXT__ left over
            Assert.DoesNotContain("__SUBTITLE_TEXT__", html);
        }

        [Fact]
        public void Render_GeneratedAtFormatted()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("t", "", "x", FixedNow);
            Assert.Contains("May 18, 2026", html);
        }

        [Fact]
        public void Render_NoUnreplacedTokensRemain()
        {
            var renderer = new PrintTemplateRenderer(Template);
            var html = renderer.Render("Title", "Sub", "x", FixedNow);
            Assert.DoesNotContain("__TITLE_TEXT__", html);
            Assert.DoesNotContain("__SUBTITLE_TEXT__", html);
            Assert.DoesNotContain("__GENERATED_AT__", html);
            Assert.DoesNotContain("__MD_INJECT__", html);
        }

        [Fact]
        public void Render_CodeBlockImagesNotStripped()
        {
            // Inside fenced code blocks the ![alt](url) syntax should remain
            // as-is (it's documentation, not an image to fetch).
            var renderer = new PrintTemplateRenderer(Template);
            var md = "```\n![keep](url)\n```";
            var html = renderer.Render("t", "", md, FixedNow);
            // Markdown is JSON-injected; verify the literal substring is there.
            Assert.Contains("![keep](url)", html);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~PrintTemplateRenderer"
```

Expected: build fails.

- [ ] **Step 3: Implement `PrintTemplateRenderer`**

Create `VSTO2/OutlookAI/Services/Export/PrintTemplateRenderer.cs`:

```csharp
using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Pure HTML composition for the PDF print template. Substitutes tokens,
    /// HTML-encodes title/subtitle, JSON-encodes markdown for safe script
    /// injection, and strips inline images outside of fenced code blocks.
    /// </summary>
    public sealed class PrintTemplateRenderer
    {
        private readonly string _template;

        public PrintTemplateRenderer(string templateHtml)
        {
            if (templateHtml == null) throw new ArgumentNullException(nameof(templateHtml));
            _template = templateHtml;
        }

        /// <summary>Test-friendly ctor that loads from a file path.</summary>
        public static PrintTemplateRenderer LoadFromFile(string templatePath)
            => new PrintTemplateRenderer(File.ReadAllText(templatePath));

        public string Render(string title, string subtitle, string markdown, DateTimeOffset generatedAt)
        {
            var encodedTitle = WebUtility.HtmlEncode(title ?? "");
            var encodedSubtitle = WebUtility.HtmlEncode(subtitle ?? "");
            var stamp = generatedAt.ToLocalTime().ToString("MMMM d, yyyy 'at' h:mm tt", CultureInfo.InvariantCulture);
            var strippedMarkdown = StripInlineImages(markdown ?? "");
            var mdLiteral = "window.__OUTLOOKAI_MD__ = " + JsonConvert.SerializeObject(strippedMarkdown) + ";";

            return _template
                .Replace("__TITLE_TEXT__", encodedTitle)
                .Replace("__SUBTITLE_TEXT__", encodedSubtitle)
                .Replace("__GENERATED_AT__", stamp)
                .Replace("__MD_INJECT__", mdLiteral);
        }

        private static readonly Regex ImageOutsideFence = new Regex(@"!\[([^\]]*)\]\([^)]+\)", RegexOptions.Compiled);

        private static string StripInlineImages(string markdown)
        {
            // Strip images only OUTSIDE fenced code blocks. Walk line-by-line.
            var lines = markdown.Replace("\r\n", "\n").Split('\n');
            bool inFence = false;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("```", StringComparison.Ordinal))
                {
                    inFence = !inFence;
                    continue;
                }
                if (inFence) continue;
                lines[i] = ImageOutsideFence.Replace(lines[i],
                    m => "[image: " + m.Groups[1].Value + "]");
            }
            return string.Join("\n", lines);
        }
    }
}
```

Register in csproj:

```xml
<Compile Include="Services\Export\PrintTemplateRenderer.cs" />
```

Test:

```xml
<Compile Include="Services\Export\PrintTemplateRendererTests.cs" />
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~PrintTemplateRenderer"
```

Expected: 8 tests pass.

- [ ] **Step 5: Full suite**

Expected: 401 (393 + 8).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/PrintTemplateRenderer.cs VSTO2/OutlookAI.Tests/Services/Export/PrintTemplateRendererTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add PDF print template renderer"
```

---

## Task 13: Add `print-template.html` and `print-styles.css` as embedded resources

The print template HTML skeleton plus a print-tuned stylesheet (A4 page, frozen `thead` repeats on every page, page-break rules, Office-ish typography).

**Files:**
- Create: `VSTO2/OutlookAI/WebUI/print-template.html`
- Create: `VSTO2/OutlookAI/WebUI/print-styles.css`
- Modify: `VSTO2/OutlookAI/OutlookAI.csproj`

- [ ] **Step 1: Create the print template**

Create `VSTO2/OutlookAI/WebUI/print-template.html`:

```html
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>__TITLE_TEXT__</title>
  <link rel="stylesheet" href="print-styles.css">
  <script src="markdown.js"></script>
</head>
<body data-render-state="loading">
  <header class="doc-header">
    <h1 class="doc-title">__TITLE_TEXT__</h1>
    <p class="doc-subtitle">__SUBTITLE_TEXT__</p>
    <p class="doc-meta">Generated by OutlookAI &middot; __GENERATED_AT__</p>
  </header>
  <main id="content"></main>
  <script>
__MD_INJECT__
    try {
      var html = (window.markdown && window.markdown.render) ? window.markdown.render(window.__OUTLOOKAI_MD__ || '') : '';
      document.getElementById('content').innerHTML = html;
      document.body.setAttribute('data-render-state', 'ready');
    } catch (e) {
      document.getElementById('content').textContent = String(window.__OUTLOOKAI_MD__ || '');
      document.body.setAttribute('data-render-state', 'error');
    }
  </script>
</body>
</html>
```

- [ ] **Step 2: Create the print stylesheet**

Create `VSTO2/OutlookAI/WebUI/print-styles.css`:

```css
@page {
  size: A4;
  margin: 0.5in;
}

html, body {
  font-family: "Calibri", "Segoe UI", Arial, sans-serif;
  font-size: 11pt;
  color: #222;
  margin: 0;
  padding: 0;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}

.doc-header {
  border-bottom: 2px solid #2b579a;
  padding-bottom: 12px;
  margin-bottom: 24px;
}
.doc-title {
  font-size: 22pt;
  font-weight: 600;
  margin: 0;
  color: #1f3a5f;
}
.doc-subtitle {
  font-size: 13pt;
  color: #555;
  margin: 4px 0 0;
}
.doc-meta {
  font-size: 9pt;
  color: #888;
  margin: 6px 0 0;
  font-style: italic;
}

main { line-height: 1.45; }

h1, h2, h3, h4, h5, h6 {
  color: #1f3a5f;
  page-break-after: avoid;
  margin-top: 14pt;
  margin-bottom: 6pt;
}
h1 { font-size: 18pt; }
h2 { font-size: 14pt; }
h3 { font-size: 12pt; }
h4, h5, h6 { font-size: 11pt; }

p { margin: 4pt 0; }

table {
  border-collapse: collapse;
  width: 100%;
  margin: 8pt 0;
  page-break-inside: auto;
}
thead { display: table-header-group; }  /* repeats header on each page */
tfoot { display: table-footer-group; }
tr { page-break-inside: avoid; }
th, td {
  border: 1px solid #ccc;
  padding: 4pt 6pt;
  text-align: left;
  vertical-align: top;
  font-size: 10pt;
}
th {
  background: #f3f4f6;
  font-weight: 600;
}

blockquote {
  border-left: 3px solid #2b579a;
  padding-left: 10pt;
  margin: 8pt 0;
  color: #555;
}

code {
  font-family: "Consolas", "Courier New", monospace;
  font-size: 9.5pt;
  background: #f6f8fa;
  padding: 1pt 3pt;
  border-radius: 2px;
}
pre {
  font-family: "Consolas", "Courier New", monospace;
  font-size: 9.5pt;
  background: #f6f8fa;
  padding: 6pt;
  border-radius: 4px;
  page-break-inside: avoid;
  white-space: pre-wrap;
  overflow-x: auto;
}
pre code {
  background: transparent;
  padding: 0;
}

a {
  color: #2b579a;
  text-decoration: none;
}

hr {
  border: 0;
  border-top: 1px solid #d0d7de;
  margin: 12pt 0;
}

@media print {
  body { margin: 0; }
}
```

- [ ] **Step 3: Register both as embedded resources**

In `VSTO2/OutlookAI/OutlookAI.csproj`, in the same `<ItemGroup>` as the existing WebUI resources, add:

```xml
<EmbeddedResource Include="WebUI\print-template.html">
  <LogicalName>OutlookAI.WebUI.print-template.html</LogicalName>
</EmbeddedResource>
<EmbeddedResource Include="WebUI\print-styles.css">
  <LogicalName>OutlookAI.WebUI.print-styles.css</LogicalName>
</EmbeddedResource>
```

Note: `WebView2Bootstrap.ResourceNameToRelativePath` treats the last `.` as the extension separator, so `print-template.html` will be extracted to `<WebUI>/print-template.html` and `print-styles.css` to `<WebUI>/print-styles.css`. Verify by reading `ResourceNameToRelativePath` (already in the codebase) and confirming the resource names map correctly. If `WebView2Bootstrap` strips the hyphen unexpectedly, rename the files to e.g. `printtemplate.html` / `printstyles.css` (no hyphens, no dots-in-stem).

- [ ] **Step 4: Build to verify embedding works**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: builds cleanly. Manual verification: after build, inspect the assembly resources to confirm both names are present:

```powershell
$asm = [Reflection.Assembly]::LoadFrom("VSTO2\OutlookAI\bin\Debug\OutlookAI.dll")
$asm.GetManifestResourceNames() | Where-Object { $_ -like "*print*" }
```

Expected output: includes `OutlookAI.WebUI.print-template.html` and `OutlookAI.WebUI.print-styles.css`.

- [ ] **Step 5: Full suite**

Expected: 401 (no test count change).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/print-template.html VSTO2/OutlookAI/WebUI/print-styles.css VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(export): add PDF print template and styles"
```

---

## Task 14: `PdfRenderer` — off-screen WebView2 with `PrintToPdfAsync`

Lazy-initialized, cached, off-screen WebView2 instance that navigates to the composed HTML and prints to a PDF file. Serialized via `SemaphoreSlim`. Disposed on add-in shutdown.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Export/PdfRenderer.cs`

This task is integration-level (real WebView2 instance) and cannot be unit-tested from `net472` xUnit. Smoke validation happens in Task 27.

- [ ] **Step 1: Implement `PdfRenderer`**

Create `VSTO2/OutlookAI/Services/Export/PdfRenderer.cs`:

```csharp
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OutlookAI.Diagnostics;
using OutlookAI.TaskPane.Chat;

namespace OutlookAI.Services.Export
{
    /// <summary>
    /// Owns an off-screen WebView2 instance that renders an HTML document and
    /// prints it to PDF via PrintToPdfAsync. Lazy-initialized, cached for the
    /// life of the add-in, serialized via SemaphoreSlim (one PDF at a time).
    /// All access must be on the UI thread (WebView2 is STA).
    /// </summary>
    public sealed class PdfRenderer : IDisposable
    {
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);
        private Form _hostForm;
        private WebView2 _webView;
        private bool _initialized;
        private bool _disposed;

        /// <summary>
        /// Render <paramref name="html"/> to a PDF file at <paramref name="outputPath"/>.
        /// Returns the file size in bytes on success. Throws
        /// <see cref="ExportException"/> on failure.
        /// </summary>
        public async Task<long> RenderAsync(string html, string outputPath, CancellationToken ct)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(PdfRenderer));
            if (string.IsNullOrEmpty(html)) throw new ArgumentException("html is empty", nameof(html));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentException("outputPath is empty", nameof(outputPath));

            TraceLog.Write(">> RenderAsync " + outputPath, "PdfRenderer");
            await _gate.WaitAsync(ct).ConfigureAwait(true);
            try
            {
                await EnsureInitializedAsync().ConfigureAwait(true);

                // Navigate and wait for completion + render-ready signal.
                var navTcs = new TaskCompletionSource<bool>();
                CoreWebView2NavigationCompletedEventHandler handler = null;
                handler = (s, e) =>
                {
                    try { _webView.CoreWebView2.NavigationCompleted -= handler; } catch { }
                    navTcs.TrySetResult(e.IsSuccess);
                };
                _webView.CoreWebView2.NavigationCompleted += handler;
                _webView.CoreWebView2.NavigateToString(html);

                using (ct.Register(() => navTcs.TrySetCanceled()))
                {
                    var ok = await navTcs.Task.ConfigureAwait(true);
                    if (!ok)
                        throw new ExportException("pdf_render_failed", "navigation reported IsSuccess=false");
                }

                // Poll render-ready up to 5 seconds.
                var ready = await WaitForRenderReadyAsync(ct).ConfigureAwait(true);
                if (!ready)
                    throw new ExportException("pdf_render_timeout", "render-ready signal didn't fire within 5s");

                // Print.
                try
                {
                    var settings = _webView.CoreWebView2.Environment.CreatePrintSettings();
                    settings.ShouldPrintBackgrounds = true;
                    settings.ShouldPrintHeaderAndFooter = false;
                    settings.Orientation = CoreWebView2PrintOrientation.Portrait;
                    settings.ScaleFactor = 1.0;
                    var success = await _webView.CoreWebView2.PrintToPdfAsync(outputPath, settings).ConfigureAwait(true);
                    if (!success)
                        throw new ExportException("pdf_print_failed", "PrintToPdfAsync returned false");
                }
                catch (ExportException) { throw; }
                catch (Exception ex)
                {
                    throw new ExportException("pdf_print_failed", ex.Message, ex);
                }

                var size = new FileInfo(outputPath).Length;
                TraceLog.Write("<< RenderAsync " + size + " bytes", "PdfRenderer");
                return size;
            }
            finally
            {
                _gate.Release();
            }
        }

        private async Task EnsureInitializedAsync()
        {
            if (_initialized) return;
            TraceLog.Write("PdfRenderer init", "PdfRenderer");

            if (!WebView2Bootstrap.IsRuntimeInstalled())
                throw new ExportException("webview2_missing",
                    "WebView2 runtime not installed. Install Edge WebView2 Evergreen Runtime.");

            // Create an invisible host form so the WebView2 has an HWND.
            _hostForm = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Size = new System.Drawing.Size(800, 600),
                Location = new System.Drawing.Point(-32000, -32000), // off-screen
                Opacity = 0,
                Visible = true   // must be created/shown to host the WebView2
            };
            _webView = new WebView2 { Dock = DockStyle.Fill };
            _hostForm.Controls.Add(_webView);

            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: WebView2Bootstrap.WebView2DataFolder,
                options: null).ConfigureAwait(true);
            await _webView.EnsureCoreWebView2Async(env).ConfigureAwait(true);

            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

            // Map the same virtual host as the chat so <script src="markdown.js">
            // and <link rel="stylesheet" href="print-styles.css"> resolve.
            _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                WebView2Bootstrap.VirtualHost,
                WebView2Bootstrap.WebUiFolder,
                CoreWebView2HostResourceAccessKind.Allow);

            _initialized = true;
        }

        private async Task<bool> WaitForRenderReadyAsync(CancellationToken ct)
        {
            const int timeoutMs = 5000;
            const int intervalMs = 50;
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                ct.ThrowIfCancellationRequested();
                var script = "document.body && document.body.getAttribute('data-render-state')";
                var raw = await _webView.CoreWebView2.ExecuteScriptAsync(script).ConfigureAwait(true);
                // ExecuteScriptAsync returns a JSON-encoded string.
                if (raw == "\"ready\"") return true;
                if (raw == "\"error\"") return false;
                await Task.Delay(intervalMs, ct).ConfigureAwait(true);
            }
            return false;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { _webView?.Dispose(); } catch { }
            try { _hostForm?.Dispose(); } catch { }
            _gate?.Dispose();
        }
    }
}
```

(Note: `CoreWebView2PrintOrientation` is the WebView2 1.0.864+ enum. If your installed WebView2 SDK is older and lacks this type, omit those lines — default orientation is Portrait.)

Register in csproj:

```xml
<Compile Include="Services\Export\PdfRenderer.cs" />
```

- [ ] **Step 2: Build to verify the WebView2 + Windows Forms references resolve**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds. If you see `CoreWebView2PrintSettings` missing, check that `Microsoft.Web.WebView2.Core` package is at version `>=1.0.864`.

- [ ] **Step 3: Construct as a singleton in `ThisAddIn`**

In `VSTO2/OutlookAI/ThisAddIn.cs`, add a private field and lazy-initialize:

```csharp
private PdfRenderer _pdfRenderer;
public PdfRenderer PdfRenderer => _pdfRenderer ?? (_pdfRenderer = new PdfRenderer());

private ExportPathResolver _exportPathResolver;
public ExportPathResolver ExportPathResolver => _exportPathResolver ?? (_exportPathResolver = new ExportPathResolver());

private IExportPathPolicy _exportPathPolicy;
public IExportPathPolicy ExportPathPolicy => _exportPathPolicy ?? (_exportPathPolicy = new ExportPathPolicy(ExportPathResolver));
```

In `ThisAddIn_Shutdown`:

```csharp
try { _pdfRenderer?.Dispose(); } catch { }
```

(Add `using OutlookAI.Services.Export;` at the top of `ThisAddIn.cs`.)

- [ ] **Step 4: Full suite**

Expected: 401 (no test count change).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Export/PdfRenderer.cs VSTO2/OutlookAI/ThisAddIn.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "feat(export): add WebView2-backed PdfRenderer"
```

---

## Task 15: `ExportPdfArgsParser`

JSON parser for the PDF tool's args. Validates `content_markdown` is non-empty and ≤ 250 000 chars; clamps `title` to 200 chars, `subtitle` to 400.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ExportPdfArgs.cs` (already created as a stub in Task 7; no changes needed — DTO is fine).
- Create: `VSTO2/OutlookAI/Services/Tools/ExportPdfArgsParser.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/ExportPdfArgsParserTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/ExportPdfArgsParserTests.cs`:

```csharp
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportPdfArgsParserTests
    {
        [Fact]
        public void Parses_ValidArgs()
        {
            var json = @"{""filename_hint"":""Quotes"",""title"":""Q1 Quotes"",""subtitle"":""May 2026"",""content_markdown"":""# Hi""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Equal("Quotes", args.FilenameHint);
            Assert.Equal("Q1 Quotes", args.Title);
            Assert.Equal("May 2026", args.Subtitle);
            Assert.Equal("# Hi", args.ContentMarkdown);
        }

        [Fact]
        public void Throws_WhenContentMarkdownMissing()
        {
            var json = @"{""filename_hint"":""x""}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(json));
            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Fact]
        public void Throws_WhenContentMarkdownEmpty()
        {
            var json = @"{""filename_hint"":""x"",""content_markdown"":"" ""}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(json));
            Assert.Equal("invalid_args", ex.Code);
        }

        [Fact]
        public void Throws_WhenContentMarkdownExceeds250k()
        {
            var json = @"{""filename_hint"":""x"",""content_markdown"":""" + new string('A', 250_001) + @"""}";
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(json));
            Assert.Equal("content_too_large", ex.Code);
        }

        [Fact]
        public void Title_ClampedTo200Chars()
        {
            var bigTitle = new string('T', 500);
            var json = @"{""title"":""" + bigTitle + @""",""content_markdown"":""x""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Equal(200, args.Title.Length);
        }

        [Fact]
        public void Subtitle_ClampedTo400Chars()
        {
            var bigSub = new string('S', 1000);
            var json = @"{""subtitle"":""" + bigSub + @""",""content_markdown"":""x""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Equal(400, args.Subtitle.Length);
        }

        [Fact]
        public void FilenameHint_FallsBackToDefault_WhenMissing()
        {
            var json = @"{""content_markdown"":""x""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Equal("OutlookAI-Report", args.FilenameHint);
        }

        [Fact]
        public void FilenameHint_FallsBackToTitle_WhenHintMissingButTitlePresent()
        {
            var json = @"{""title"":""Quarterly Report"",""content_markdown"":""x""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Equal("Quarterly Report", args.FilenameHint);
        }

        [Fact]
        public void Title_AndSubtitle_DefaultToNullWhenMissing()
        {
            var json = @"{""filename_hint"":""x"",""content_markdown"":""y""}";
            var args = ExportPdfArgsParser.Parse(json);
            Assert.Null(args.Title);
            Assert.Null(args.Subtitle);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ExportPdfArgsParser"
```

Expected: build fails (parser missing).

- [ ] **Step 3: Implement the parser**

Create `VSTO2/OutlookAI/Services/Tools/ExportPdfArgsParser.cs`:

```csharp
using System;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    public static class ExportPdfArgsParser
    {
        private const int MaxMarkdown = 250_000;
        private const int MaxTitle = 200;
        private const int MaxSubtitle = 400;
        private const string DefaultFilenameHint = "OutlookAI-Report";

        public static ExportPdfArgs Parse(string json)
        {
            JObject root;
            try { root = JObject.Parse(json ?? ""); }
            catch (Exception ex)
            {
                throw new ToolArgValidationException("invalid_args", "could not parse JSON: " + ex.Message);
            }

            var content = (string)root["content_markdown"];
            if (string.IsNullOrWhiteSpace(content))
                throw new ToolArgValidationException("invalid_args", "content_markdown is required");
            if (content.Length > MaxMarkdown)
                throw new ToolArgValidationException("content_too_large",
                    $"content_markdown is {content.Length} chars; max allowed is {MaxMarkdown}");

            var title = ClampNullable((string)root["title"], MaxTitle);
            var subtitle = ClampNullable((string)root["subtitle"], MaxSubtitle);

            var hintRaw = (string)root["filename_hint"];
            string filenameHint;
            if (!string.IsNullOrWhiteSpace(hintRaw)) filenameHint = hintRaw.Trim();
            else if (!string.IsNullOrWhiteSpace(title)) filenameHint = title.Trim();
            else filenameHint = DefaultFilenameHint;

            return new ExportPdfArgs
            {
                FilenameHint = filenameHint,
                Title = title,
                Subtitle = subtitle,
                ContentMarkdown = content
            };
        }

        private static string ClampNullable(string s, int max)
        {
            if (s == null) return null;
            var trimmed = s.Trim();
            if (trimmed.Length == 0) return null;
            return trimmed.Length > max ? trimmed.Substring(0, max) : trimmed;
        }
    }
}
```

Register in csproj:

```xml
<Compile Include="Services\Tools\ExportPdfArgsParser.cs" />
```

Test:

```xml
<Compile Include="Services\Tools\ExportPdfArgsParserTests.cs" />
```

- [ ] **Step 4: Run tests + full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 410` (401 + 9).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ExportPdfArgsParser.cs VSTO2/OutlookAI.Tests/Services/Tools/ExportPdfArgsParserTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(export): add PDF args parser"
```

---

## Task 16: Implement `LiveOutlookSurface.ExportPdf`

Replace the `NotImplementedException` stub from Task 7. Composes the print template, calls the singleton `PdfRenderer`, returns a `FileSavedResult`.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`

- [ ] **Step 1: Replace the stub `ExportPdf` method**

In `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`, replace the `ExportPdf` stub with:

```csharp
public FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
{
    if (args == null) throw new System.ArgumentNullException(nameof(args));
    ct.ThrowIfCancellationRequested();

    var resolver = Globals.ThisAddIn?.ExportPathResolver ?? new ExportPathResolver();
    string baseDir;
    try
    {
        resolver.EnsureExists();
        baseDir = resolver.ResolveBaseDir();
    }
    catch (System.IO.IOException ex)
    {
        throw new ExportException("path_unavailable", ex.Message, ex);
    }

    var filename = ExportFilenameSanitizer.Build(
        args.FilenameHint, ".pdf", System.DateTimeOffset.Now,
        candidate => System.IO.File.Exists(System.IO.Path.Combine(baseDir, candidate)));
    var fullPath = System.IO.Path.Combine(baseDir, filename);

    // Load template HTML from the extracted WebUI folder.
    var templatePath = System.IO.Path.Combine(WebView2Bootstrap.WebUiFolder, "print-template.html");
    if (!System.IO.File.Exists(templatePath))
        throw new ExportException("pdf_render_failed", "print-template.html missing from WebUI folder");
    var templateHtml = System.IO.File.ReadAllText(templatePath);

    var rendererC = new PrintTemplateRenderer(templateHtml);
    var html = rendererC.Render(
        args.Title ?? args.FilenameHint,
        args.Subtitle,
        args.ContentMarkdown,
        System.DateTimeOffset.Now);

    var pdfRenderer = Globals.ThisAddIn?.PdfRenderer
        ?? throw new ExportException("pdf_render_failed", "PdfRenderer not initialized");

    // PrintToPdfAsync must run on the UI thread. Marshal through the
    // existing OutlookThreadMarshaller.
    long bytes;
    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
    if (marshaller != null
        && System.Threading.Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
    {
        bytes = marshaller.RunAsync(() => pdfRenderer.RenderAsync(html, fullPath, ct), ct)
            .ConfigureAwait(false).GetAwaiter().GetResult();
    }
    else
    {
        bytes = pdfRenderer.RenderAsync(html, fullPath, ct)
            .ConfigureAwait(false).GetAwaiter().GetResult();
    }

    return new FileSavedResult
    {
        Path = fullPath,
        FileUrl = new System.Uri(fullPath).AbsoluteUri,
        Format = "pdf",
        Bytes = bytes,
        Filename = filename
    };
}
```

(Imports needed at top of file: `using OutlookAI.Services.Export;`, `using OutlookAI.TaskPane.Chat;`.)

(Verify `OutlookThreadMarshaller.RunAsync<T>(Func<Task<T>>, ct)` exists with that signature. If not, check the existing pattern — there may be a slightly different overload. The principle is "marshal to UI thread for WebView2 calls".)

- [ ] **Step 2: Build**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds.

- [ ] **Step 3: Full suite**

Expected: 410 (no test count change; this is integration code).

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(export): wire LiveOutlookSurface.ExportPdf"
```

---

## Task 17: `OutlookExportPdfTool`

`IOutlookTool` wrapper. Mirrors `OutlookExportExcelTool` but for PDF.

**Files:**
- Create: `VSTO2/OutlookAI/Services/Tools/OutlookExportPdfTool.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportPdfToolTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportPdfToolTests.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportPdfToolTests
    {
        private sealed class FakeSurface : MinimalSurface
        {
            public ExportPdfArgs Captured;
            public FileSavedResult ResultToReturn;
            public ExportException ToThrow;

            public override FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
            {
                Captured = args;
                if (ct.IsCancellationRequested) throw new OperationCanceledException(ct);
                if (ToThrow != null) throw ToThrow;
                return ResultToReturn;
            }
        }

        private const string Valid = @"{""filename_hint"":""Quotes"",""content_markdown"":""# Hi""}";

        [Fact]
        public async Task ReturnsFileSavedEnvelopeOnSuccess()
        {
            var surface = new FakeSurface
            {
                ResultToReturn = new FileSavedResult
                {
                    Path = "p", FileUrl = "file:///p", Format = "pdf", Bytes = 100, Filename = "x.pdf"
                }
            };
            var tool = new OutlookExportPdfTool();
            var json = await tool.DispatchAsync(Valid, surface, CancellationToken.None);
            var obj = JObject.Parse(json);
            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.Equal("pdf", (string)obj["format"]);
        }

        [Fact]
        public async Task ReturnsErrorOnInvalidArgs()
        {
            var surface = new FakeSurface();
            var tool = new OutlookExportPdfTool();
            var json = await tool.DispatchAsync(@"{}", surface, CancellationToken.None);
            Assert.Equal("invalid_args", (string)JObject.Parse(json)["error"]);
        }

        [Fact]
        public async Task ReturnsErrorOnContentTooLarge()
        {
            var surface = new FakeSurface();
            var big = new string('A', 250_001);
            var json = "{\"content_markdown\":\"" + big + "\"}";
            var tool = new OutlookExportPdfTool();
            var result = await tool.DispatchAsync(json, surface, CancellationToken.None);
            Assert.Equal("content_too_large", (string)JObject.Parse(result)["error"]);
        }

        [Fact]
        public async Task ReturnsErrorOnExportException()
        {
            var surface = new FakeSurface
            {
                ToThrow = new ExportException("webview2_missing", "install runtime")
            };
            var tool = new OutlookExportPdfTool();
            var json = await tool.DispatchAsync(Valid, surface, CancellationToken.None);
            Assert.Equal("webview2_missing", (string)JObject.Parse(json)["error"]);
        }

        [Fact]
        public async Task PropagatesCancellation()
        {
            var surface = new FakeSurface();
            var cts = new CancellationTokenSource();
            cts.Cancel();
            var tool = new OutlookExportPdfTool();
            var json = await tool.DispatchAsync(Valid, surface, cts.Token);
            Assert.Equal("cancelled", (string)JObject.Parse(json)["error"]);
        }

        [Fact]
        public async Task NameIs_outlook_export_pdf()
        {
            Assert.Equal("outlook_export_pdf", new OutlookExportPdfTool().Name);
            await Task.CompletedTask;
        }
    }
}
```

- [ ] **Step 2: Implement `OutlookExportPdfTool`**

Create `VSTO2/OutlookAI/Services/Tools/OutlookExportPdfTool.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public sealed class OutlookExportPdfTool : IOutlookTool
    {
        public string Name => "outlook_export_pdf";

        public Task<string> DispatchAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ExportPdfArgs args;
            try
            {
                args = ExportPdfArgsParser.Parse(argsJson);
            }
            catch (ToolArgValidationException vex)
            {
                return Task.FromResult(ErrorEnvelope(vex.Code, vex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(ErrorEnvelope("invalid_args", ex.Message));
            }

            try
            {
                if (ct.IsCancellationRequested)
                    return Task.FromResult(ErrorEnvelope("cancelled", "tool cancelled before dispatch"));

                var result = surface.ExportPdf(args, ct);
                return Task.FromResult(SuccessEnvelope(result));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(ErrorEnvelope("cancelled", "tool cancelled"));
            }
            catch (ExportException eex)
            {
                return Task.FromResult(ErrorEnvelope(eex.Code, eex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(ErrorEnvelope("pdf_render_failed", ex.Message));
            }
        }

        private static string SuccessEnvelope(FileSavedResult r)
            => new JObject(
                new JProperty("result_type", "file_saved"),
                new JProperty("path", r.Path),
                new JProperty("file_url", r.FileUrl),
                new JProperty("format", r.Format),
                new JProperty("bytes", r.Bytes),
                new JProperty("filename", r.Filename)).ToString(Newtonsoft.Json.Formatting.None);

        private static string ErrorEnvelope(string code, string detail)
            => new JObject(
                new JProperty("error", code),
                new JProperty("detail", detail ?? "")).ToString(Newtonsoft.Json.Formatting.None);
    }
}
```

Register in csproj:

```xml
<Compile Include="Services\Tools\OutlookExportPdfTool.cs" />
```

Test:

```xml
<Compile Include="Services\Tools\OutlookExportPdfToolTests.cs" />
```

- [ ] **Step 3: Run tests + full suite**

Expected: `Passed: 416` (410 + 6).

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookExportPdfTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookExportPdfToolTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(tools): add outlook_export_pdf tool"
```

---

## Task 18: Add `outlook_export_pdf` schema

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`

- [ ] **Step 1: Write the failing tests**

Append to `ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void OutlookExportPdf_IsRegistered()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    Assert.Contains(arr, t => (string)t["name"] == "outlook_export_pdf");
}

[Fact]
public void OutlookExportPdf_Description_TeachesPolishedReportUseCase()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_pdf");
    var desc = ((string)entry["description"]) ?? "";
    Assert.Contains("PDF", desc);
    Assert.Contains("polished", desc, System.StringComparison.OrdinalIgnoreCase);
    Assert.Contains("markdown", desc, System.StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void OutlookExportPdf_Schema_HasContentMarkdown()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_pdf");
    var props = entry["parameters"]?["properties"];
    Assert.NotNull(props["filename_hint"]);
    Assert.NotNull(props["content_markdown"]);
    Assert.NotNull(props["title"]);
    Assert.NotNull(props["subtitle"]);
}

[Fact]
public void OutlookExportPdf_TeachesWhenToChooseVsExcel()
{
    var arr = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var entry = arr.First(t => (string)t["name"] == "outlook_export_pdf");
    var desc = ((string)entry["description"]) ?? "";
    Assert.Contains("Excel", desc); // mentions when to prefer Excel instead
}
```

- [ ] **Step 2: Add the schema entry**

In `BuildResponsesToolsArray`, add after the Excel export entry:

```csharp
BuildToolEntry("outlook_export_pdf",
    "Save a polished, printable report to a PDF (.pdf) file. Use when the user asks for a PDF, printable, document, or 'send me a report I can share'. "
    + "Pass polished markdown - title + headings + tables + lists. The tool renders it into a clean A4 document with a header bar, no chat-UI chrome. "
    + "For tabular structured data (lists of messages, vendor breakdowns, etc.) PREFER outlook_export_excel - it produces a real spreadsheet with formatted cells, autofilter, and frozen headers. Use PDF for narrative reports, summaries, action items, digest-style outputs. "
    + "You may COMPOSE FRESH markdown specifically for the PDF (better headings, polished phrasing) instead of reusing what you just said in chat - the PDF is a separate artifact the user will share. "
    + "Example: user says 'make a PDF summarizing my week with this customer' -> compose markdown with title, sections per topic, action items, and call this tool. "
    + "The file is saved to ~\\Documents\\OutlookAI\\Reports\\ with an auto-generated timestamped filename. After success, mention the filename briefly; the UI surfaces an Open / Show-in-folder card automatically. "
    + "Max content_markdown length: 250,000 characters.",
    new JObject(
        new JProperty("type", "object"),
        new JProperty("properties", new JObject(
            new JProperty("filename_hint", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Short human-readable hint for the filename. Optional - defaults to title or 'OutlookAI-Report'."))),
            new JProperty("title", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Document title shown at the top of the PDF. Optional but recommended. Max 200 chars."))),
            new JProperty("subtitle", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Optional subtitle/byline below the title. Max 400 chars."))),
            new JProperty("content_markdown", new JObject(
                new JProperty("type", "string"),
                new JProperty("description", "Markdown body of the report. Supports headings (#), tables (GFM), lists, blockquotes, code blocks, bold/italic, links. Inline images are stripped. Max 250,000 chars."))))),
        new JProperty("required", new JArray("content_markdown")),
        new JProperty("additionalProperties", false))),
```

- [ ] **Step 3: Run tests + full suite**

Expected: `Passed: 420` (416 + 4).

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(schema): teach model when to use PDF export"
```

---

## Task 19: Register `OutlookExportPdfTool` in `OutlookToolHost`

**Files:**
- Modify: `VSTO2/OutlookAI/Services/OutlookToolHost.cs`

- [ ] **Step 1: Add the tool to the list**

After `new OutlookExportExcelTool(),` add:

```csharp
new OutlookExportPdfTool(),             // Phase 5: PDF export
```

- [ ] **Step 2: Build + full suite**

Expected: `Passed: 420`.

- [ ] **Step 3: Commit**

```powershell
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs
git commit -m "feat(host): register PDF export tool"
```

---

## Task 20: `ExportBridge` — handles `export_pdf`, `open_file`, `reveal_in_explorer` web messages

Encapsulates the three new WebMessage types so each chat controller can wire them with one line. Calls `LiveOutlookSurface.ExportPdf` directly (bypassing the model) for the per-message PDF button. Uses `IExportPathPolicy` to guard `open_file` / `reveal_in_explorer`.

**Files:**
- Create: `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs`
- Create: `VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/Chat/ChatController.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotController.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs`

- [ ] **Step 1: Write the failing tests**

Create `VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;
using Xunit;

namespace OutlookAI.Tests.TaskPane.Chat
{
    public class ExportBridgeTests
    {
        private sealed class FakeSurface : MinimalSurface
        {
            public ExportPdfArgs CapturedPdf;
            public FileSavedResult PdfResult = new FileSavedResult { Path = "p", Format = "pdf", Filename = "x.pdf", Bytes = 0, FileUrl = "file:///p" };
            public override FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
            {
                CapturedPdf = args;
                return PdfResult;
            }
        }

        private sealed class FakePolicy : IExportPathPolicy
        {
            public List<string> Validated = new List<string>();
            public bool Reject;
            public void RequireInsideReportsDir(string path)
            {
                Validated.Add(path);
                if (Reject) throw new UnauthorizedExportPathException("not allowed");
            }
        }

        private sealed class FakeJsRunner
        {
            public List<string> Scripts = new List<string>();
            public Task RunScript(string s) { Scripts.Add(s); return Task.CompletedTask; }
        }

        [Fact]
        public async Task ExportPdf_PostsBackFileCardScript()
        {
            var surface = new FakeSurface();
            var policy = new FakePolicy();
            var js = new FakeJsRunner();
            var bridge = new ExportBridge(surface, policy, js.RunScript);

            var payload = new JObject(
                new JProperty("message_id", "asst_3"),
                new JProperty("filename_hint", "x"),
                new JProperty("content_markdown", "# Hi"));
            var handled = await bridge.HandleAsync("export_pdf", payload, CancellationToken.None);

            Assert.True(handled);
            Assert.NotNull(surface.CapturedPdf);
            Assert.Equal("# Hi", surface.CapturedPdf.ContentMarkdown);
            // JS-side outlookai.onFileSaved("asst_3", {...}) script was invoked
            Assert.Contains(js.Scripts, s => s.Contains("onFileSaved") && s.Contains("asst_3"));
        }

        [Fact]
        public async Task ExportPdf_OnExceptionPostsErrorCard()
        {
            var surface = new FakeSurface();
            // FakeSurface always succeeds; use a derived surface that throws
            var throwingSurface = new ThrowingSurface();
            var policy = new FakePolicy();
            var js = new FakeJsRunner();
            var bridge = new ExportBridge(throwingSurface, policy, js.RunScript);

            var payload = new JObject(
                new JProperty("message_id", "asst_3"),
                new JProperty("filename_hint", "x"),
                new JProperty("content_markdown", "# Hi"));
            var handled = await bridge.HandleAsync("export_pdf", payload, CancellationToken.None);

            Assert.True(handled);
            Assert.Contains(js.Scripts, s => s.Contains("onExportError") && s.Contains("pdf_render_failed"));
        }

        private sealed class ThrowingSurface : MinimalSurface
        {
            public override FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
                => throw new ExportException("pdf_render_failed", "boom");
        }

        [Fact]
        public async Task OpenFile_ValidatesPathThenPostsBackOk()
        {
            var surface = new FakeSurface();
            var policy = new FakePolicy();
            var js = new FakeJsRunner();
            // Use a launcher we control instead of Process.Start so the test stays hermetic.
            var launched = new List<string>();
            var bridge = new ExportBridge(surface, policy, js.RunScript, openFileLauncher: launched.Add, revealInExplorerLauncher: _ => { });

            var payload = new JObject(new JProperty("path", @"C:\fake\Reports\file.xlsx"));
            var handled = await bridge.HandleAsync("open_file", payload, CancellationToken.None);

            Assert.True(handled);
            Assert.Single(policy.Validated);
            Assert.Equal(@"C:\fake\Reports\file.xlsx", launched[0]);
        }

        [Fact]
        public async Task OpenFile_RejectedPathLogsToastNotProcessStart()
        {
            var surface = new FakeSurface();
            var policy = new FakePolicy { Reject = true };
            var js = new FakeJsRunner();
            var launched = new List<string>();
            var bridge = new ExportBridge(surface, policy, js.RunScript, openFileLauncher: launched.Add, revealInExplorerLauncher: _ => { });

            var payload = new JObject(new JProperty("path", @"C:\Windows\System32\cmd.exe"));
            var handled = await bridge.HandleAsync("open_file", payload, CancellationToken.None);

            Assert.True(handled);
            Assert.Empty(launched);
            // Toast script invoked
            Assert.Contains(js.Scripts, s => s.Contains("showError") || s.Contains("onExportError"));
        }

        [Fact]
        public async Task UnknownType_NotHandled()
        {
            var surface = new FakeSurface();
            var policy = new FakePolicy();
            var js = new FakeJsRunner();
            var bridge = new ExportBridge(surface, policy, js.RunScript);
            var handled = await bridge.HandleAsync("send", new JObject(), CancellationToken.None);
            Assert.False(handled);
        }
    }
}
```

- [ ] **Step 2: Implement `ExportBridge`**

Create `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs`:

```csharp
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.Chat
{
    /// <summary>
    /// Handles three WebMessage types posted by the WebUI:
    /// <list type="bullet">
    ///   <item><c>export_pdf</c> - per-message Save-as-PDF button click.</item>
    ///   <item><c>open_file</c> - open a saved file via the OS default app.</item>
    ///   <item><c>reveal_in_explorer</c> - open Explorer with the file selected.</item>
    /// </list>
    /// Each chat controller delegates to <see cref="HandleAsync"/> first; if it
    /// returns true, the controller skips its own switch.
    /// </summary>
    public sealed class ExportBridge
    {
        private readonly IOutlookSurface _surface;
        private readonly IExportPathPolicy _pathPolicy;
        private readonly Func<string, Task> _runScript;
        private readonly Action<string> _openFile;
        private readonly Action<string> _revealInExplorer;

        public ExportBridge(IOutlookSurface surface, IExportPathPolicy pathPolicy, Func<string, Task> runScript,
            Action<string> openFileLauncher = null,
            Action<string> revealInExplorerLauncher = null)
        {
            _surface = surface ?? throw new ArgumentNullException(nameof(surface));
            _pathPolicy = pathPolicy ?? throw new ArgumentNullException(nameof(pathPolicy));
            _runScript = runScript ?? throw new ArgumentNullException(nameof(runScript));
            _openFile = openFileLauncher ?? DefaultOpenFile;
            _revealInExplorer = revealInExplorerLauncher ?? DefaultReveal;
        }

        public async Task<bool> HandleAsync(string type, JObject payload, CancellationToken ct)
        {
            if (type == null) return false;
            switch (type)
            {
                case "export_pdf":
                    await HandleExportPdfAsync(payload, ct).ConfigureAwait(false);
                    return true;
                case "open_file":
                    HandleOpenFile(payload);
                    return true;
                case "reveal_in_explorer":
                    HandleReveal(payload);
                    return true;
                default:
                    return false;
            }
        }

        private async Task HandleExportPdfAsync(JObject payload, CancellationToken ct)
        {
            var messageId = (string)payload?["message_id"] ?? "";
            var hint = (string)payload?["filename_hint"] ?? "";
            var content = (string)payload?["content_markdown"] ?? "";

            TraceLog.Write("export_pdf for " + messageId + " (" + content.Length + " chars)", "ExportBridge");

            try
            {
                var args = new ExportPdfArgs
                {
                    FilenameHint = string.IsNullOrWhiteSpace(hint) ? "OutlookAI-Report" : hint,
                    Title = null,
                    Subtitle = null,
                    ContentMarkdown = content
                };
                var result = _surface.ExportPdf(args, ct);

                var resultJson = new JObject(
                    new JProperty("path", result.Path),
                    new JProperty("file_url", result.FileUrl),
                    new JProperty("format", result.Format),
                    new JProperty("bytes", result.Bytes),
                    new JProperty("filename", result.Filename));
                var script = "outlookai.onFileSaved(" +
                             JsString(messageId) + ", " +
                             resultJson.ToString(Newtonsoft.Json.Formatting.None) + ");";
                await _runScript(script).ConfigureAwait(false);
            }
            catch (ExportException ex)
            {
                TraceLog.Write("export_pdf error: " + ex.Code + " - " + ex.Message, "ExportBridge");
                await PostError(messageId, ex.Code, ex.Message).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                TraceLog.Write("export_pdf UNHANDLED: " + ex, "ExportBridge");
                await PostError(messageId, "pdf_render_failed", ex.Message).ConfigureAwait(false);
            }
        }

        private async Task PostError(string messageId, string code, string detail)
        {
            var err = new JObject(
                new JProperty("error", code),
                new JProperty("detail", detail ?? ""));
            var script = "outlookai.onExportError(" + JsString(messageId) + ", " +
                         err.ToString(Newtonsoft.Json.Formatting.None) + ");";
            await _runScript(script).ConfigureAwait(false);
        }

        private void HandleOpenFile(JObject payload)
        {
            var path = (string)payload?["path"] ?? "";
            try
            {
                _pathPolicy.RequireInsideReportsDir(path);
                TraceLog.Write("open_file " + path, "ExportBridge");
                _openFile(path);
            }
            catch (UnauthorizedExportPathException ex)
            {
                TraceLog.Write("open_file REJECTED: " + ex.Message, "ExportBridge");
                _ = _runScript("outlookai.onExportError(null, {error:'unauthorized_path', detail:" +
                    JsString(ex.Message) + "});");
            }
            catch (Exception ex)
            {
                _ = _runScript("outlookai.onExportError(null, {error:'open_failed', detail:" +
                    JsString(ex.Message) + "});");
            }
        }

        private void HandleReveal(JObject payload)
        {
            var path = (string)payload?["path"] ?? "";
            try
            {
                _pathPolicy.RequireInsideReportsDir(path);
                TraceLog.Write("reveal_in_explorer " + path, "ExportBridge");
                _revealInExplorer(path);
            }
            catch (UnauthorizedExportPathException ex)
            {
                _ = _runScript("outlookai.onExportError(null, {error:'unauthorized_path', detail:" +
                    JsString(ex.Message) + "});");
            }
            catch (Exception ex)
            {
                _ = _runScript("outlookai.onExportError(null, {error:'reveal_failed', detail:" +
                    JsString(ex.Message) + "});");
            }
        }

        private static void DefaultOpenFile(string path)
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }

        private static void DefaultReveal(string path)
        {
            Process.Start("explorer.exe", "/select,\"" + path + "\"");
        }

        private static string JsString(string s)
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");
    }
}
```

Register in csproj:

```xml
<Compile Include="TaskPane\Chat\ExportBridge.cs" />
```

Test:

```xml
<Compile Include="TaskPane\Chat\ExportBridgeTests.cs" />
```

(The test folder may not exist yet — create it.)

- [ ] **Step 3: Wire `ExportBridge` into each controller**

In each of `ChatController.cs`, `InboxCopilotController.cs`, `InboxReportsController.cs`:

Add a field:

```csharp
private ExportBridge _exportBridge;
```

Initialize in the constructor (or in `InitializeAsync` right after `_surface` is available):

```csharp
_exportBridge = new ExportBridge(
    _surface,
    Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver()),
    RunScript);
```

In `HandleHostMessage(string type, JObject payload)`, at the **top** of the switch, add a guard that runs the bridge first:

```csharp
private async void HandleHostMessage(string type, JObject payload)
{
    if (_exportBridge != null)
    {
        var handled = await _exportBridge.HandleAsync(type, payload, _activeCts?.Token ?? CancellationToken.None).ConfigureAwait(true);
        if (handled) return;
    }
    switch (type) { /* existing cases */ }
}
```

(Adjust method signature from `void` to `async void` if needed. The pattern is similar to `WebMessageReceived → HandleHostMessage` you already use; it's the JSON parsing wrapper that gets the type string.)

- [ ] **Step 4: Build + run tests + full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 425` (420 + 5).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs VSTO2/OutlookAI/TaskPane/Chat/ChatController.cs VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotController.cs VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs VSTO2/OutlookAI/OutlookAI.csproj VSTO2/OutlookAI.Tests/OutlookAI.Tests.csproj
git commit -m "feat(bridge): add export PDF and file-action methods"
```

---

## Task 21: Render `result_type:"file_saved"` tool results as a FileCard in `chat.js`

When the model dispatches `outlook_export_excel` or `outlook_export_pdf`, the tool result contains `result_type:"file_saved"`. The chat renderer should detect this and render a rich card instead of raw JSON.

**Files:**
- Modify: `VSTO2/OutlookAI/WebUI/chat.js`
- Modify: `VSTO2/OutlookAI/WebUI/styles.css`

- [ ] **Step 1: Locate where tool results are currently rendered**

```powershell
Select-String -Path "VSTO2\OutlookAI\WebUI\chat.js" -Pattern "tool_result|function_call|onToolResult|appendTool" | Select-Object -First 20
```

The existing path likely has something like `outlookai.appendToolResult(messageId, toolName, jsonText)` or it bakes the tool round-trip into the assistant message text. Identify the exact entry point and how it currently renders.

- [ ] **Step 2: Add `appendFileCardToMessage` and `formatBytes` helpers to `chat.js`**

Insert near the other DOM helpers in `chat.js`:

```javascript
function formatBytes(b) {
    if (b == null) return '';
    if (b < 1024) return b + ' B';
    if (b < 1024 * 1024) return (b / 1024).toFixed(1) + ' KB';
    return (b / 1024 / 1024).toFixed(1) + ' MB';
}

function formatLabel(format) {
    switch (String(format || '').toLowerCase()) {
        case 'xlsx': return 'Excel Workbook';
        case 'pdf':  return 'PDF Document';
        default:     return (format || '').toUpperCase();
    }
}

function appendFileCardToMessage(messageId, fileInfo) {
    if (!fileInfo || !fileInfo.path) return;
    var msgEl = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]');
    if (!msgEl) {
        // Fallback: append to last assistant message
        msgEl = document.querySelector('.msg-assistant:last-child');
    }
    if (!msgEl) return;

    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
        attach = document.createElement('div');
        attach.className = 'msg-attachments';
        msgEl.appendChild(attach);
    }

    var card = document.createElement('div');
    card.className = 'file-card';
    card.setAttribute('data-format', (fileInfo.format || '').toLowerCase());
    card.innerHTML =
        '<div class="file-card-icon"></div>' +
        '<div class="file-card-meta">' +
            '<div class="file-card-name" title="' + escapeAttr(fileInfo.filename || '') + '">' +
                escapeText(fileInfo.filename || '') + '</div>' +
            '<div class="file-card-sub">' + formatBytes(fileInfo.bytes) + ' &middot; ' +
                escapeText(formatLabel(fileInfo.format)) + '</div>' +
        '</div>' +
        '<div class="file-card-actions">' +
            '<button type="button" class="file-card-btn" data-action="open">Open</button>' +
            '<button type="button" class="file-card-btn" data-action="reveal">Show in folder</button>' +
        '</div>';

    card.querySelector('[data-action="open"]').addEventListener('click', function () {
        postHost('open_file', { path: fileInfo.path });
    });
    card.querySelector('[data-action="reveal"]').addEventListener('click', function () {
        postHost('reveal_in_explorer', { path: fileInfo.path });
    });

    attach.appendChild(card);
}

// Expose for C# host -> JS calls.
window.outlookai = window.outlookai || {};
window.outlookai.onFileSaved = function (messageId, fileInfo) {
    appendFileCardToMessage(messageId, fileInfo);
};

function postHost(type, payload) {
    if (!window.chrome || !window.chrome.webview) return;
    window.chrome.webview.postMessage(JSON.stringify({ type: type, payload: payload || {} }));
}

function escapeText(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escapeAttr(s) {
    return escapeText(s).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function cssEscape(s) {
    return String(s).replace(/(["\\])/g, '\\$1');
}
```

(`postHost`, `escapeText`, `escapeAttr`, `cssEscape` may already exist under different names — if so, reuse them. The pattern above is a baseline.)

- [ ] **Step 3: Hook tool-result handling**

Find the existing function that processes tool results (look for "tool_result", "function_call_output", or similar). Add:

```javascript
function handleToolResult(messageId, toolName, jsonText) {
    try {
        var obj = JSON.parse(jsonText);
        if (obj && obj.result_type === 'file_saved') {
            appendFileCardToMessage(messageId, obj);
            return; // do NOT render the raw JSON inline
        }
        if (obj && obj.error) {
            // Will be handled by error-card task; for now skip the raw JSON
            renderInlineErrorCard(messageId, obj);
            return;
        }
    } catch (_) { /* fall through to default */ }
    // ...existing render path
}
```

Wire into the existing tool-result entry point. If the codebase currently calls something like `outlookai.appendToolResult(messageId, toolName, json)`, redirect that to `handleToolResult` first.

- [ ] **Step 4: Add styles to `styles.css`**

Append to `VSTO2/OutlookAI/WebUI/styles.css`:

```css
.msg-attachments {
    margin-top: 8px;
}

.file-card {
    display: flex;
    align-items: center;
    gap: 10px;
    background: #f6f8fa;
    border: 1px solid #d0d7de;
    border-radius: 6px;
    padding: 8px 12px;
    margin: 8px 0;
    max-width: 480px;
}
.file-card-icon {
    font-size: 22px;
    line-height: 1;
    flex: 0 0 24px;
}
.file-card[data-format="pdf"] .file-card-icon::before  { content: "\1F4C4"; }   /* page-facing-up */
.file-card[data-format="xlsx"] .file-card-icon::before { content: "\1F4CA"; }   /* bar chart */
.file-card-meta { flex: 1; min-width: 0; }
.file-card-name {
    font-weight: 500;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.file-card-sub {
    font-size: 11px;
    color: #656d76;
    margin-top: 2px;
}
.file-card-actions {
    display: flex;
    gap: 4px;
}
.file-card-btn {
    background: white;
    border: 1px solid #d0d7de;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 11px;
    cursor: pointer;
    color: #1f2328;
}
.file-card-btn:hover { background: #f3f4f6; }
.file-card-btn:active { background: #ebebeb; }
```

- [ ] **Step 5: Manual smoke**

After the publish + install at the end of the plan, ask the model to "make me an Excel of my unread messages today". Confirm the FileCard renders with the correct icon, size, format label, and the Open/Show-in-folder buttons work.

- [ ] **Step 6: Full suite**

Expected: 425 (no new C# tests).

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/chat.js VSTO2/OutlookAI/WebUI/styles.css
git commit -m "feat(webui): render file-saved tool results as file card"
```

---

## Task 22: Per-message "Save as PDF" button + `handleExportPdf` flow

Adds a small `📄` button to the top-right of every assistant message, visible on hover/focus. Clicking it posts an `export_pdf` web message with the message's raw markdown; the bridge produces the file and the FileCard renders inline.

**Files:**
- Modify: `VSTO2/OutlookAI/WebUI/chat.js`
- Modify: `VSTO2/OutlookAI/WebUI/styles.css`

- [ ] **Step 1: Modify the assistant-message renderer**

Find `outlookai.appendAssistantMessage` (or equivalent) in `chat.js`. Update it to include the action button:

```javascript
outlookai.appendAssistantMessage = function (messageId, markdown) {
    var el = document.createElement('div');
    el.className = 'msg msg-assistant';
    el.dataset.messageId = messageId;
    el.dataset.state = 'streaming';
    el.innerHTML =
        '<button type="button" class="msg-action msg-action-pdf" title="Save as PDF" aria-label="Save message as PDF" tabindex="-1">' +
            '\u{1F4C4}' +   // page-facing-up emoji
        '</button>' +
        '<div class="msg-body"></div>' +
        '<div class="msg-attachments"></div>';
    el.querySelector('.msg-body').innerHTML = window.markdown.render(markdown || '');
    el.querySelector('.msg-action-pdf').addEventListener('click', function () {
        handleExportPdf(messageId);
    });
    chatLog.appendChild(el);
    return el;
};
```

(`chatLog` is the existing container element; preserve whatever name the existing code uses.)

- [ ] **Step 2: Add `handleExportPdf` and helper**

In `chat.js`:

```javascript
function deriveFilenameHint(markdown) {
    if (!markdown) return 'OutlookAI Report';
    var m = String(markdown).match(/^#{1,3}\s+(.+)$/m);
    var hint = m ? m[1] : 'OutlookAI Report';
    return hint.substring(0, 60).trim() || 'OutlookAI Report';
}

async function handleExportPdf(messageId) {
    var el = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]');
    if (!el || el.dataset.state !== 'complete') return;
    var btn = el.querySelector('.msg-action-pdf');
    if (!btn || btn.disabled) return;

    var markdown = (conversationStore && conversationStore.getMarkdown)
        ? conversationStore.getMarkdown(messageId)
        : el.querySelector('.msg-body').innerText;

    if (!markdown) return;

    btn.disabled = true;
    btn.textContent = '\u23F3'; // hourglass-flowing-sand
    try {
        postHost('export_pdf', {
            message_id: messageId,
            filename_hint: deriveFilenameHint(markdown),
            content_markdown: markdown
        });
    } finally {
        // Re-enable on response (via onFileSaved or onExportError below).
    }
}

// Re-enable the button when the file card OR error card arrives.
var _onFileSavedOriginal = window.outlookai.onFileSaved;
window.outlookai.onFileSaved = function (messageId, fileInfo) {
    if (_onFileSavedOriginal) _onFileSavedOriginal(messageId, fileInfo);
    var btn = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"] .msg-action-pdf');
    if (btn) { btn.disabled = false; btn.textContent = '\u{1F4C4}'; }
};
```

- [ ] **Step 3: Mark messages "complete" so the button is active**

Find `outlookai.finalizeAssistantMessage` and add:

```javascript
outlookai.finalizeAssistantMessage = function (messageId, opts) {
    var el = document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]');
    if (el) el.dataset.state = 'complete';
    // ...existing finalize logic
};
```

- [ ] **Step 4: Add hover/focus styles**

Append to `styles.css`:

```css
.msg-assistant { position: relative; }

.msg-action {
    position: absolute;
    top: 6px;
    right: 6px;
    opacity: 0;
    transition: opacity 120ms;
    background: transparent;
    border: 1px solid #d0d7de;
    border-radius: 4px;
    padding: 2px 6px;
    cursor: pointer;
    font-size: 14px;
    line-height: 1;
    color: #1f2328;
    z-index: 2;
}
.msg-assistant:hover .msg-action,
.msg-action:focus-visible {
    opacity: 1;
}
.msg-assistant[data-state="streaming"] .msg-action { display: none; }
.msg-action[disabled] {
    opacity: 1;
    cursor: progress;
    background: #f3f4f6;
}
```

- [ ] **Step 5: Manual smoke**

After publish + install, hover any assistant message → button appears → click → spinner → file card renders below message. Confirm `Open` works.

- [ ] **Step 6: Full suite**

Expected: 425.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/chat.js VSTO2/OutlookAI/WebUI/styles.css
git commit -m "feat(webui): add per-message Save as PDF button"
```

---

## Task 23: Render export errors as actionable error cards

When the tool or bridge returns `{ error, detail }`, render an error card inline (instead of inline JSON) with a Retry button for retryable error codes.

**Files:**
- Modify: `VSTO2/OutlookAI/WebUI/chat.js`
- Modify: `VSTO2/OutlookAI/WebUI/styles.css`

- [ ] **Step 1: Add error-card renderer**

In `chat.js`:

```javascript
var RETRYABLE_CODES = ['file_locked', 'webview2_missing', 'path_timeout', 'pdf_render_timeout'];

function renderInlineErrorCard(messageId, err) {
    var msgEl = messageId
        ? document.querySelector('[data-message-id="' + cssEscape(messageId) + '"]')
        : document.querySelector('.msg-assistant:last-child');
    if (!msgEl) return;
    var attach = msgEl.querySelector('.msg-attachments');
    if (!attach) {
        attach = document.createElement('div');
        attach.className = 'msg-attachments';
        msgEl.appendChild(attach);
    }
    var card = document.createElement('div');
    card.className = 'error-card';
    var retryable = RETRYABLE_CODES.indexOf(err.error) >= 0;
    card.innerHTML =
        '<div class="error-card-icon">\u26A0</div>' +
        '<div class="error-card-body">' +
            '<div class="error-card-title">Export failed</div>' +
            '<div class="error-card-detail">' + escapeText(err.detail || err.error) + '</div>' +
        '</div>' +
        (retryable
            ? '<div class="error-card-actions"><button type="button" class="error-card-btn" data-action="retry">Retry</button></div>'
            : '');
    if (retryable && err.error === 'webview2_missing') {
        var a = document.createElement('a');
        a.href = 'https://developer.microsoft.com/microsoft-edge/webview2/';
        a.target = '_blank';
        a.textContent = 'Install WebView2 Runtime';
        a.className = 'error-card-link';
        card.querySelector('.error-card-body').appendChild(a);
    }
    if (retryable) {
        card.querySelector('[data-action="retry"]').addEventListener('click', function () {
            handleExportPdf(messageId); // re-attempt
        });
    }
    attach.appendChild(card);

    // Re-enable the PDF button if it's still disabled.
    var btn = msgEl.querySelector('.msg-action-pdf');
    if (btn) { btn.disabled = false; btn.textContent = '\u{1F4C4}'; }
}

window.outlookai.onExportError = function (messageId, err) {
    renderInlineErrorCard(messageId, err || { error: 'unknown', detail: 'Unknown export error.' });
};
```

- [ ] **Step 2: Add error-card styles**

Append to `styles.css`:

```css
.error-card {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    background: #fff8f8;
    border: 1px solid #f3c8c8;
    border-radius: 6px;
    padding: 8px 12px;
    margin: 8px 0;
    max-width: 480px;
}
.error-card-icon {
    font-size: 18px;
    line-height: 1.2;
    color: #d1242f;
    flex: 0 0 20px;
}
.error-card-body { flex: 1; min-width: 0; }
.error-card-title {
    font-weight: 600;
    color: #d1242f;
    margin-bottom: 2px;
}
.error-card-detail {
    color: #4f3636;
    font-size: 12px;
}
.error-card-actions { display: flex; gap: 4px; align-items: center; }
.error-card-btn {
    background: white;
    border: 1px solid #d1242f;
    color: #d1242f;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 11px;
    cursor: pointer;
}
.error-card-btn:hover { background: #fff0f0; }
.error-card-link {
    display: inline-block;
    margin-top: 4px;
    font-size: 11px;
    color: #2b579a;
}
```

- [ ] **Step 3: Full suite**

Expected: 425.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/chat.js VSTO2/OutlookAI/WebUI/styles.css
git commit -m "feat(webui): render export errors as actionable error card"
```

---

## Task 24: Verify `conversationStore` retains raw markdown

For the per-message PDF button to work, the WebUI's conversation store must hold the **markdown source** of every assistant message — not the rendered HTML. If it currently stores HTML, we need to fix that.

**Files:**
- Audit: `VSTO2/OutlookAI/WebUI/chat.js`
- Audit: `VSTO2/OutlookAI/WebUI/conversation-store.js` (if it exists as a separate file)
- Modify: relevant file
- Add: regression test (JS-side via the C# unit tests is awkward; covered by smoke instead)

- [ ] **Step 1: Find the conversation store**

```powershell
Select-String -Path "VSTO2\OutlookAI\WebUI\*.js" -Pattern "conversationStore|store\.append|getMarkdown" | Select-Object -First 30
```

- [ ] **Step 2: Audit how assistant messages are stored**

If the existing code does `conversationStore.appendAssistant(htmlString)`, change it to receive markdown:

```javascript
conversationStore.appendAssistant = function (messageId, markdownSource) {
    this._messages.push({ id: messageId, role: 'assistant', markdown: markdownSource });
};

conversationStore.getMarkdown = function (messageId) {
    var m = this._messages.find(function (x) { return x.id === messageId; });
    return m ? m.markdown : null;
};
```

Adjust the appendAssistantMessage / finalize flow so the markdown source is captured (it may already be passed as a parameter from C# but discarded after rendering).

If the store already holds markdown — verify by adding a `console.log(conversationStore._messages[0])` in dev tools and confirm the `markdown` field is plain text with `#`/`|` characters, not `<h1>`/`<table>` HTML — skip the implementation.

- [ ] **Step 3: Manual smoke check**

After all changes are in, open Outlook → chat with the model → produce a markdown response → click the per-message PDF button → confirm the generated PDF has clean markdown rendering (tables, headings present and formatted). If the PDF looks "compiled HTML" rather than markdown, the store is feeding wrong content; revisit.

- [ ] **Step 4: Full suite**

Expected: 425.

- [ ] **Step 5: Commit (if changes were needed)**

```powershell
git add VSTO2/OutlookAI/WebUI/chat.js   # or conversation-store.js
git commit -m "fix(store): preserve raw markdown for assistant messages"
```

Skip this step if the audit showed no fix was needed.

---

## Task 25: Mark Phase 5 implemented in docs + final full Debug test run

**Files:**
- Modify: `docs/superpowers/specs/2026-05-18-phase-5-exports-design.md`
- Modify: `docs/superpowers/plans/2026-05-18-phase-5-exports.md` (this file)

- [x] **Step 1: Update spec status**

Changed the spec header's `Status:` line from `Proposed (pending user review)` to `Implemented (Debug tests pass; install/smoke pending)`.

- [x] **Step 2: Record current plan progress without marking pending release/smoke work complete**

Tasks 1-24 are implemented and reviewed. This plan intentionally leaves Tasks 26-28 unchecked because publish/install, live smoke, and push are still pending.

- [x] **Step 3: Final full test run**

```powershell
node --check VSTO2/OutlookAI/WebUI/chat.js
node --check VSTO2/OutlookAI/WebUI/markdown.js
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Passed: 499` (actual count before Task 25 docs update was 499/499; rerun after docs changes and record below).

Result:
- `node --check VSTO2/OutlookAI/WebUI/chat.js` passed.
- `node --check VSTO2/OutlookAI/WebUI/markdown.js` passed.
- MSBuild Debug Any CPU succeeded with existing `MSB3277` warnings.
- VSTest reported `499/499` passed (`Total tests: 499`, `Passed: 499`).

- [x] **Step 4: Commit docs**

```powershell
git add docs/superpowers/specs/2026-05-18-phase-5-exports-design.md docs/superpowers/plans/2026-05-18-phase-5-exports.md
git commit -m "docs(superpowers): mark Phase 5 exports implemented"
```

---

## Task 26: Publish Release + install elevated

**Files:** none (deployment only — no commit).

- [ ] **Step 1: Confirm Outlook is closed**

```powershell
$procs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) { "Outlook RUNNING pid=$($procs[0].Id) - close before install" } else { "Outlook closed" }
```

- [ ] **Step 2: Publish Release into staging**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" /p:PublishDir="$staging\" /v:minimal /nologo
```

Expected: `OutlookAI -> ...\bin\Release\OutlookAI.dll`.

- [ ] **Step 3: Copy installer + run elevated**

```powershell
Copy-Item -LiteralPath "Deploy\Install-OutlookAI.ps1" -Destination "$staging\" -Force
$script = Join-Path $staging "Install-OutlookAI.ps1"
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -SourcePath `"$staging`""
$proc = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -Wait -PassThru
"installer exit=$($proc.ExitCode)"
```

- [ ] **Step 4: Verify hash match**

```powershell
$staged = "$staging\OutlookAI.dll"
$installed = "C:\Program Files\OutlookAI\OutlookAI.dll"
$a = (Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash
$b = (Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash
"match=$($a -eq $b) hash=$a"
```

Expected: `match=True`.

---

## Task 27: End-to-end smoke

**Files:** none — manual interaction only.

- [ ] **Step 1: Open Outlook**

Wait for the add-in to load (taskpane buttons appear in the ribbon).

- [ ] **Step 2: Smoke A — model-initiated Excel**

Open Inbox Copilot (or Reports). Ask:

> Make me an Excel of my unread messages from today.

Expected:
- Model dispatches `outlook_search_messages` (or `outlook_count_messages`) then `outlook_export_excel`.
- A file card appears with a `.xlsx` filename, the file size (e.g. "8 KB · Excel Workbook"), and Open/Show-in-folder buttons.
- Clicking `Open` launches Excel with: row 1 = bold gray-filled header (Date, Subject, Sender, etc.), frozen at row 1, autofilter dropdowns visible, date column formatted, currency formatted if used.

- [ ] **Step 3: Smoke B — model-initiated PDF**

Continue the chat:

> Now save that as a polished PDF.

Expected:
- Model dispatches `outlook_export_pdf` with markdown content + title + subtitle.
- A file card appears with `.pdf` filename.
- Clicking `Open` shows a clean A4 PDF with: header bar in OutlookAI navy, title + subtitle, "Generated by OutlookAI ·" timestamp, the table preserved (header repeats across pages if multi-page), no chat-UI chrome.

- [ ] **Step 4: Smoke C — per-message PDF button**

Find any arbitrary assistant message in the chat. Hover. The 📄 button appears top-right. Click.

Expected:
- Brief spinner state on the button.
- File card appears inline below the message body.
- Clicking `Open` shows the rendered PDF using the SAME markdown the chat displayed (no fresh polish from the model).

- [ ] **Step 5: Smoke D — path-traversal rejection**

Open DevTools in the WebView (if dev mode is on) or temporarily enable them. Run:

```javascript
chrome.webview.postMessage(JSON.stringify({
    type: 'open_file',
    payload: { path: 'C:\\Windows\\System32\\cmd.exe' }
}));
```

Expected:
- No process launches.
- An error card appears with `unauthorized_path` code.
- Trace log records `open_file REJECTED`.

- [ ] **Step 6: Inspect trace log**

```powershell
Get-Content -LiteralPath "C:\Users\MDASR\AppData\Local\OutlookAI\trace.log" -Tail 100 | Select-String -Pattern "export\.|ExportBridge" | Select-Object -Last 30
```

Expected: clean `export.*` events for all three smoke flows, no exceptions.

---

## Task 28: Push branch

- [ ] **Step 1: Push**

```powershell
git push origin feature/codex-oauth-migration
```

Expected: push succeeds; remote updates with all Phase 5 commits.

---

## Self-Review Checklist (run before declaring Phase 5 complete)

**1. Spec coverage:** Walk through every section of `docs/superpowers/specs/2026-05-18-phase-5-exports-design.md`. Each requirement should map to at least one task above.

- Excel tool: Tasks 4, 5, 6, 7, 8, 9, 10 ✓
- PDF tool: Tasks 11, 12, 13, 14, 15, 16, 17, 18, 19 ✓
- Filename + path: Tasks 2, 3 ✓
- Per-message button: Task 22 ✓
- FileCard UI: Task 21 ✓
- Error UI: Task 23 ✓
- Bridge security: Tasks 3 (policy) + 20 (bridge enforcement) ✓
- Conversation store: Task 24 ✓
- Tracing: covered by `TraceLog.Write` calls in each implementation step ✓

**2. Placeholder scan:** Re-read this plan top to bottom. No `TODO`, no "handle errors appropriately", no "similar to Task N" without inline code. (If any are found, fix inline before starting execution.)

**3. Type consistency:** Method signatures used across tasks match. `FileSavedResult` shape used in Task 7 matches Task 17, 20, 21. `ExportPdfArgs` properties (`FilenameHint`, `Title`, `Subtitle`, `ContentMarkdown`) consistent between Tasks 7, 15, 16, 17, 20.

**4. Scope check:** Single plan, ~24 commits. Comparable to Phase 4 (17 commits). Manageable in one focused session.

---

## Final State After Phase 5

- 24 new commits on `feature/codex-oauth-migration`.
- Test count: **317 → 499** (182 new tests before Task 25 docs-only verification).
- Two new model tools: `outlook_export_excel`, `outlook_export_pdf`.
- One new UI affordance: per-message "Save as PDF" button.
- One new shared infra: `markdown.js` (used by chat AND PDF).
- New helper subsystems under `Services/Export/`.
- One file-card UX pattern used by both tools and the button.
- One off-screen WebView2 instance for headless PDF rendering.
- ClosedXML pinned 0.102.x.
- Reports saved to `~\Documents\OutlookAI\Reports\` with auto-generated timestamped filenames.
- Path-policy security gate on all file-action bridge methods.

Phase 5 implementation is ready for publish/install and live smoke. No dependencies on Phase 6+ items (settings UI, multi-sheet Excel, PDF page numbers, etc.).
