# Issues #2–#5 Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Resolve the four outstanding follow-up issues on `master` in one PR — fix the UNC file-launch bug (#2), tighten tool-schema guidance to keep large Excel exports inside the context window (#3, #4), and remove the dead `Config.MaxTokens` setting (#5).

**Architecture:** Four independent surgical changes plus a verification gate, a Release publish + install, smoke, and a single PR merged into `master`. Tasks 1–4 each ship as their own commit so the diff stays auditable. No new services, no schema migrations, no new dependencies.

**Tech Stack:** C# .NET Framework 4.7.2 (VSTO add-in), Newtonsoft.Json schema builder, xUnit tests, PowerShell installer.

**Source issues:**
- #2 — https://github.com/kirklandsig/OutlookAI/issues/2 — `File card "Open" can fail with "UNC paths are not supported"`
- #3 — https://github.com/kirklandsig/OutlookAI/issues/3 — `Steer model toward metadata-only synthesis for tabular Excel exports`
- #4 — https://github.com/kirklandsig/OutlookAI/issues/4 — `Add paginated bulk-extract pattern for large Excel exports that need bodies`
- #5 — https://github.com/kirklandsig/OutlookAI/issues/5 — `Remove unused Config.MaxTokens` (Option 1, recommended)

---

## File Structure

**Modified files:**

- `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs` — set `ProcessStartInfo.WorkingDirectory` to a local temp path before `UseShellExecute = true` in `OpenWithDefaultApp` and `RevealWithExplorer`. Task 1.
- `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` — extend description strings for `outlook_search_messages`, `outlook_read_messages`, `outlook_export_excel` to steer toward metadata-only synthesis (Task 2) and to teach the date-windowed pagination pattern (Task 3).
- `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs` — pin new guidance strings. Tasks 2 and 3.
- `VSTO2/OutlookAI/Config.cs` — delete `DefaultMaxTokens`, `MaxTokens`, the load logic, and the reset assignment. Update the comment at the top of the file. Task 4.
- `VSTO2/OutlookAI.Tests/ConfigTests.cs` — drop the four references to `MaxTokens` (asserting default 65536, and the two `<MaxTokens>` strings inside config XML fixtures). Task 4.
- `Deploy/Install-OutlookAI.ps1` — drop `<MaxTokens>65536</MaxTokens>` from the default config written by the script, and remove `MaxTokens` from the docstring at the top. Task 4.
- `Deploy/README.txt` — drop `MaxTokens` from the "Server-authoritative defaults" sentence. Task 4.

**Files NOT touched:**

- `README.md` — no `MaxTokens` references.
- `docs/superpowers/specs/*` and `docs/superpowers/plans/*` — historical record; left alone deliberately.
- `Services/CodexChatService.cs` — never referenced `MaxTokens` to begin with.
- Service-layer surfaces, parsers, tool implementations — no functional changes for #3 or #4 (description-only).
- `handoff.md` — gitignored; updated locally after merge in Task 6.

**Verification commands** (run from
`C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-issues-2-5-cleanup`):

```powershell
node --check VSTO2\OutlookAI\WebUI\chat.js
node --check VSTO2\OutlookAI\WebUI\markdown.js
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Baseline before starting: **546 tests passing**. Target at end: **≥ 546** (tasks add tests, may push the count up by ~6–8; do not assert a precise number unless we end up wanting to).

---

## Task 1: Fix UNC path failure in `ExportBridge` (Issue #2)

`Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })` can throw `UNC paths are not supported` when Windows tries to implicitly set the spawned process working directory to a UNC path (Folder Redirection on `~\Documents` → `\\fileserver\…`). Fix is to set `WorkingDirectory` to a known local path before launching.

Same logic applies to `explorer.exe /select,"…"` in `RevealWithExplorer`.

**Files:**
- Modify: `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs:136-147`
- Modify: `VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs` (add source-level assertions)

- [ ] **Step 1: Add failing source-level tests**

Append two new tests to `VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs`. These read the source file as text and pin the fix patterns — exact same style as the existing `Controllers_ObserveEntireAsyncHostMessageHandler` tests in that file.

```csharp
[Fact]
public void OpenWithDefaultApp_SetsLocalWorkingDirectory_ToAvoidUncFailure()
{
    var source = File.ReadAllText(LocateExportBridgeSource());
    Assert.Contains("OpenWithDefaultApp", source);
    Assert.Contains("WorkingDirectory = Path.GetTempPath()", source);
}

[Fact]
public void RevealWithExplorer_SetsLocalWorkingDirectory_ToAvoidUncFailure()
{
    var source = File.ReadAllText(LocateExportBridgeSource());
    Assert.Contains("RevealWithExplorer", source);
    Assert.Contains("WorkingDirectory = Path.GetTempPath()", source);
}

private static string LocateExportBridgeSource()
{
    var baseDir = AppContext.BaseDirectory;
    var dir = new DirectoryInfo(baseDir);
    while (dir != null)
    {
        var candidate = Path.Combine(dir.FullName, "VSTO2", "OutlookAI", "TaskPane", "Chat", "ExportBridge.cs");
        if (File.Exists(candidate)) return candidate;
        dir = dir.Parent;
    }
    throw new FileNotFoundException("ExportBridge.cs not found relative to test base dir.");
}
```

If `ExportBridgeTests.cs` does not already `using System.IO;` and `using System;`, add them.

- [ ] **Step 2: Run the new tests to confirm RED**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /p:Platform="AnyCPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SetsLocalWorkingDirectory"
```

Expected: both tests **fail** with `Assert.Contains() Failure ... Not found: "WorkingDirectory = Path.GetTempPath()"`.

- [ ] **Step 3: Implement the fix**

Replace the two private methods at `VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs:136-147` so they set a local working directory:

```csharp
private static void OpenWithDefaultApp(string path)
{
    Process.Start(new ProcessStartInfo(path)
    {
        UseShellExecute = true,
        // Force a local working directory so ShellExecute does not try to set
        // CWD to a UNC path on Folder-Redirected Documents. The launched
        // process inherits this CWD; the target file path is unaffected.
        WorkingDirectory = Path.GetTempPath(),
    });
}

private static void RevealWithExplorer(string path)
{
    Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + path.Replace("\"", "\\\"") + "\"")
    {
        UseShellExecute = true,
        WorkingDirectory = Path.GetTempPath(),
    });
}
```

Confirm `using System.IO;` is present in `ExportBridge.cs` so `Path.GetTempPath()` resolves; add it if missing.

- [ ] **Step 4: Run the new tests to confirm GREEN**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SetsLocalWorkingDirectory"
```

Expected: both tests pass.

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/Chat/ExportBridge.cs `
        VSTO2/OutlookAI.Tests/TaskPane/Chat/ExportBridgeTests.cs
git commit -m "fix(export): avoid UNC working-directory failure when opening reports"
```

---

## Task 2: Steer model toward metadata-only synthesis (Issue #3)

Three description tweaks in `ToolCatalogSchema.cs` that nudge the model to build tabular exports from search snippets instead of full bodies.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` (the `outlook_search_messages`, `outlook_read_messages`, `outlook_export_excel` description strings)
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs` (pin new strings)

- [ ] **Step 1: Add failing schema-guidance tests**

Append to `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void SearchMessages_Description_SteersTowardSnippetForBulkExports()
{
    var description = GetToolDescription("outlook_search_messages");
    Assert.Contains("snippet", description, StringComparison.OrdinalIgnoreCase);
    Assert.Contains("metadata-only", description, StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void ReadMessages_Description_WarnsAgainstLargeBatchesForExports()
{
    var description = GetToolDescription("outlook_read_messages");
    Assert.Contains("metadata", description, StringComparison.OrdinalIgnoreCase);
    Assert.Contains("snippet", description, StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void ExportExcel_Description_PrefersSearchSnippetsOverBodyReads()
{
    var description = GetToolDescription("outlook_export_excel");
    Assert.Contains("metadata-only", description, StringComparison.OrdinalIgnoreCase);
    Assert.Contains("snippet", description, StringComparison.OrdinalIgnoreCase);
}
```

If `GetToolDescription(string toolName)` does not already exist as a helper in this test file, add it once at the bottom of the class — it reads the catalog JSON and returns the matching tool's `description` field:

```csharp
private static string GetToolDescription(string toolName)
{
    var catalog = ToolCatalogSchema.Build();
    var match = catalog
        .OfType<Newtonsoft.Json.Linq.JObject>()
        .FirstOrDefault(t => (string)t["function"]?["name"] == toolName);
    Assert.NotNull(match);
    return (string)match["function"]["description"];
}
```

(If `ToolCatalogSchema.Build()` returns a different shape, adapt accordingly — there is precedent in this file. Match the existing helper if one is already present.)

- [ ] **Step 2: Run the new tests to confirm RED**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SteersTowardSnippet|FullyQualifiedName~WarnsAgainstLargeBatches|FullyQualifiedName~PrefersSearchSnippets"
```

Expected: all three new tests **fail** with `Assert.Contains() Failure`.

- [ ] **Step 3: Update `outlook_search_messages` description**

In `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`, find the `outlook_search_messages` BuildToolEntry (around line 33–55). Append two sentences to the description string, just before `"Prefer one precise call over many. After search, use outlook_read_message on the most-relevant id for full body."`:

```csharp
+ "For tabular Excel exports over many messages, prefer metadata-only synthesis: the snippet field already contains the first ~200 chars of each match, so build Excel rows directly from search results and only call outlook_read_messages when a column genuinely requires full body content. "
+ "Reading full bodies for 100+ messages will exceed the context window. "
```

- [ ] **Step 4: Update `outlook_read_messages` description**

Find the `outlook_read_messages` BuildToolEntry (around line 141). Append to the description:

```csharp
+ " Avoid calling on more than ~25 IDs when each body is long. For large result sets you intend to export, work from search metadata + snippet instead of bulk-reading bodies."
```

- [ ] **Step 5: Update `outlook_export_excel` description**

Find the `outlook_export_excel` BuildToolEntry (around line 202). Append to the description:

```csharp
+ " For exports over 50+ messages, prefer projecting columns (subject, from, to, received_at, snippet) from outlook_search_messages results without reading full bodies; this is metadata-only synthesis and stays within the context window. Reading full bodies for 100+ messages will exceed context."
```

- [ ] **Step 6: Run the new tests to confirm GREEN**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~SteersTowardSnippet|FullyQualifiedName~WarnsAgainstLargeBatches|FullyQualifiedName~PrefersSearchSnippets"
```

Expected: all three new tests pass.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs `
        VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(schema): steer model toward metadata-only synthesis for Excel exports"
```

---

## Task 3: Add paginated bulk-extract guidance (Issue #4)

When body content genuinely is required for many messages, the model should batch by date window. Add one paragraph each to `outlook_search_messages` and `outlook_export_excel`.

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`

- [ ] **Step 1: Add failing pagination-guidance tests**

Append to `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void SearchMessages_Description_TeachesDateWindowedPagination()
{
    var description = GetToolDescription("outlook_search_messages");
    Assert.Contains("page by date window", description, StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void ExportExcel_Description_MentionsPaginatedExtractionForLargeSets()
{
    var description = GetToolDescription("outlook_export_excel");
    Assert.Contains("page by date window", description, StringComparison.OrdinalIgnoreCase);
}
```

- [ ] **Step 2: Run the new tests to confirm RED**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~TeachesDateWindowedPagination|FullyQualifiedName~MentionsPaginatedExtraction"
```

Expected: both new tests **fail**.

- [ ] **Step 3: Update `outlook_search_messages` description**

Append to the description string (after the metadata-only sentences added in Task 2):

```csharp
+ "When body content really is required for an Excel/PDF report over many matches, page by date window: do separate searches per quarter / month, read up to 25 bodies per window, extract rows, accumulate across turns, then call the export tool once at the end. "
```

- [ ] **Step 4: Update `outlook_export_excel` description**

Append to the description string (after the metadata-only sentence added in Task 2):

```csharp
+ " For exports that genuinely require body-derived columns over a large result set, page by date window across multiple turns (search Q1 -> read 25 -> extract rows; search Q2 -> read 25 -> extract rows; ... -> outlook_export_excel once with all accumulated rows)."
```

- [ ] **Step 5: Run the new tests to confirm GREEN**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" `
  /TestCaseFilter:"FullyQualifiedName~TeachesDateWindowedPagination|FullyQualifiedName~MentionsPaginatedExtraction"
```

Expected: both new tests pass.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs `
        VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs
git commit -m "feat(schema): teach paginated bulk-extract pattern for large exports"
```

---

## Task 4: Remove unused `Config.MaxTokens` (Issue #5, Option 1)

Delete the dead config field plus its loader, fix tests, fix the installer default config, fix `Deploy/README.txt`. Existing v2 `<MaxTokens>` elements in user configs continue to load fine — the unknown element is silently ignored, same as any v1 field today.

**Files:**
- Modify: `VSTO2/OutlookAI/Config.cs`
- Modify: `VSTO2/OutlookAI.Tests/ConfigTests.cs`
- Modify: `Deploy/Install-OutlookAI.ps1`
- Modify: `Deploy/README.txt`

- [ ] **Step 1: Update `ConfigTests.cs` to remove `MaxTokens` assertions**

`VSTO2/OutlookAI.Tests/ConfigTests.cs` currently has these references at lines 27, 39, 44, 51:

```csharp
Assert.Equal(65536, Config.MaxTokens);   // line 27 (defaults test)
...
+ "<MaxTokens>65536</MaxTokens>"          // line 39 (global config XML fixture)
+ "<MaxTokens>2048</MaxTokens>"           // line 44 (per-user override fixture)
...
Assert.Equal(65536, Config.MaxTokens);    // line 51 (override test asserting user MaxTokens is rejected)
```

Remove all four. The "per-user override is rejected" assertion at line 51 should be replaced with a comment / a still-meaningful assertion that some other server-authoritative field (e.g. `Model`) is NOT overridden by the user fixture; the existing test most likely already asserts that — leave the existing non-MaxTokens assertions in place.

After edit, the tests in question should still build and pass against the (about-to-be-modified) Config.cs.

- [ ] **Step 2: Run the modified tests to confirm RED**

Build will fail until `Config.MaxTokens` is also deleted, so we cannot run tests yet — that is the expected RED for the field removal. Continue to Step 3.

- [ ] **Step 3: Remove `Config.MaxTokens` from `Config.cs`**

In `VSTO2/OutlookAI/Config.cs`:

- Line 13 (comment): change `Server-authoritative fields (CodexAuthPath, Model, MaxTokens)` to `Server-authoritative fields (CodexAuthPath, Model)`.
- Line 23: delete `public const int DefaultMaxTokens = 65536;`.
- Line 31: delete `public static int MaxTokens { get; set; } = DefaultMaxTokens;`.
- Line 169: in `ResetDefaults()`, delete `MaxTokens = DefaultMaxTokens;`.
- Lines 263-267: delete the entire `var maxTokens = root.Element("MaxTokens"); …` block.

- [ ] **Step 4: Update `Deploy/Install-OutlookAI.ps1`**

Two places:

1. The docstring near the top (around line 15): change `Model / MaxTokens / CodexAuthPath settings` to `Model / CodexAuthPath settings`.
2. The default global config the installer writes (around line 375): delete the line `  <MaxTokens>65536</MaxTokens>`.

If there are any inline comments referencing `MaxTokens`, update them to match.

- [ ] **Step 5: Update `Deploy/README.txt`**

Around line 40 the file reads: `Server-authoritative defaults: Model, MaxTokens, CodexAuthPath.` Change to `Server-authoritative defaults: Model, CodexAuthPath.`

- [ ] **Step 6: Build and run full VSTest suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: build succeeds (only the pre-existing `MSB3277` warning), full VSTest reports `Total tests: ≥ 546` and **0 failures**. The new tests from Tasks 1–3 are included; the `MaxTokens` assertions are gone.

- [ ] **Step 7: Commit**

```powershell
git add VSTO2/OutlookAI/Config.cs `
        VSTO2/OutlookAI.Tests/ConfigTests.cs `
        Deploy/Install-OutlookAI.ps1 `
        Deploy/README.txt
git commit -m "chore(config): remove unused MaxTokens setting"
```

---

## Task 5: Verification gate

Confirm the cumulative change is clean before publish/install.

**Files:** none (verification only).

- [ ] **Step 1: WebUI syntax checks**

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

Expected: build succeeds; only the pre-existing `MSB3277` warning is emitted.

- [ ] **Step 3: Full VSTest suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" `
  "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: `Total tests: ≥ 546   Passed: same number   Failed: 0`.

- [ ] **Step 4: Confirm working tree clean**

```powershell
git status --short
```

Expected: empty output.

- [ ] **Step 5: Confirm no stray `MaxTokens` references in shipped code**

```powershell
Select-String -Path "VSTO2\OutlookAI\Config.cs","VSTO2\OutlookAI.Tests\ConfigTests.cs","Deploy\Install-OutlookAI.ps1","Deploy\README.txt","README.md" -Pattern "MaxTokens"
```

Expected: empty output. (Historical specs/plans under `docs/superpowers/` are excluded on purpose.)

---

## Task 6: Publish Release, install, smoke, PR, merge

**Files:** none (deployment + GitHub only — no new commits on the branch beyond Tasks 1-4).

- [ ] **Step 1: Confirm Outlook is closed**

```powershell
$procs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) { "Outlook RUNNING pid=$($procs[0].Id) - close before install" } else { "Outlook closed" }
```

Expected: `Outlook closed`.

- [ ] **Step 2: Publish Release into staging**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
  /p:PublishDir="$staging\" /v:minimal /nologo
```

Expected: `OutlookAI -> …\bin\Release\OutlookAI.dll`.

- [ ] **Step 3: Copy installer and run elevated**

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
Copy-Item -LiteralPath "Deploy\Install-OutlookAI.ps1" -Destination "$staging\" -Force
$script = Join-Path $staging "Install-OutlookAI.ps1"
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -SourcePath `"$staging`""
"launching elevated installer"
$proc = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -Wait -PassThru
"installer exit=$($proc.ExitCode)"
```

User must approve UAC. Expected exit code 0.

- [ ] **Step 4: Verify hash match**

```powershell
$staging   = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
$staged    = (Get-FileHash -LiteralPath "$staging\OutlookAI.dll"   -Algorithm SHA256).Hash
$installed = (Get-FileHash -LiteralPath "C:\Program Files\OutlookAI\OutlookAI.dll" -Algorithm SHA256).Hash
"match=$($staged -eq $installed) staged=$staged installed=$installed"
```

Expected: `match=True`.

- [ ] **Step 5: Smoke**

Open Outlook → AI Assistant → at minimum:

1. Generate one Excel export or per-message PDF (any tool result that surfaces a file card).
2. Click **Open** on the file card → file opens in its default app, no `UNC paths are not supported` error.
3. Click **Show in folder** on the file card → Explorer opens with the file selected.

If the dev machine does **not** have Folder Redirection configured (most likely), the UNC bug is not reproducible here. The hash match plus the local-WorkingDirectory source-test in Task 1 are the verifiable evidence; record this in the PR comment.

- [ ] **Step 6: Push the branch**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git push -u origin feature/issues-2-5-cleanup
```

- [ ] **Step 7: Create the PR via `--body-file`** (PowerShell gotcha — see `handoff.md` §11)

```powershell
$tmp = New-TemporaryFile
Set-Content -LiteralPath $tmp.FullName -Encoding UTF8 -Value @'
## Summary
Clears the four follow-up issues filed after the v2 launch.

- #2 (`bug`) — `ExportBridge.OpenWithDefaultApp` / `RevealWithExplorer` now set `ProcessStartInfo.WorkingDirectory = Path.GetTempPath()` so Folder-Redirected Documents (UNC-backed) no longer break `Open` / `Show in folder`.
- #3 (`enhancement`) — Tool schema steers the model toward metadata-only synthesis for tabular Excel exports (build rows from `outlook_search_messages` snippets instead of bulk-reading bodies).
- #4 (`enhancement`) — Tool schema teaches the date-windowed pagination pattern for exports that genuinely need body content over many messages.
- #5 (`bug`) — Removes the dead `Config.MaxTokens` setting (Option 1 in the issue). `Config.cs`, `ConfigTests.cs`, `Install-OutlookAI.ps1`, and `Deploy/README.txt` all updated. Existing user configs with `<MaxTokens>` continue to load (element silently ignored).

## Test Plan
- `node --check VSTO2\OutlookAI\WebUI\chat.js` and `markdown.js` both pass.
- VS MSBuild Debug Any CPU succeeded; only the pre-existing `MSB3277` warning emitted.
- Full VSTest suite green (count includes new tests from Tasks 1-3).
- Release publish + elevated install succeeded; staged and installed `OutlookAI.dll` SHA256 match.
- Smoke: file card `Open` and `Show in folder` both work on a non-UNC Documents folder. UNC repro requires Folder Redirection; not testable on the dev machine. The new source-level test pins the `WorkingDirectory = Path.GetTempPath()` fix as evidence.

Closes #2, #3, #4, #5.
'@
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr create --repo kirklandsig/OutlookAI --base master --head feature/issues-2-5-cleanup `
    --title "Resolve issues #2-#5 (UNC fix, schema steering, MaxTokens removal)" `
    --body-file $tmp.FullName
Remove-Item -LiteralPath $tmp.FullName -Force
```

Expected: the command prints the new PR URL.

- [ ] **Step 8: Merge the PR**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr view --json mergeStateStatus,mergeable
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh pr merge --merge --subject "Merge: resolve issues #2-#5"
```

Expected: `mergeable: MERGEABLE`, then `Merged pull request`.

- [ ] **Step 9: Fast-forward local `master` in the main worktree**

```powershell
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" fetch origin master
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" pull --ff-only origin master
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" log -1 --oneline
```

Expected: latest commit on `master` is the merge.

- [ ] **Step 10: Confirm issues auto-closed**

```powershell
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
gh issue list --repo kirklandsig/OutlookAI --state open --json number,title
```

Expected: issues #2-#5 are not in the open list (`Closes #2, #3, #4, #5` in the PR body should have closed them on merge).

- [ ] **Step 11: Clean up the feature branch and worktree**

From the main worktree directory:

```powershell
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" worktree remove "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-issues-2-5-cleanup"
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" worktree prune
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" branch -d feature/issues-2-5-cleanup
Remove-Item Env:GITHUB_TOKEN -ErrorAction SilentlyContinue
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" push origin --delete feature/issues-2-5-cleanup
git -C "C:\Users\MDASR\Desktop\Projects\OutlookAI" branch -a
```

Expected: only `master` and `remotes/origin/master` remain.

- [ ] **Step 12: Update `handoff.md`** (local-only; gitignored)

In `C:\Users\MDASR\Desktop\Projects\OutlookAI\handoff.md`, in §13 "Open follow-ups / known gaps" → "Open GitHub issues":

- Remove the bullets for #2, #3, #4, #5 (now closed).
- In §11 PowerShell gotchas, the `gh issue create --body $var` note stays (still relevant).
- In §1 snapshot, refresh the merge commit hash and the installed DLL hash.
- In §12 "Recent history", append a one-line entry: "Resolved issues #2-#5 (UNC fix in `ExportBridge`, two schema-guidance changes, removal of unused `Config.MaxTokens`)."

This is the only post-merge maintenance step. Handoff is not committed.

---

## Self-Review

**1. Spec coverage:**

| Issue | Tasks |
|---|---|
| #2 UNC paths | Task 1 |
| #3 metadata-only synthesis | Task 2 |
| #4 paginated extraction | Task 3 |
| #5 MaxTokens removal | Task 4 |
| Verification | Task 5 |
| Publish, install, smoke, PR, merge, cleanup | Task 6 |

Every issue has at least one task; every task closes one issue or runs a required gate.

**2. Placeholder scan:** No `TBD`, `TODO`, `implement later`, or `similar to Task N` phrases. All file contents are inlined verbatim.

**3. Type / name consistency:**

- `Path.GetTempPath()` used in both `OpenWithDefaultApp` and `RevealWithExplorer` (Task 1).
- `GetToolDescription(string toolName)` helper introduced once in Task 2 Step 1 and reused in Task 3 Step 1.
- `outlook_search_messages`, `outlook_read_messages`, `outlook_export_excel` tool names consistent across Tasks 2 and 3.
- `Config.MaxTokens`, `DefaultMaxTokens`, `<MaxTokens>` consistently removed in Task 4 — no stragglers in `Config.cs`, `ConfigTests.cs`, `Install-OutlookAI.ps1`, `Deploy/README.txt`.

---

## Final State After This Plan

- `master` advances by one merge commit ahead of `84da4ee`.
- Issues #2, #3, #4, #5 closed on GitHub.
- No new test failures; baseline grows by ~7 tests (2 for #2, 3 for #3, 2 for #4).
- Release DLL re-published and installed; SHA256 verified.
- `feature/issues-2-5-cleanup` branch and worktree gone.
- `handoff.md` reflects the new state.
