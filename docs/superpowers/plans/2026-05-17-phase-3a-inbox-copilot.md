# Phase 3a Implementation Plan — Inbox Copilot

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the Inbox Copilot chat pane that opens from the Outlook Explorer ribbon, anchored per-Explorer, fed by the current folder + selection, with quick-action chips and a substantially enhanced `outlook_search_messages` so the model can express precise mailbox queries in a single call.

**Architecture:** Reuses every piece of Phase 2 (WebView2 surface, `chat.js`, `ConversationStore`, `ChatEventSink`, `CodexChatService.RunTurnAsync`, `OutlookThreadMarshaller`, `OutlookToolHost`) and adds one new tool (`outlook_get_current_selection`), one new pane (`InboxCopilotPane`) with a thin controller (`InboxCopilotController`), one ribbon group on `TabMail`, structured fields on `outlook_search_messages` + `outlook_count_messages`, and small additions to `chat.js` for quick-action chips and an enhanced context strip.

**Tech Stack:** C# 7.3 / .NET Framework 4.7.2, VSTO Outlook add-in, WinForms + WebView2 (Microsoft.Web.WebView2 1.0.2849.39), Newtonsoft.Json 13, xUnit 2.9.3 / Microsoft.NET.Test.Sdk 17.14.1, MSBuild 18, Outlook 16 COM (`Microsoft.Office.Interop.Outlook`).

**Test commands** (used at each task's verification step):

```powershell
# Restore (only when packages.config or csproj changes)
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj

# Build
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal

# Run all tests
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"

# Run a filtered subset
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SomePattern"
```

**Push checkpoints:** after Task 4 (enhanced search shipped), Task 9 (selection tool shipped), Task 13 (Inbox Copilot pane functional end-to-end), Task 15 (final).

**Baseline test count at HEAD `25ab2d7`:** 109 / 109. Target after Phase 3a: ~125.

---

## Task 1: Extend `SearchMessagesArgs` with the new structured fields

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`

POCO-only change. New fields default to null/unset so existing call sites keep working unchanged. No tests yet — those land in Task 2 once the consumers can read the fields.

- [ ] **Step 1: Extend `SearchMessagesArgs`**

Locate `public sealed class SearchMessagesArgs` in `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`. Replace the whole class declaration with:

```csharp
public sealed class SearchMessagesArgs
{
    public string Query { get; set; }
    public string From { get; set; }
    public string SubjectContains { get; set; }
    public string BodyContains { get; set; }
    public bool? HasAttachment { get; set; }
    public bool? IsUnread { get; set; }
    public bool? IsFlagged { get; set; }
    /// <summary>One of "low" | "normal" | "high"; null = unset.</summary>
    public string Importance { get; set; }
    public string FolderId { get; set; }
    public DateTimeOffset? DateFrom { get; set; }
    public DateTimeOffset? DateTo { get; set; }
    public int MaxResults { get; set; } = 25;
}
```

- [ ] **Step 2: Build**

Run the build command from the plan header. Expected: clean build, no warnings.

- [ ] **Step 3: Run the full suite to confirm no regressions**

Run the test command from the plan header. Expected: 109/109 pass.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs
git commit -m "Phase 3a Task 1: extend SearchMessagesArgs with structured filter fields"
```

---

## Task 2: Expose `BuildRestrictFilter` for unit testing, add the DASL clauses

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Tools\BuildRestrictFilterTests.cs`

The existing `BuildRestrictFilter` is `private static`. Make it `internal static` so the test assembly can call it directly. Then TDD the new DASL clauses one at a time.

Test assembly already has `InternalsVisibleTo` access by default for SDK-style xUnit projects? **No** — net472 / VSTO classic csproj requires an explicit `[assembly: InternalsVisibleTo]`. Add it.

- [ ] **Step 1: Add `InternalsVisibleTo` to the product assembly**

Open `VSTO2\OutlookAI\Properties\AssemblyInfo.cs`. Append at the end:

```csharp
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("OutlookAI.Tests")]
```

- [ ] **Step 2: Change `BuildRestrictFilter` from private to internal**

In `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`, find the line:

```csharp
private static string BuildRestrictFilter(SearchMessagesArgs args)
```

Change to:

```csharp
internal static string BuildRestrictFilter(SearchMessagesArgs args)
```

- [ ] **Step 3: Write the failing tests**

Create `VSTO2\OutlookAI.Tests\Services\Tools\BuildRestrictFilterTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class BuildRestrictFilterTests
    {
        [Fact]
        public void EmptyArgs_ReturnsNullFilter()
        {
            var filter = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs());
            Assert.Null(filter);
        }

        [Fact]
        public void QueryOnly_BuildsSubjectOrBodyLike()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Query = "Q4" });
            Assert.StartsWith("@SQL=", f);
            Assert.Contains("urn:schemas:httpmail:subject LIKE '%Q4%'", f);
            Assert.Contains("urn:schemas:httpmail:textdescription LIKE '%Q4%'", f);
        }

        [Fact]
        public void From_MatchesDisplayNameOrEmail()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { From = "jane" });
            Assert.Contains("urn:schemas:httpmail:fromname LIKE '%jane%'", f);
            Assert.Contains("urn:schemas:httpmail:fromemail LIKE '%jane%'", f);
        }

        [Fact]
        public void SubjectContains_AddsSubjectLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { SubjectContains = "plan" });
            Assert.Contains("urn:schemas:httpmail:subject LIKE '%plan%'", f);
            Assert.DoesNotContain("textdescription", f);
        }

        [Fact]
        public void BodyContains_AddsBodyLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { BodyContains = "draft" });
            Assert.Contains("urn:schemas:httpmail:textdescription LIKE '%draft%'", f);
            Assert.DoesNotContain("subject LIKE", f);
        }

        [Theory]
        [InlineData(true,  "urn:schemas:httpmail:hasattachment = 1")]
        [InlineData(false, "urn:schemas:httpmail:hasattachment = 0")]
        public void HasAttachment_MapsToBool(bool value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { HasAttachment = value });
            Assert.Contains(expected, f);
        }

        [Theory]
        [InlineData(true,  "urn:schemas:httpmail:read = 0")]
        [InlineData(false, "urn:schemas:httpmail:read = 1")]
        public void IsUnread_MapsToInverseRead(bool value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { IsUnread = value });
            Assert.Contains(expected, f);
        }

        [Theory]
        [InlineData(true,  "= 2")]
        [InlineData(false, "<> 2")]
        public void IsFlagged_UsesFlagStatusProperty(bool value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { IsFlagged = value });
            Assert.Contains("0x10900003", f);
            Assert.Contains(expected, f);
        }

        [Theory]
        [InlineData("low",    "= 0")]
        [InlineData("normal", "= 1")]
        [InlineData("high",   "= 2")]
        public void Importance_MapsToMapiPropertyValues(string ui, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Importance = ui });
            Assert.Contains("0x00170003", f);
            Assert.Contains(expected, f);
        }

        [Fact]
        public void Importance_UnknownValue_IsIgnored()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Importance = "Extreme" });
            Assert.Null(f);
        }

        [Fact]
        public void Compound_AllFieldsAndedTogether()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                Query = "Q4",
                From = "jane",
                IsUnread = true,
                HasAttachment = true,
                Importance = "high",
                DateFrom = new DateTimeOffset(2026, 5, 10, 0, 0, 0, TimeSpan.Zero),
            });
            Assert.StartsWith("@SQL=", f);
            // AND-joined; count of ' AND ' separators = clauses - 1.
            var ands = System.Text.RegularExpressions.Regex.Matches(f, " AND ").Count;
            Assert.Equal(5, ands); // query, from, unread, attachment, importance, datefrom = 6 clauses -> 5 ANDs
        }

        [Fact]
        public void EscapesSingleQuotes()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Query = "Jane's Q4" });
            // Single quote inside DASL string is escaped by doubling.
            Assert.Contains("'%Jane''s Q4%'", f);
        }
    }
}
```

- [ ] **Step 4: Run the new tests; expect FAIL**

Run with the filter on the new class name. Expected: 11 failures because the existing `BuildRestrictFilter` only honors `Query`/`DateFrom`/`DateTo`.

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~BuildRestrictFilterTests"
```

- [ ] **Step 5: Replace `BuildRestrictFilter` with the full implementation**

In `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`, replace the entire `BuildRestrictFilter` method with:

```csharp
internal static string BuildRestrictFilter(SearchMessagesArgs args)
{
    if (args == null) return null;
    var clauses = new List<string>();

    if (!string.IsNullOrEmpty(args.Query))
    {
        var q = Escape(args.Query);
        clauses.Add("(urn:schemas:httpmail:subject LIKE '%" + q + "%' OR " +
                    "urn:schemas:httpmail:textdescription LIKE '%" + q + "%')");
    }
    if (!string.IsNullOrEmpty(args.From))
    {
        var v = Escape(args.From);
        clauses.Add("(urn:schemas:httpmail:fromname LIKE '%" + v + "%' OR " +
                    "urn:schemas:httpmail:fromemail LIKE '%" + v + "%')");
    }
    if (!string.IsNullOrEmpty(args.SubjectContains))
    {
        clauses.Add("urn:schemas:httpmail:subject LIKE '%" + Escape(args.SubjectContains) + "%'");
    }
    if (!string.IsNullOrEmpty(args.BodyContains))
    {
        clauses.Add("urn:schemas:httpmail:textdescription LIKE '%" + Escape(args.BodyContains) + "%'");
    }
    if (args.HasAttachment.HasValue)
    {
        clauses.Add("urn:schemas:httpmail:hasattachment = " + (args.HasAttachment.Value ? "1" : "0"));
    }
    if (args.IsUnread.HasValue)
    {
        // urn:schemas:httpmail:read is 0 for unread, 1 for read.
        clauses.Add("urn:schemas:httpmail:read = " + (args.IsUnread.Value ? "0" : "1"));
    }
    if (args.IsFlagged.HasValue)
    {
        // PR_FLAG_STATUS (0x1090) is PT_LONG (0x0003) => 0x10900003. Value 2 = followup flagged.
        clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" " +
                    (args.IsFlagged.Value ? "= 2" : "<> 2"));
    }
    if (!string.IsNullOrEmpty(args.Importance))
    {
        // PR_IMPORTANCE (0x0017) PT_LONG => 0x00170003. 0=low, 1=normal, 2=high.
        var imp = args.Importance.Trim().ToLowerInvariant();
        if (imp == "low" || imp == "normal" || imp == "high")
        {
            var val = imp == "low" ? "0" : imp == "normal" ? "1" : "2";
            clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x00170003\" = " + val);
        }
        // Unknown importance values are silently ignored - tool layer
        // could surface this, but tolerating noise here is safer.
    }
    if (args.DateFrom.HasValue)
    {
        clauses.Add("urn:schemas:httpmail:datereceived >= '" +
            args.DateFrom.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
    }
    if (args.DateTo.HasValue)
    {
        clauses.Add("urn:schemas:httpmail:datereceived <= '" +
            args.DateTo.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
    }

    if (clauses.Count == 0) return null;
    return "@SQL=" + string.Join(" AND ", clauses);
}
```

- [ ] **Step 6: Run all BuildRestrictFilter tests; expect PASS**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~BuildRestrictFilterTests"
```

Expected: 11 / 11 pass.

Note on the MAPI proptag clauses: the test only asserts on the substrings `0x10900003`, `0x00170003`, `= 2`, `<> 2`, `= 0`, `= 1`. The quote marks around the proptag URI are not asserted - leave them in the implementation to satisfy Outlook's DASL parser (it requires quotes around URIs that contain `://`).

- [ ] **Step 7: Run the full suite to confirm no regressions**

Run the full test command. Expected: 120 / 120 (109 baseline + 11 new).

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Properties/AssemblyInfo.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/BuildRestrictFilterTests.cs
git commit -m "Phase 3a Task 2: BuildRestrictFilter handles all structured search fields"
```

---

## Task 3: Update `ToolCatalogSchema` for `outlook_search_messages` + `outlook_count_messages`

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\ToolCatalogSchema.cs`

The schema sent to the model must advertise the new fields. Existing entries describe only `query` / `folder_id` / `date_from` / `date_to` / `max_results`. Replace each with the full 12-field shape.

- [ ] **Step 1: Replace the `outlook_search_messages` entry**

Find `BuildToolEntry("outlook_search_messages", ...)` in `ToolCatalogSchema.cs` and replace the entire entry (BuildToolEntry call) with:

```csharp
BuildToolEntry("outlook_search_messages",
    "Search messages. Combine any subset of filters via AND. Returns id+metadata+snippet for up to max_results (default 25, hard cap 100). Prefer structured filters (from, subject_contains, body_contains, has_attachment, is_unread, is_flagged, importance) over free-form 'query' when you can - one precise call beats five sequential searches.",
    new JObject(
        new JProperty("type", "object"),
        new JProperty("properties", new JObject(
            new JProperty("query",            new JObject(new JProperty("type","string"),
                                              new JProperty("description","Free-form text matched against subject + body. Leave empty if filtering by structured fields only."))),
            new JProperty("from",             new JObject(new JProperty("type","string"),
                                              new JProperty("description","Sender substring; matches display name OR email (case-insensitive)."))),
            new JProperty("subject_contains", new JObject(new JProperty("type","string"))),
            new JProperty("body_contains",    new JObject(new JProperty("type","string"))),
            new JProperty("has_attachment",   new JObject(new JProperty("type","boolean"))),
            new JProperty("is_unread",        new JObject(new JProperty("type","boolean"))),
            new JProperty("is_flagged",       new JObject(new JProperty("type","boolean"))),
            new JProperty("importance",       new JObject(new JProperty("type","string"),
                                              new JProperty("enum", new JArray("low","normal","high")))),
            new JProperty("folder_id",        new JObject(new JProperty("type","string"),
                                              new JProperty("description","Default: Inbox."))),
            new JProperty("date_from",        new JObject(new JProperty("type","string"),
                                                          new JProperty("format","date-time"))),
            new JProperty("date_to",          new JObject(new JProperty("type","string"),
                                                          new JProperty("format","date-time"))),
            new JProperty("max_results",      new JObject(new JProperty("type","integer"),
                                                          new JProperty("minimum",1),
                                                          new JProperty("maximum",100))))),
        new JProperty("additionalProperties", false)))
```

- [ ] **Step 2: Replace the `outlook_count_messages` entry**

Find `BuildToolEntry("outlook_count_messages", ...)` and replace the entire entry with the same property set (same `SearchMessagesArgs`-shape, different description):

```csharp
BuildToolEntry("outlook_count_messages",
    "Count messages matching the given filters without returning bodies. Same filter fields as outlook_search_messages.",
    new JObject(
        new JProperty("type", "object"),
        new JProperty("properties", new JObject(
            new JProperty("query",            new JObject(new JProperty("type","string"))),
            new JProperty("from",             new JObject(new JProperty("type","string"))),
            new JProperty("subject_contains", new JObject(new JProperty("type","string"))),
            new JProperty("body_contains",    new JObject(new JProperty("type","string"))),
            new JProperty("has_attachment",   new JObject(new JProperty("type","boolean"))),
            new JProperty("is_unread",        new JObject(new JProperty("type","boolean"))),
            new JProperty("is_flagged",       new JObject(new JProperty("type","boolean"))),
            new JProperty("importance",       new JObject(new JProperty("type","string"),
                                              new JProperty("enum", new JArray("low","normal","high")))),
            new JProperty("folder_id",        new JObject(new JProperty("type","string"))),
            new JProperty("date_from",        new JObject(new JProperty("type","string"),
                                                          new JProperty("format","date-time"))),
            new JProperty("date_to",          new JObject(new JProperty("type","string"),
                                                          new JProperty("format","date-time"))))),
        new JProperty("additionalProperties", false)))
```

Neither tool declares a `required` array — every field is optional. The model may call them with zero arguments (returns all messages in the default folder, capped by max_results).

- [ ] **Step 3: Build**

Run the build command. Expected: clean.

- [ ] **Step 4: Run the full suite**

Expected: 120 / 120 still pass. Schema changes are invisible to existing tests (no test asserts on the JSON shape directly; they just confirm tools are findable by name).

- [ ] **Step 5: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs
git commit -m "Phase 3a Task 3: advertise structured search fields in tool catalog"
```

---

## Task 4: Extend `OutlookSearchMessagesTool` + `OutlookCountMessagesTool` arg parsing + tests; PUSH CHECKPOINT

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\OutlookSearchMessagesTool.cs`
- Modify: `VSTO2\OutlookAI\Services\Tools\OutlookCountMessagesTool.cs`
- Modify: `VSTO2\OutlookAI.Tests\Services\Tools\OutlookSearchMessagesToolTests.cs`
- Modify: `VSTO2\OutlookAI.Tests\Services\Tools\OutlookCountMessagesToolTests.cs`
- Modify: `VSTO2\OutlookAI.Tests\Services\Tools\MinimalSurface.cs`

Each tool's `Execute` currently parses only `query`/`folder_id`/`date_from`/`date_to`/`max_results` from the input JSON. Extend the parsing to fill all the new `SearchMessagesArgs` fields. The tools themselves are thin — they just hand the args object to the surface — so the change is in arg-extraction.

Tests use `MinimalSurface` (a stub surface). Extend the stub's `SearchMessages`/`CountMessages` to capture and expose the args it received so tests can assert the parsed values.

- [ ] **Step 1: Write the new failing test for `OutlookSearchMessagesTool`**

Open `VSTO2\OutlookAI.Tests\Services\Tools\OutlookSearchMessagesToolTests.cs`. Append a new fact at the end (before the closing brace):

```csharp
[Fact]
public async Task Execute_ParsesAllStructuredFields_AndPassesToSurface()
{
    var captured = (SearchMessagesArgs)null;
    var surface = new MinimalSurface
    {
        SearchMessagesImpl = args => { captured = args; return new MessageSummary[0]; }
    };
    var tool = new OutlookSearchMessagesTool();
    var argsJson = @"{
        ""query"": ""Q4"",
        ""from"": ""jane@acme.com"",
        ""subject_contains"": ""plan"",
        ""body_contains"": ""draft"",
        ""has_attachment"": true,
        ""is_unread"": true,
        ""is_flagged"": false,
        ""importance"": ""high"",
        ""date_from"": ""2026-05-10T00:00:00Z"",
        ""date_to"":   ""2026-05-17T00:00:00Z"",
        ""max_results"": 50
    }";

    await tool.ExecuteAsync(surface, JObject.Parse(argsJson), CancellationToken.None);

    Assert.NotNull(captured);
    Assert.Equal("Q4", captured.Query);
    Assert.Equal("jane@acme.com", captured.From);
    Assert.Equal("plan", captured.SubjectContains);
    Assert.Equal("draft", captured.BodyContains);
    Assert.Equal(true,  captured.HasAttachment);
    Assert.Equal(true,  captured.IsUnread);
    Assert.Equal(false, captured.IsFlagged);
    Assert.Equal("high", captured.Importance);
    Assert.Equal(50, captured.MaxResults);
    Assert.Equal(new DateTimeOffset(2026, 5, 10, 0, 0, 0, TimeSpan.Zero), captured.DateFrom);
    Assert.Equal(new DateTimeOffset(2026, 5, 17, 0, 0, 0, TimeSpan.Zero), captured.DateTo);
}
```

Confirm the file already has `using OutlookAI.Services.Tools;` and `using Newtonsoft.Json.Linq;` at the top; add them if missing.

- [ ] **Step 2: Update `MinimalSurface` to capture args**

Open `VSTO2\OutlookAI.Tests\Services\Tools\MinimalSurface.cs`. Find the existing `SearchMessages(SearchMessagesArgs)` implementation (likely returning a fixed array). Replace the surface's search/count members with a delegate-based capture pattern. The relevant excerpt should look like:

```csharp
// At the top of the class, alongside existing fields:
public Func<SearchMessagesArgs, IReadOnlyList<MessageSummary>> SearchMessagesImpl { get; set; }
    = _ => Array.Empty<MessageSummary>();
public Func<SearchMessagesArgs, int> CountMessagesImpl { get; set; } = _ => 0;

public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args)
    => SearchMessagesImpl(args);
public int CountMessages(SearchMessagesArgs args)
    => CountMessagesImpl(args);
```

If existing tests already pass `SearchMessagesImpl` differently (older shape), reconcile by leaving the old delegate name in place. The exact pattern in `MinimalSurface` was set in Task 11-19 of the Phase 2 plan — if your inspection shows a different name, prefer the existing one and update the new test to match. Either way, the goal is "tests can inject behavior + capture args."

- [ ] **Step 3: Run the new test; expect FAIL**

Filter on the new test name:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~Execute_ParsesAllStructuredFields"
```

Expected: FAIL (parsing in the tool only reads `Query`/`FolderId`/`DateFrom`/`DateTo`/`MaxResults`; the new fields are still null/default).

- [ ] **Step 4: Update `OutlookSearchMessagesTool` arg parsing**

Open `VSTO2\OutlookAI\Services\Tools\OutlookSearchMessagesTool.cs`. Find the block that constructs `SearchMessagesArgs` and replace it with:

```csharp
var sma = new SearchMessagesArgs
{
    Query           = (string)args["query"],
    From            = (string)args["from"],
    SubjectContains = (string)args["subject_contains"],
    BodyContains    = (string)args["body_contains"],
    HasAttachment   = args["has_attachment"] != null ? (bool?)(bool)args["has_attachment"] : null,
    IsUnread        = args["is_unread"]      != null ? (bool?)(bool)args["is_unread"]      : null,
    IsFlagged       = args["is_flagged"]     != null ? (bool?)(bool)args["is_flagged"]     : null,
    Importance      = (string)args["importance"],
    FolderId        = (string)args["folder_id"],
    MaxResults      = args["max_results"] != null
        ? Math.Max(1, Math.Min(100, (int)args["max_results"]))
        : 25,
};
if (args["date_from"] != null && DateTimeOffset.TryParse((string)args["date_from"], out var df))
    sma.DateFrom = df;
if (args["date_to"] != null && DateTimeOffset.TryParse((string)args["date_to"], out var dt))
    sma.DateTo = dt;
```

Make sure `using System;` is present at the top of the file (for `DateTimeOffset.TryParse` and `Math`).

- [ ] **Step 5: Update `OutlookCountMessagesTool` arg parsing**

Same change as Step 4, applied to `VSTO2\OutlookAI\Services\Tools\OutlookCountMessagesTool.cs`. The two tools have identical arg-extraction blocks.

- [ ] **Step 6: Add a parallel test for `OutlookCountMessagesTool`**

Open `VSTO2\OutlookAI.Tests\Services\Tools\OutlookCountMessagesToolTests.cs` and append:

```csharp
[Fact]
public async Task Execute_ParsesAllStructuredFields_AndPassesToSurface()
{
    var captured = (SearchMessagesArgs)null;
    var surface = new MinimalSurface
    {
        CountMessagesImpl = args => { captured = args; return 7; }
    };
    var tool = new OutlookCountMessagesTool();
    var argsJson = @"{
        ""from"": ""jane@acme.com"",
        ""is_unread"": true,
        ""has_attachment"": true,
        ""importance"": ""high""
    }";

    var resultJson = await tool.ExecuteAsync(surface, JObject.Parse(argsJson), CancellationToken.None);

    Assert.NotNull(captured);
    Assert.Equal("jane@acme.com", captured.From);
    Assert.Equal(true, captured.IsUnread);
    Assert.Equal(true, captured.HasAttachment);
    Assert.Equal("high", captured.Importance);
    var result = JObject.Parse(resultJson);
    Assert.Equal(7, (int)result["count"]);
}
```

- [ ] **Step 7: Run the new tests; expect PASS**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookSearchMessagesToolTests|FullyQualifiedName~OutlookCountMessagesToolTests"
```

Expected: every existing test in those two classes still passes, plus the two new ones.

- [ ] **Step 8: Full suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: 122 / 122 (120 baseline + 2 new).

- [ ] **Step 9: Commit + push (search-enhancement checkpoint)**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs
git commit -m "Phase 3a Task 4: search + count tools parse structured fields end-to-end"
git push origin feature/codex-oauth-migration
```

---

## Task 5: `CurrentSelectionResult` POCO + `IOutlookSurface.GetCurrentSelection`

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`

POCO + interface method only. No implementation yet (LiveOutlookSurface in Task 8; MinimalSurface in Task 6).

- [ ] **Step 1: Add the POCO**

In `VSTO2\OutlookAI\Services\Tools\IOutlookSurface.cs`, append a new public sealed class at the bottom of the file (after the existing POCOs, inside the same namespace):

```csharp
public sealed class CurrentSelectionResult
{
    /// <summary>Display name of the currently-active folder (e.g. "Inbox").</summary>
    public string Folder { get; set; }
    /// <summary>Stable short-id for the folder; matches outlook_list_folders ids.</summary>
    public string FolderId { get; set; }
    /// <summary>Total selection count (may exceed Messages.Count if MaxItems clamped).</summary>
    public int Count { get; set; }
    /// <summary>Up to <c>MaxItems</c> selected messages, freshest first.</summary>
    public IReadOnlyList<MessageDetail> Messages { get; set; }
}
```

`MessageDetail` (not `MessageSummary`) so the same shape supports both snippet-only and full-body modes — `BodyPlaintext` and `BodyTruncated` are already optional/nullable on that type.

- [ ] **Step 2: Add the interface method**

In the same file, find:

```csharp
public interface IOutlookSurface
{
    ...
    void SetCategory(string messageId, string category);
}
```

Add one new method before the closing brace:

```csharp
/// <summary>
/// Returns the messages currently selected in the active Explorer.
/// When the surface was constructed without an Explorer reference
/// (e.g. for a compose-only Inspector pane), returns an empty result.
/// </summary>
CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems);
```

- [ ] **Step 3: Build (expect FAIL — `LiveOutlookSurface` doesn't implement the method yet)**

Run the build. Expected error:

```
'LiveOutlookSurface' does not implement interface member 'IOutlookSurface.GetCurrentSelection(bool, int)'
```

`MinimalSurface` will also fail to compile until Task 6.

This is expected — we're TDDing the type signature. Tasks 6 + 8 land the actual implementations.

- [ ] **Step 4: Add a no-op default to `LiveOutlookSurface` so the build passes**

Open `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`. Append before the closing brace of the class (after `SetCategory` impl):

```csharp
// Real implementation lands in Task 8 once the Explorer ctor parameter
// is wired. For now return an empty result so this file compiles.
public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) =>
    new CurrentSelectionResult
    {
        Folder = "",
        FolderId = "",
        Count = 0,
        Messages = new MessageDetail[0],
    };
```

- [ ] **Step 5: Build expected PASS, `MinimalSurface` still fails**

Run the build. `OutlookAI.dll` should compile. `OutlookAI.Tests.dll` will fail with:

```
'MinimalSurface' does not implement interface member 'IOutlookSurface.GetCurrentSelection(bool, int)'
```

That's the cliffhanger Task 6 resolves.

- [ ] **Step 6: Commit the POCO + interface change (the test-helper completion is Task 6)**

```powershell
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "Phase 3a Task 5: IOutlookSurface.GetCurrentSelection + CurrentSelectionResult POCO"
```

---

## Task 6: `MinimalSurface.GetCurrentSelection` stub + delegate hook for tests

**Files:**
- Modify: `VSTO2\OutlookAI.Tests\Services\Tools\MinimalSurface.cs`

Same delegate-injection pattern as the search/count methods.

- [ ] **Step 1: Add the delegate + implementation**

In `MinimalSurface.cs`, add a new public field alongside the existing `SearchMessagesImpl` / `CountMessagesImpl`:

```csharp
public Func<bool, int, CurrentSelectionResult> GetCurrentSelectionImpl { get; set; }
    = (_, __) => new CurrentSelectionResult
    {
        Folder = "Inbox",
        FolderId = "fld_root",
        Count = 0,
        Messages = new MessageDetail[0],
    };

public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
    => GetCurrentSelectionImpl(includeFullBodies, maxItems);
```

- [ ] **Step 2: Build**

Expected: clean.

- [ ] **Step 3: Run full suite**

Expected: 122 / 122 (no behavior change; just type completion).

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs
git commit -m "Phase 3a Task 6: MinimalSurface.GetCurrentSelection stub + capture hook"
```

---

## Task 7: `OutlookGetCurrentSelectionTool` + tests

**Files:**
- Create: `VSTO2\OutlookAI\Services\Tools\OutlookGetCurrentSelectionTool.cs`
- Create: `VSTO2\OutlookAI.Tests\Services\Tools\OutlookGetCurrentSelectionToolTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj` (Compile entry)

Tool reads the args, calls `surface.GetCurrentSelection`, projects to the wire JSON shape.

- [ ] **Step 1: Write the failing tests**

Create `VSTO2\OutlookAI.Tests\Services\Tools\OutlookGetCurrentSelectionToolTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookGetCurrentSelectionToolTests
    {
        [Fact]
        public async Task Execute_DefaultArgs_ReturnsFolderCountAndEmptyList()
        {
            var surface = new MinimalSurface();   // default impl returns empty
            var tool = new OutlookGetCurrentSelectionTool();

            var resultJson = await tool.ExecuteAsync(surface, new JObject(), CancellationToken.None);

            var result = JObject.Parse(resultJson);
            Assert.Equal("Inbox", (string)result["folder"]);
            Assert.Equal("fld_root", (string)result["folder_id"]);
            Assert.Equal(0, (int)result["count"]);
            Assert.NotNull(result["messages"]);
            Assert.Empty((JArray)result["messages"]);
        }

        [Fact]
        public async Task Execute_RespectsMaxItems_PassedToSurface()
        {
            bool capturedIncludeBodies = false;
            int capturedMaxItems = -1;
            var surface = new MinimalSurface
            {
                GetCurrentSelectionImpl = (incl, max) =>
                {
                    capturedIncludeBodies = incl;
                    capturedMaxItems = max;
                    return new CurrentSelectionResult
                    {
                        Folder = "Inbox",
                        FolderId = "fld_root",
                        Count = 0,
                        Messages = new MessageDetail[0],
                    };
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();
            var args = new JObject(new JProperty("max_items", 3), new JProperty("include_full_bodies", true));

            await tool.ExecuteAsync(surface, args, CancellationToken.None);

            Assert.Equal(3, capturedMaxItems);
            Assert.True(capturedIncludeBodies);
        }

        [Fact]
        public async Task Execute_ClampsMaxItemsToHardCap()
        {
            int capturedMaxItems = -1;
            var surface = new MinimalSurface
            {
                GetCurrentSelectionImpl = (incl, max) =>
                {
                    capturedMaxItems = max;
                    return new CurrentSelectionResult { Messages = new MessageDetail[0] };
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();
            var args = new JObject(new JProperty("max_items", 999));

            await tool.ExecuteAsync(surface, args, CancellationToken.None);

            Assert.Equal(20, capturedMaxItems);   // hard cap per the spec
        }

        [Fact]
        public async Task Execute_ProjectsMessageDetail_ToWireShape()
        {
            var msg = new MessageDetail
            {
                Id = "msg_abc",
                Subject = "Re: Q4 plan",
                From = "Jane Doe <jane@acme.com>",
                ReceivedAt = new DateTimeOffset(2026, 5, 17, 9, 14, 0, TimeSpan.Zero),
                BodyPlaintext = "Hi team --- thoughts on regional split...",
                BodyTruncated = false,
                Attachments = new[] { new AttachmentSummary { Filename = "plan.xlsx", SizeBytes = 4096 } },
                ConversationTopic = "Q4 plan",
            };
            var surface = new MinimalSurface
            {
                GetCurrentSelectionImpl = (incl, max) => new CurrentSelectionResult
                {
                    Folder = "Inbox",
                    FolderId = "fld_inbox",
                    Count = 1,
                    Messages = new[] { msg },
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            var resultJson = await tool.ExecuteAsync(surface, new JObject(), CancellationToken.None);

            var result = JObject.Parse(resultJson);
            Assert.Equal("Inbox", (string)result["folder"]);
            Assert.Equal(1, (int)result["count"]);
            var arr = (JArray)result["messages"];
            Assert.Single(arr);
            var m = (JObject)arr[0];
            Assert.Equal("msg_abc", (string)m["id"]);
            Assert.Equal("Re: Q4 plan", (string)m["subject"]);
            Assert.Equal("Jane Doe <jane@acme.com>", (string)m["from"]);
            Assert.Equal("Q4 plan", (string)m["conversation_topic"]);
            Assert.True((bool)m["has_attachments"]);
        }
    }
}
```

- [ ] **Step 2: Run; expect FAIL with "tool not found"**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookGetCurrentSelectionToolTests"
```

Expected: build fails because `OutlookGetCurrentSelectionTool` doesn't exist. Good — TDD red.

- [ ] **Step 3: Implement the tool**

Create `VSTO2\OutlookAI\Services\Tools\OutlookGetCurrentSelectionTool.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Returns the messages currently selected in the user's active
    /// Explorer (Inbox view). Read-only; the model uses this when the
    /// user says things like "reply to this" or "summarize the selected".
    /// </summary>
    public sealed class OutlookGetCurrentSelectionTool : IOutlookTool
    {
        public string Name => "outlook_get_current_selection";

        public Task<string> ExecuteAsync(
            IOutlookSurface surface,
            JObject args,
            CancellationToken cancellationToken)
        {
            args = args ?? new JObject();
            var includeBodies = (bool?)args["include_full_bodies"] ?? false;
            var maxItemsRaw = (int?)args["max_items"] ?? 5;
            var maxItems = Math.Max(1, Math.Min(20, maxItemsRaw));

            var result = surface.GetCurrentSelection(includeBodies, maxItems);
            var json = ProjectToJson(result, includeBodies);
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static JObject ProjectToJson(CurrentSelectionResult r, bool includeBodies)
        {
            var arr = new JArray();
            if (r?.Messages != null)
            {
                foreach (var m in r.Messages)
                {
                    var item = new JObject(
                        new JProperty("id", m.Id ?? ""),
                        new JProperty("subject", m.Subject ?? ""),
                        new JProperty("from", m.From ?? ""),
                        new JProperty("received_at", m.ReceivedAt.ToString("o")),
                        new JProperty("conversation_topic", m.ConversationTopic ?? ""),
                        new JProperty("has_attachments", m.Attachments != null && m.Attachments.Count > 0));
                    if (includeBodies)
                    {
                        item.Add("body_plaintext", m.BodyPlaintext ?? "");
                        item.Add("body_truncated", m.BodyTruncated);
                    }
                    else
                    {
                        // Snippet = first ~200 chars of body for context without
                        // shipping the whole thing every turn.
                        var snippet = m.BodyPlaintext ?? "";
                        if (snippet.Length > 200) snippet = snippet.Substring(0, 200);
                        item.Add("snippet", snippet);
                    }
                    arr.Add(item);
                }
            }
            return new JObject(
                new JProperty("folder", r?.Folder ?? ""),
                new JProperty("folder_id", r?.FolderId ?? ""),
                new JProperty("count", r?.Count ?? 0),
                new JProperty("messages", arr));
        }
    }
}
```

- [ ] **Step 4: Register in csproj**

Open `VSTO2\OutlookAI\OutlookAI.csproj`. Find the existing tool registrations (look for `OutlookSetCategoryTool.cs`). Insert one new Compile entry next to it:

```xml
<Compile Include="Services\Tools\OutlookGetCurrentSelectionTool.cs" />
```

- [ ] **Step 5: Build**

Expected: clean.

- [ ] **Step 6: Run the new tests; expect PASS**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookGetCurrentSelectionToolTests"
```

Expected: 4 / 4.

- [ ] **Step 7: Full suite**

Expected: 126 / 126 (122 baseline + 4 new).

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookGetCurrentSelectionTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookGetCurrentSelectionToolTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 3a Task 7: outlook_get_current_selection tool + 4 unit tests"
```

---

## Task 8: `ToolCatalogSchema` advertises the selection tool

**Files:**
- Modify: `VSTO2\OutlookAI\Services\Tools\ToolCatalogSchema.cs`

- [ ] **Step 1: Add the schema entry**

In `BuildResponsesToolsArray`, find the existing read-tool block (somewhere near `outlook_list_recent_threads_with`). Insert one new entry into the `arr` initializer or adjacent `.Add` call — placement near the other read tools is fine. The full call:

```csharp
arr.Add(BuildToolEntry("outlook_get_current_selection",
    "Read the messages currently selected in the user's active Explorer (e.g. messages they highlighted in the reading pane). Useful for 'reply to this', 'summarize this thread', etc. Returns empty when nothing is selected or when there is no active Explorer (e.g. the chat is anchored to a compose window).",
    new JObject(
        new JProperty("type", "object"),
        new JProperty("properties", new JObject(
            new JProperty("include_full_bodies", new JObject(
                new JProperty("type","boolean"),
                new JProperty("description","If true, returns full message body per item (up to 32 KB). Default false: 200-char snippet only."))),
            new JProperty("max_items", new JObject(
                new JProperty("type","integer"),
                new JProperty("minimum",1),
                new JProperty("maximum",20),
                new JProperty("description","Hard cap on items returned. Default 5."))))),
        new JProperty("additionalProperties", false))));
```

- [ ] **Step 2: Build + run full suite**

Expected: 126 / 126.

- [ ] **Step 3: Commit**

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs
git commit -m "Phase 3a Task 8: advertise outlook_get_current_selection in tool catalog"
```

---

## Task 9: `OutlookToolHost` registers the selection tool + `LiveOutlookSurface` Explorer plumbing; PUSH CHECKPOINT

**Files:**
- Modify: `VSTO2\OutlookAI\Services\OutlookToolHost.cs`
- Modify: `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`

Two changes:
1. `OutlookToolHost` adds `new OutlookGetCurrentSelectionTool()` to its registered set (always, like other read tools).
2. `LiveOutlookSurface` gets a new `_explorer` field, a new constructor that accepts an `Outlook.Explorer`, and a real `GetCurrentSelection` implementation replacing the Task-5 stub.

The Phase 2 Inspector-only constructor stays so existing callers (`AITaskPane.Bind` in compose mode) continue to compile.

- [ ] **Step 1: Register the new tool in `OutlookToolHost`**

Open `VSTO2\OutlookAI\Services\OutlookToolHost.cs`. Find the `tools` list initializer in the constructor. Add one entry alongside the other always-on read tools (e.g. after `OutlookListRecentThreadsWithTool`):

```csharp
new OutlookGetCurrentSelectionTool(),
```

- [ ] **Step 2: Extend `LiveOutlookSurface` with `_explorer` field + constructor overload**

In `VSTO2\OutlookAI\Services\Tools\LiveOutlookSurface.cs`:

a. Add a private field next to the existing `_composeInspector`:

```csharp
private readonly Outlook.Explorer _explorer;
```

b. Replace the existing single constructor with two:

```csharp
public LiveOutlookSurface(
    Outlook.Application application,
    OutlookThreadMarshaller marshaller,
    IdResolver ids,
    Outlook.Inspector composeInspector)
    : this(application, marshaller, ids, composeInspector, explorer: null)
{
}

public LiveOutlookSurface(
    Outlook.Application application,
    OutlookThreadMarshaller marshaller,
    IdResolver ids,
    Outlook.Inspector composeInspector,
    Outlook.Explorer explorer)
{
    _application = application ?? throw new ArgumentNullException(nameof(application));
    _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
    _ids = ids ?? throw new ArgumentNullException(nameof(ids));
    _composeInspector = composeInspector;
    _explorer = explorer;
}
```

(The single-arg form preserves binary compatibility with `AITaskPane.Bind`.)

- [ ] **Step 3: Replace the Task-5 stub `GetCurrentSelection` with the real implementation**

Find the stub `public CurrentSelectionResult GetCurrentSelection(...)` we wrote in Task 5 Step 4 and replace it with:

```csharp
public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) =>
    Run(() =>
    {
        if (_explorer == null)
        {
            return new CurrentSelectionResult
            {
                Folder = "",
                FolderId = "",
                Count = 0,
                Messages = new MessageDetail[0],
            };
        }
        var folder = _explorer.CurrentFolder;
        var folderName = folder?.Name ?? "";
        var folderId = folder != null ? _ids.Shorten(folder.EntryID ?? "") : "";

        var selection = _explorer.Selection;
        int totalCount = 0;
        try { totalCount = selection.Count; } catch (COMException) { }

        var picked = new List<MessageDetail>();
        try
        {
            int taken = 0;
            // Outlook Selection is 1-based. Iterate up to min(Selection.Count, maxItems).
            for (int i = 1; i <= totalCount && taken < maxItems; i++)
            {
                object item = null;
                try { item = selection[i]; } catch (COMException) { continue; }
                if (!(item is Outlook.MailItem mi)) continue;

                var body = mi.Body ?? "";
                bool truncated = false;
                if (!includeFullBodies && body.Length > 1000)
                {
                    body = body.Substring(0, 1000);
                    truncated = true;
                }
                else if (body.Length > MaxBodyChars)
                {
                    body = body.Substring(0, MaxBodyChars);
                    truncated = true;
                }

                var atts = new List<AttachmentSummary>();
                try
                {
                    foreach (Outlook.Attachment att in mi.Attachments)
                    {
                        atts.Add(new AttachmentSummary
                        {
                            Filename = att.FileName,
                            SizeBytes = att.Size,
                        });
                    }
                }
                catch (COMException) { }

                picked.Add(new MessageDetail
                {
                    Id = _ids.Shorten(mi.EntryID ?? ""),
                    Subject = mi.Subject ?? "",
                    From = (mi.SenderName ?? "") +
                           (string.IsNullOrEmpty(mi.SenderEmailAddress) ? "" :
                            " <" + mi.SenderEmailAddress + ">"),
                    To = SplitAddresses(mi.To),
                    Cc = SplitAddresses(mi.CC),
                    ReceivedAt = ToOffset(mi.ReceivedTime),
                    BodyPlaintext = body,
                    BodyTruncated = truncated,
                    Attachments = atts,
                    InReplyToMessageId = null,
                    ConversationTopic = mi.ConversationTopic ?? "",
                });
                taken++;
            }
        }
        catch (COMException) { }

        return new CurrentSelectionResult
        {
            Folder = folderName,
            FolderId = folderId,
            Count = totalCount,
            Messages = picked,
        };
    });
```

- [ ] **Step 4: Build**

Expected: clean.

- [ ] **Step 5: Full suite**

Expected: 126 / 126 (no behavior change for existing callers; the new constructor isn't exercised yet — that happens in Task 12).

- [ ] **Step 6: Commit + push (selection-tool checkpoint)**

```powershell
git add VSTO2/OutlookAI/Services/OutlookToolHost.cs VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "Phase 3a Task 9: register selection tool + LiveOutlookSurface Explorer plumbing"
git push origin feature/codex-oauth-migration
```

---

## Task 10: `QuickActionChip` POCO + `ComputeChipsForSelectionCount` helper + tests

**Files:**
- Create: `VSTO2\OutlookAI\TaskPane\InboxCopilot\QuickActionChip.cs`
- Create: `VSTO2\OutlookAI.Tests\TaskPane\InboxCopilot\QuickActionChipTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

Pure C#, no Outlook deps, fully unit-testable. The chip set is computed server-side and pushed to the WebUI per § 6 of the spec.

- [ ] **Step 1: Create folders**

```powershell
New-Item -ItemType Directory -Path "VSTO2\OutlookAI\TaskPane\InboxCopilot" -Force | Out-Null
New-Item -ItemType Directory -Path "VSTO2\OutlookAI.Tests\TaskPane\InboxCopilot" -Force | Out-Null
```

- [ ] **Step 2: Write the failing tests**

Create `VSTO2\OutlookAI.Tests\TaskPane\InboxCopilot\QuickActionChipTests.cs`:

```csharp
using System.Linq;
using OutlookAI.TaskPane.InboxCopilot;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxCopilot
{
    public class QuickActionChipTests
    {
        [Fact]
        public void NoSelection_ReturnsThreeStaticChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(0);
            Assert.Equal(3, chips.Count);
            Assert.Contains(chips, c => c.Label == "What needs my attention?");
            Assert.Contains(chips, c => c.Label == "Summarize unread");
            Assert.Contains(chips, c => c.Label == "Today's emails");
        }

        [Fact]
        public void SingleSelection_AddsSingleSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(1);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Summarize this thread");
            Assert.Contains(chips, c => c.Label == "Draft a reply");
            // Static three still present:
            Assert.Contains(chips, c => c.Label == "Today's emails");
        }

        [Fact]
        public void MultiSelection_AddsMultiSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(3);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Summarize all selected");
            Assert.Contains(chips, c => c.Label == "Triage selected");
            // Static three still present:
            Assert.Contains(chips, c => c.Label == "What needs my attention?");
            // Single-selection chips NOT present:
            Assert.DoesNotContain(chips, c => c.Label == "Summarize this thread");
            Assert.DoesNotContain(chips, c => c.Label == "Draft a reply");
        }

        [Fact]
        public void Prompts_AreNotEmpty()
        {
            // Every chip must have a non-empty prompt the model will receive.
            foreach (var n in new[] { 0, 1, 5 })
            {
                var chips = QuickActionChip.ComputeChipsForSelectionCount(n);
                Assert.All(chips, c => Assert.False(string.IsNullOrWhiteSpace(c.Prompt)));
            }
        }
    }
}
```

- [ ] **Step 3: Run; expect FAIL (type doesn't exist)**

Build should fail. Good.

- [ ] **Step 4: Implement `QuickActionChip`**

Create `VSTO2\OutlookAI\TaskPane\InboxCopilot\QuickActionChip.cs`:

```csharp
using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// A pre-canned prompt shown as a clickable chip above the Inbox
    /// Copilot composer. Clicking a chip pre-fills the textarea and
    /// auto-sends (per spec). The set is computed server-side based on
    /// how many messages the user has selected in the active Explorer.
    /// </summary>
    public sealed class QuickActionChip
    {
        public string Label { get; set; }    // shown on the button
        public string Prompt { get; set; }   // pre-filled into the textarea

        /// <summary>
        /// Build the default chip set for a given selection count.
        /// Three static chips plus 0 or 2 dynamic chips:
        ///   0 selected -> static only
        ///   1 selected -> static + "Summarize this thread" + "Draft a reply"
        ///   2+ selected -> static + "Summarize all selected" + "Triage selected"
        /// </summary>
        public static IReadOnlyList<QuickActionChip> ComputeChipsForSelectionCount(int selectionCount)
        {
            var list = new List<QuickActionChip>
            {
                new QuickActionChip
                {
                    Label = "What needs my attention?",
                    Prompt = "Look at my inbox and tell me what needs attention. Prioritize by recency, importance, and sender. Be concise.",
                },
                new QuickActionChip
                {
                    Label = "Summarize unread",
                    Prompt = "Summarize all my unread messages. Group by sender or topic. Be concise.",
                },
                new QuickActionChip
                {
                    Label = "Today's emails",
                    Prompt = "Show me everything I received today, grouped by sender. Highlight anything that looks urgent.",
                },
            };

            if (selectionCount == 1)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Summarize this thread",
                    Prompt = "Summarize the selected message and the rest of its conversation thread.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Draft a reply",
                    Prompt = "Draft a reply to the selected message. Match the tone of the sender.",
                });
            }
            else if (selectionCount >= 2)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Summarize all selected",
                    Prompt = "Summarize all the selected messages.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Triage selected",
                    Prompt = "Triage the selected messages -- which need action, which can be archived, which can be marked read?",
                });
            }

            return list;
        }
    }
}
```

- [ ] **Step 5: Add Compile entry to csproj**

Open `VSTO2\OutlookAI\OutlookAI.csproj`. Find the existing TaskPane entries (look for `TaskPane\Variants\VariantsController.cs`). Insert one new Compile entry below it:

```xml
<Compile Include="TaskPane\InboxCopilot\QuickActionChip.cs" />
```

- [ ] **Step 6: Build + run tests; expect PASS**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~QuickActionChipTests"
```

Expected: 4 / 4 pass.

- [ ] **Step 7: Full suite**

Expected: 130 / 130.

- [ ] **Step 8: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxCopilot/QuickActionChip.cs VSTO2/OutlookAI.Tests/TaskPane/InboxCopilot/QuickActionChipTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 3a Task 10: QuickActionChip POCO + chip-set computation + 4 tests"
```

---

## Task 11: `InboxCopilotPromptBuilder` + tests

**Files:**
- Create: `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPromptBuilder.cs`
- Create: `VSTO2\OutlookAI.Tests\TaskPane\InboxCopilot\InboxCopilotPromptBuilderTests.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

Pure-function helper that builds the `SystemInstructions` string for a given folder + selection state. Tested standalone before the controller wires it up.

- [ ] **Step 1: Write the failing tests**

Create `VSTO2\OutlookAI.Tests\TaskPane\InboxCopilot\InboxCopilotPromptBuilderTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.InboxCopilot;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxCopilot
{
    public class InboxCopilotPromptBuilderTests
    {
        [Fact]
        public void NoSelection_PromptIncludesFolderAndUnreadCountOnly()
        {
            var prompt = InboxCopilotPromptBuilder.Build(
                folderName: "Inbox",
                unreadCount: 47,
                totalCount: 1284,
                selection: null);

            Assert.Contains("Inbox", prompt);
            Assert.Contains("47 unread", prompt);
            Assert.Contains("1284 total", prompt);
            Assert.DoesNotContain("Selected:", prompt);
        }

        [Fact]
        public void SingleSelection_PromptIncludesSelectedMessageBlock()
        {
            var sel = new CurrentSelectionResult
            {
                Folder = "Inbox",
                FolderId = "fld",
                Count = 1,
                Messages = new[]
                {
                    new MessageDetail
                    {
                        Id = "m1",
                        Subject = "Re: Q4 plan",
                        From = "Jane Doe <jane@acme.com>",
                        ReceivedAt = new DateTimeOffset(2026, 5, 17, 9, 14, 0, TimeSpan.Zero),
                        BodyPlaintext = "Hi team - thoughts on regional split. ",
                    }
                }
            };
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 47, 1284, sel);

            Assert.Contains("Selected:", prompt);
            Assert.Contains("Re: Q4 plan", prompt);
            Assert.Contains("Jane Doe", prompt);
        }

        [Fact]
        public void MultiSelection_PromptSummarizesCount()
        {
            var sel = new CurrentSelectionResult
            {
                Folder = "Inbox",
                FolderId = "fld",
                Count = 4,
                Messages = new MessageDetail[0],   // count without listing every detail
            };
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 47, 1284, sel);

            Assert.Contains("4 messages selected", prompt);
            Assert.DoesNotContain("Re: Q4 plan", prompt);   // no individual subject listed
        }

        [Fact]
        public void Prompt_AlwaysIncludesRolePreamble()
        {
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 0, 0, null);
            Assert.Contains("Inbox Copilot", prompt);
        }
    }
}
```

- [ ] **Step 2: Implement `InboxCopilotPromptBuilder`**

Create `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPromptBuilder.cs`:

```csharp
using System.Text;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Builds the per-turn system instructions for the Inbox Copilot.
    /// Pure function over current folder state + selection so the chat
    /// service always sends fresh context.
    /// </summary>
    public static class InboxCopilotPromptBuilder
    {
        public static string Build(
            string folderName,
            int unreadCount,
            int totalCount,
            CurrentSelectionResult selection)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are the Outlook Inbox Copilot. The user is viewing their mailbox.");
            sb.AppendLine("Help them search, summarize, triage, and act on messages. You have");
            sb.AppendLine("mailbox tools available; prefer one well-targeted tool call over many.");
            sb.AppendLine();
            sb.AppendLine("Current context:");
            sb.Append("- Folder: ").Append(folderName ?? "Inbox");
            sb.Append(" (").Append(unreadCount).Append(" unread, ").Append(totalCount).Append(" total)");
            sb.AppendLine();

            if (selection != null && selection.Count > 0)
            {
                if (selection.Count == 1 && selection.Messages != null && selection.Messages.Count > 0)
                {
                    var m = selection.Messages[0];
                    sb.Append("- Selected: ").AppendLine(m.Subject ?? "");
                    if (!string.IsNullOrEmpty(m.From))
                        sb.Append("  From: ").AppendLine(m.From);
                    sb.Append("  Received: ").AppendLine(m.ReceivedAt.ToString("o"));
                    var snippet = (m.BodyPlaintext ?? "").Replace("\r", " ").Replace("\n", " ");
                    if (snippet.Length > 200) snippet = snippet.Substring(0, 200);
                    if (!string.IsNullOrEmpty(snippet))
                    {
                        sb.Append("  Snippet: ").AppendLine(snippet);
                    }
                }
                else
                {
                    sb.Append("- ").Append(selection.Count).AppendLine(" messages selected");
                }
            }

            sb.AppendLine();
            sb.AppendLine("Reply concisely; the user is busy.");
            return sb.ToString();
        }
    }
}
```

- [ ] **Step 3: Add Compile entry**

In `VSTO2\OutlookAI\OutlookAI.csproj`, next to `QuickActionChip.cs`:

```xml
<Compile Include="TaskPane\InboxCopilot\InboxCopilotPromptBuilder.cs" />
```

- [ ] **Step 4: Build + run new tests**

Expected: 4 / 4.

- [ ] **Step 5: Full suite**

Expected: 134 / 134.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPromptBuilder.cs VSTO2/OutlookAI.Tests/TaskPane/InboxCopilot/InboxCopilotPromptBuilderTests.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 3a Task 11: InboxCopilotPromptBuilder + 4 tests"
```

---

## Task 12: `InboxCopilotPane` UserControl + `.Designer.cs`

**Files:**
- Create: `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPane.cs`
- Create: `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPane.Designer.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

Minimal WinForms UserControl. Holds a single `Panel` that the controller fills with a WebView2. Parallels Phase 2's `AITaskPane` but without the TabControl.

- [ ] **Step 1: Create the Designer file**

Create `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPane.Designer.cs`:

```csharp
namespace OutlookAI.TaskPane.InboxCopilot
{
    partial class InboxCopilotPane
    {
        private System.ComponentModel.IContainer components = null;

        partial void DisposeCustomResources();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                DisposeCustomResources();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.chatHost = new System.Windows.Forms.Panel();
            this.SuspendLayout();

            // chatHost - controller will Dock-Fill a WebView2 into this.
            this.chatHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chatHost.Location = new System.Drawing.Point(0, 0);
            this.chatHost.Size = new System.Drawing.Size(340, 600);
            this.chatHost.TabIndex = 0;

            // InboxCopilotPane
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = false;
            this.BackColor = System.Drawing.Color.FromArgb(250, 249, 248);
            this.Controls.Add(this.chatHost);
            this.Name = "InboxCopilotPane";
            this.Size = new System.Drawing.Size(340, 600);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel chatHost;
    }
}
```

- [ ] **Step 2: Create the code-behind**

Create `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotPane.cs`:

```csharp
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Per-Explorer task pane that hosts the Inbox Copilot chat surface.
    /// Construction is cheap; <see cref="Bind"/> wires up the per-Explorer
    /// LiveOutlookSurface, OutlookToolHost, ConversationStore, and
    /// InboxCopilotController.
    /// </summary>
    public partial class InboxCopilotPane : UserControl
    {
        private Outlook.Explorer _explorer;
        private LiveOutlookSurface _surface;
        private OutlookToolHost _toolHost;
        private ConversationStore _conversationStore;
        private InboxCopilotController _controller;

        public InboxCopilotPane()
        {
            using (TraceLog.Scope("ctor", "InboxCopilotPane"))
            {
                InitializeComponent();
                this.HandleCreated += (s, e) => TraceLog.Write("HandleCreated", "InboxCopilotPane");
                this.VisibleChanged += (s, e) => TraceLog.Write("VisibleChanged Visible=" + this.Visible, "InboxCopilotPane");
            }
        }

        private CodexChatService ChatService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.ChatService : null;

        /// <summary>
        /// Bind this pane to its owning Outlook Explorer. Called by
        /// <see cref="ThisAddIn.ShowExplorerTaskPane"/> immediately after
        /// construction. Builds the per-Explorer surface + tool host + chat
        /// controller and fires the controller's async init.
        /// </summary>
        public void Bind(Outlook.Explorer explorer)
        {
            using (TraceLog.Scope("Bind", "InboxCopilotPane"))
            {
                _explorer = explorer;
                try
                {
                    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                    var ids = Globals.ThisAddIn?.IdResolver;
                    var app = Globals.ThisAddIn?.Application;
                    TraceLog.Write("Services: marshaller=" + (marshaller != null) +
                        " ids=" + (ids != null) + " app=" + (app != null), "InboxCopilotPane");
                    if (marshaller != null && ids != null && app != null)
                    {
                        _surface = new LiveOutlookSurface(app, marshaller, ids,
                            composeInspector: null, explorer: explorer);
                        _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
                        TraceLog.Write("surface + toolHost constructed", "InboxCopilotPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("surface/toolHost error: " + ex, "InboxCopilotPane");
                }

                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        _conversationStore = new ConversationStore();
                        _controller = new InboxCopilotController(
                            chatHost, ChatService, _toolHost, _surface, _conversationStore, explorer);
                        TraceLog.Write("Controller constructed; firing InitializeAsync", "InboxCopilotPane");
                        var initTask = _controller.InitializeAsync();
                        initTask.ContinueWith(t =>
                        {
                            if (t.IsFaulted)
                                TraceLog.Write("InitializeAsync FAULTED: " + t.Exception, "InboxCopilotPane");
                            else if (t.IsCanceled)
                                TraceLog.Write("InitializeAsync CANCELLED", "InboxCopilotPane");
                            else
                                TraceLog.Write("InitializeAsync completed", "InboxCopilotPane");
                        }, TaskScheduler.Default);
                    }
                    else
                    {
                        TraceLog.Write("Controller NOT created (ChatService/toolHost/surface null)", "InboxCopilotPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("Controller construction error: " + ex, "InboxCopilotPane");
                }
            }
        }

        partial void DisposeCustomResources()
        {
            try { _controller?.Dispose(); } catch { }
        }
    }
}
```

- [ ] **Step 3: Add Compile entries (UserControl + Designer)**

In `VSTO2\OutlookAI\OutlookAI.csproj`, after the existing `InboxCopilot\*` entries:

```xml
<Compile Include="TaskPane\InboxCopilot\InboxCopilotPane.cs">
  <SubType>UserControl</SubType>
</Compile>
<Compile Include="TaskPane\InboxCopilot\InboxCopilotPane.Designer.cs">
  <DependentUpon>InboxCopilotPane.cs</DependentUpon>
</Compile>
```

- [ ] **Step 4: Build (expect FAIL — `InboxCopilotController` doesn't exist)**

Build will fail with `'InboxCopilotController' is not defined`. Task 13 lands the controller.

This intentional cliffhanger keeps the pane → controller dependency obvious. Don't try to fix the build before Task 13.

- [ ] **Step 5: Commit (build broken but Step 4 documents why; Task 13 unblocks)**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.cs VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPane.Designer.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 3a Task 12: InboxCopilotPane UserControl + Designer (controller follows in Task 13)"
```

---

## Task 13: `InboxCopilotController` + push checkpoint

**Files:**
- Create: `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotController.cs`
- Modify: `VSTO2\OutlookAI\OutlookAI.csproj`

Parallel to Phase 2's `ChatController`. Reads §§ 3, 5, 6 of the spec.

- [ ] **Step 1: Implement the controller**

Create `VSTO2\OutlookAI\TaskPane\InboxCopilot\InboxCopilotController.cs`:

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Drives the per-Explorer Inbox Copilot chat surface. Mirrors
    /// Phase 2's ChatController but is anchored to an Explorer instead
    /// of an Inspector. Builds a fresh system prompt + quick-action chip
    /// set on every selection change and on every turn.
    /// </summary>
    public sealed class InboxCopilotController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CodexChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Outlook.Explorer _explorer;

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public InboxCopilotController(
            Control hostContainer,
            CodexChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store,
            Outlook.Explorer explorer)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            _explorer = explorer;
        }

        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "InboxCopilot");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                ShowFallback("WebView2 runtime not installed.\r\nRun the installer or download:\r\n" +
                             "https://developer.microsoft.com/microsoft-edge/webview2/");
                return;
            }
            _webView = new WebView2 { Dock = DockStyle.Fill };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");

                // Subscribe to Explorer events for live context refresh.
                if (_explorer != null)
                {
                    _explorer.SelectionChange += OnExplorerSelectionChange;
                    _explorer.FolderSwitch += OnExplorerFolderSwitch;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "InboxCopilot");
                ShowFallback("WebView2 failed to initialize: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null) { _fallbackLabel.Text = message; return; }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message,
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnExplorerSelectionChange()
        {
            if (_isDisposed || !_isReady) return;
            PushContextStripAndChips();
        }

        private void OnExplorerFolderSwitch()
        {
            if (_isDisposed || !_isReady) return;
            PushContextStripAndChips();
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "InboxCopilot");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                HandleHostMessage(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "InboxCopilot");
            }
        }

        private void HandleHostMessage(string type, JObject payload)
        {
            switch (type)
            {
                case "ready":
                    OnWebViewReady();
                    break;
                case "send":
                    _ = StartTurnAsync(
                        (string)payload?["text"] ?? "",
                        (string)payload?["reasoning"]);
                    break;
                case "stop":
                    try { _activeCts?.Cancel(); } catch { }
                    break;
                case "clear":
                    _store.Clear();
                    _ = RunScript("outlookai.clear();");
                    break;
                case "copy":
                    var clip = _store.ExportForClipboard();
                    try { Clipboard.SetText(clip ?? ""); } catch { }
                    break;
            }
        }

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "InboxCopilot");
            _isReady = true;
            _ = RunScript("outlookai.applyTheme('light');");
            PushReasoningOptions();
            PushContextStripAndChips();
            TraceLog.Write("OnWebViewReady completed", "InboxCopilot");
        }

        private void PushReasoningOptions()
        {
            try
            {
                var efforts = Config.ReasoningEffortsForModel(Config.Model);
                var arr = new JArray();
                foreach (var e in efforts) arr.Add(e);
                _ = RunScript("outlookai.setReasoningOptions(" +
                    arr.ToString(Newtonsoft.Json.Formatting.None) + ", '');");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushReasoningOptions error: " + ex.Message, "InboxCopilot");
            }
        }

        private void PushContextStripAndChips()
        {
            try
            {
                CurrentSelectionResult sel = null;
                string folderName = "";
                int unreadCount = 0, totalCount = 0;
                try
                {
                    sel = _surface.GetCurrentSelection(includeFullBodies: false, maxItems: 5);
                    folderName = sel?.Folder ?? "";
                    var folder = _explorer?.CurrentFolder;
                    try { if (folder != null) { unreadCount = folder.UnReadItemCount; totalCount = folder.Items.Count; } }
                    catch { }
                }
                catch (Exception ex) { TraceLog.Write("PushContextStrip surface error: " + ex.Message, "InboxCopilot"); }

                var ctx = new JObject(
                    new JProperty("folder", folderName),
                    new JProperty("unread_count", unreadCount),
                    new JProperty("total_count", totalCount));
                if (sel != null && sel.Count > 0 && sel.Messages != null && sel.Messages.Count > 0)
                {
                    var first = sel.Messages[0];
                    ctx.Add("selection", new JObject(
                        new JProperty("count", sel.Count),
                        new JProperty("subject", first.Subject ?? ""),
                        new JProperty("from", first.From ?? "")));
                }
                _ = RunScript("outlookai.setContextStrip(" +
                    ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");

                var selectionCount = sel?.Count ?? 0;
                var chips = QuickActionChip.ComputeChipsForSelectionCount(selectionCount);
                var chipsArr = new JArray();
                foreach (var c in chips)
                {
                    chipsArr.Add(new JObject(
                        new JProperty("label", c.Label),
                        new JProperty("prompt", c.Prompt)));
                }
                _ = RunScript("outlookai.setQuickActions(" +
                    chipsArr.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStripAndChips error: " + ex.Message, "InboxCopilot");
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxCopilot");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxCopilot");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("outlookai.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildSystemInstructionsForCurrentState(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    ReasoningEffortOverride = string.IsNullOrEmpty(reasoningOverride) ? null : reasoningOverride,
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                {
                    _store.Append(ctx.History[i]);
                }
                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                await RunScript("outlookai.showError(" + JsString(ex.Message ?? "") + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "InboxCopilot");
            }
        }

        private string BuildSystemInstructionsForCurrentState()
        {
            try
            {
                var sel = _surface?.GetCurrentSelection(includeFullBodies: false, maxItems: 1);
                int unreadCount = 0, totalCount = 0;
                string folderName = sel?.Folder ?? "Inbox";
                try
                {
                    var folder = _explorer?.CurrentFolder;
                    if (folder != null) { unreadCount = folder.UnReadItemCount; totalCount = folder.Items.Count; }
                }
                catch { }
                return InboxCopilotPromptBuilder.Build(folderName, unreadCount, totalCount, sel);
            }
            catch (Exception ex)
            {
                TraceLog.Write("BuildSystemInstructions error: " + ex, "InboxCopilot");
                return "You are the Outlook Inbox Copilot. Help the user with their mailbox.";
            }
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
                {
                    await marshaller.RunAsync(() =>
                    {
                        if (_isDisposed) return;
                        var core = _webView?.CoreWebView2;
                        if (core == null) return;
                        _ = core.ExecuteScriptAsync(script);
                    }, CancellationToken.None).ConfigureAwait(false);
                    return;
                }
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "InboxCopilot");
            }
        }

        private static string JsString(string s)
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try
            {
                if (_explorer != null)
                {
                    _explorer.SelectionChange -= OnExplorerSelectionChange;
                    _explorer.FolderSwitch -= OnExplorerFolderSwitch;
                }
            }
            catch { }
            try { _webView?.Dispose(); } catch { }
        }

        private sealed class WebViewSink : ChatEventSink
        {
            private readonly InboxCopilotController _owner;
            private readonly string _assistantId;
            public WebViewSink(InboxCopilotController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                _ = _owner.RunScript("outlookai.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                _ = _owner.RunScript("outlookai.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                _ = _owner.RunScript("outlookai.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
            }
            public override void OnError(string message)
            {
                _ = _owner.RunScript("outlookai.showError(" + JsString(message ?? "") + ");");
            }
        }
    }
}
```

- [ ] **Step 2: Add Compile entry**

In `OutlookAI.csproj`, next to the existing InboxCopilot entries:

```xml
<Compile Include="TaskPane\InboxCopilot\InboxCopilotController.cs" />
```

- [ ] **Step 3: Build**

Expected: clean. The earlier Task 12 cliffhanger now resolves.

- [ ] **Step 4: Full suite**

Expected: 134 / 134 (no new tests in this task — controller is integration-tested manually).

- [ ] **Step 5: Commit + push (Inbox Copilot pane functional checkpoint)**

```powershell
git add VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotController.cs VSTO2/OutlookAI/OutlookAI.csproj
git commit -m "Phase 3a Task 13: InboxCopilotController + Bind+InitializeAsync end-to-end"
git push origin feature/codex-oauth-migration
```

---

## Task 14: WebUI — chat.js `setQuickActions` + enhanced `setContextStrip` + index.html chip div + styles.css

**Files:**
- Modify: `VSTO2\OutlookAI\WebUI\index.html`
- Modify: `VSTO2\OutlookAI\WebUI\chat.js`
- Modify: `VSTO2\OutlookAI\WebUI\styles.css`

Adds the quick-action chip row to the HTML, the JS API, and CSS pill styling. Enhances `setContextStrip` to render the Inbox shape (folder + selection) in addition to the compose shape (subject + recipients + thread).

- [ ] **Step 1: Add the chip container to `index.html`**

Open `VSTO2\OutlookAI\WebUI\index.html`. Find the `<div id="messages" ...>...</div>` line. Immediately after the closing `</div>` of `#messages` (and before the `<div id="composer" ...>` line), insert:

```html
    <div id="quickActions" class="quick-actions"></div>
```

- [ ] **Step 2: Add the `setQuickActions` API + enhance `setContextStrip` in `chat.js`**

Open `VSTO2\OutlookAI\WebUI\chat.js`. Near the top, after the existing `$reasoning` DOM ref, add:

```javascript
var $quickActions = document.getElementById('quickActions');
```

Find the existing `setContextStrip:` method on the `api` object and replace it entirely with:

```javascript
setContextStrip: function(ctx) {
  ctx = ctx || {};
  // Inbox shape: folder + selection. Compose shape: subject + recipients + thread.
  if (ctx.folder !== undefined) {
    var unread = (ctx.unread_count != null) ? (' (' + ctx.unread_count + ' unread)') : '';
    $ctxSubject.textContent = 'In: ' + ctx.folder + unread;
    if (ctx.selection && ctx.selection.count > 0) {
      if (ctx.selection.count === 1) {
        $ctxRecipients.textContent = 'Selected: ' + (ctx.selection.subject || '') +
          (ctx.selection.from ? (' \u2014 ' + ctx.selection.from) : '');
      } else {
        $ctxRecipients.textContent = 'Selected: ' + ctx.selection.count + ' messages';
      }
    } else {
      $ctxRecipients.textContent = '';
    }
    $ctxThread.textContent = '';
    return;
  }
  // Compose shape (Phase 2 behaviour unchanged):
  $ctxSubject.textContent = ctx.subject ? ('Re: ' + ctx.subject) : 'New email';
  var recipients = (ctx.recipients || []).join(', ');
  $ctxRecipients.textContent = recipients ? ('To: ' + recipients) : '';
  $ctxThread.textContent = ctx.thread || '';
},
```

Then, on the same `api` object (anywhere is fine; convention is alongside the other setters), add:

```javascript
/**
 * Render the row of quick-action chips above the composer. Each chip
 * is { label, prompt }. Clicking a chip fills the textarea with the
 * prompt AND immediately sends, per the Phase 3a spec.
 */
setQuickActions: function(chips) {
  if (!$quickActions) return;
  while ($quickActions.firstChild) $quickActions.removeChild($quickActions.firstChild);
  (chips || []).forEach(function(chip) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'qa-chip';
    btn.textContent = chip.label;
    btn.title = chip.prompt;
    btn.addEventListener('click', function() {
      $input.value = chip.prompt;
      sendInput();   // existing helper that sends + clears the textarea
    });
    $quickActions.appendChild(btn);
  });
},
```

- [ ] **Step 3: Add chip styling to `styles.css`**

Open `VSTO2\OutlookAI\WebUI\styles.css`. Append at the end of the file (after the existing rules, before any closing brace):

```css
/* ---- Quick-action chip row (Phase 3a) ---- */
.quick-actions {
  flex: 0 0 auto;
  padding: 6px 10px 0 10px;
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  background: var(--panel);
  border-top: 1px solid var(--border);
}
.quick-actions:empty { padding: 0; border-top: none; }

.qa-chip {
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 4px 10px;
  background: var(--tool-bg);
  color: var(--text);
  font: 12px var(--font-stack);
  cursor: pointer;
}
.qa-chip:hover { background: var(--accent); color: white; border-color: var(--accent); }
.qa-chip:focus { outline: 2px solid var(--accent); outline-offset: -1px; }
```

The `:empty` rule collapses the row to zero height when there are no chips — important because the WebView is reused for compose-pane chat too (Phase 2), and compose-pane chat never calls `setQuickActions`.

- [ ] **Step 4: Build**

Expected: clean. (WebUI files are embedded resources; their content lands in the DLL.)

- [ ] **Step 5: Full suite**

Expected: 134 / 134 (no test changes for WebUI).

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/WebUI/index.html VSTO2/OutlookAI/WebUI/chat.js VSTO2/OutlookAI/WebUI/styles.css
git commit -m "Phase 3a Task 14: WebUI quick-action chips + Inbox-shape context strip"
```

---

## Task 15: `ThisAddIn` context-routes Explorer vs Inspector + per-Explorer lifecycle

**Files:**
- Modify: `VSTO2\OutlookAI\ThisAddIn.cs`

`ShowTaskPane` becomes a dispatcher. Adds `ShowExplorerTaskPane`. Hooks `Explorers.NewExplorer` so we can attach a `Close` handler per Explorer for pane cleanup.

- [ ] **Step 1: Refactor `ShowTaskPane`**

Open `VSTO2\OutlookAI\ThisAddIn.cs`. Locate `public void ShowTaskPane()`. Replace it with the dispatch + helper methods:

```csharp
public void ShowTaskPane()
{
    using (TraceLog.Scope("ShowTaskPane", "ThisAddIn"))
    try
    {
        var activeWindow = this.Application.ActiveWindow();
        if (activeWindow is Outlook.Inspector insp)
        {
            ShowInspectorTaskPane(insp);
            return;
        }
        if (activeWindow is Outlook.Explorer expl)
        {
            ShowExplorerTaskPane(expl);
            return;
        }
        TraceLog.Write("No Inspector or Explorer is the active window; ignoring", "ThisAddIn");
        System.Windows.Forms.MessageBox.Show(
            "Open Outlook to your Inbox or compose an email, then click AI Assistant.",
            "AI Assistant",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
        TraceLog.Write("ShowTaskPane error: " + ex, "ThisAddIn");
    }
}

private void ShowInspectorTaskPane(Outlook.Inspector inspector)
{
    foreach (CustomTaskPane pane in this.CustomTaskPanes)
    {
        if (pane.Window == inspector)
        {
            TraceLog.Write("Reusing existing Inspector CustomTaskPane", "ThisAddIn");
            if (!pane.Visible)
            {
                var existingControl = pane.Control as AITaskPane;
                existingControl?.ResetForNewEmail();
            }
            pane.Visible = !pane.Visible;
            return;
        }
    }
    TraceLog.Write("Creating new AITaskPane for Inspector", "ThisAddIn");
    var taskPaneControl = new AITaskPane();
    taskPaneControl.Bind(inspector);
    var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
    customTaskPane.Width = 340;
    customTaskPane.Visible = true;
}

private void ShowExplorerTaskPane(Outlook.Explorer explorer)
{
    foreach (CustomTaskPane pane in this.CustomTaskPanes)
    {
        if (pane.Window == explorer)
        {
            TraceLog.Write("Reusing existing Explorer CustomTaskPane (toggle visibility)", "ThisAddIn");
            pane.Visible = !pane.Visible;
            return;
        }
    }
    TraceLog.Write("Creating new InboxCopilotPane for Explorer", "ThisAddIn");
    var paneControl = new InboxCopilotPane();
    paneControl.Bind(explorer);
    var ctp = this.CustomTaskPanes.Add(paneControl, "AI Assistant", explorer);
    ctp.Width = 340;
    ctp.Visible = true;
}
```

- [ ] **Step 2: Add `using OutlookAI.TaskPane.InboxCopilot;` at the top of the file**

In `VSTO2\OutlookAI\ThisAddIn.cs`, add to the imports:

```csharp
using OutlookAI.TaskPane.InboxCopilot;
```

(`OutlookAI.TaskPane` is already imported for `AITaskPane`; we need the nested namespace.)

- [ ] **Step 3: Hook `Explorers.NewExplorer` for per-Explorer Close cleanup (defense in depth)**

In `ThisAddIn_Startup`, after the existing service-init block, add:

```csharp
try
{
    this.Application.Explorers.NewExplorer += OnNewExplorer;
    TraceLog.Write("Subscribed to Explorers.NewExplorer", "ThisAddIn");
}
catch (Exception ex)
{
    TraceLog.Write("Could not subscribe to Explorers.NewExplorer: " + ex, "ThisAddIn");
}
```

Add the handler near the other private methods in the same class:

```csharp
private void OnNewExplorer(Outlook.Explorer explorer)
{
    try
    {
        TraceLog.Write("NewExplorer raised", "ThisAddIn");
        // VSTO's CustomTaskPanes collection handles disposal when the
        // Explorer closes, so there is nothing critical to do here for
        // now. Keeping the subscription gives us a hook point if Phase
        // 3.x adds persistent state we need to flush.
    }
    catch (Exception ex) { TraceLog.Write("OnNewExplorer error: " + ex, "ThisAddIn"); }
}
```

- [ ] **Step 4: Build**

Expected: clean.

- [ ] **Step 5: Full suite**

Expected: 134 / 134.

- [ ] **Step 6: Commit**

```powershell
git add VSTO2/OutlookAI/ThisAddIn.cs
git commit -m "Phase 3a Task 15: ThisAddIn context-routes Inspector vs Explorer; ShowExplorerTaskPane"
```

---

## Task 16: `Ribbon.xml` adds AI Assistant group on `TabMail`

**Files:**
- Modify: `VSTO2\OutlookAI\Ribbon.xml`

- [ ] **Step 1: Add the new tab/group entry**

Open `VSTO2\OutlookAI\Ribbon.xml`. Find the existing `<tabs>` block:

```xml
<tabs>
  <tab idMso="TabNewMailMessage">
    <group id="AIAssistantGroup" label="AI Assistant" insertAfterMso="GroupClipboard">
      <button id="btnAIAssistant"
              label="AI Assistant"
              size="large"
              onAction="OnAIAssistantClick"
              supertip="Open the AI Writing Assistant to help proofread, revise, and write emails."
              imageMso="SmartArtChangeColorsGallery" />
    </group>
  </tab>
</tabs>
```

Replace it with the two-tab version:

```xml
<tabs>
  <tab idMso="TabNewMailMessage">
    <group id="AIAssistantGroup" label="AI Assistant" insertAfterMso="GroupClipboard">
      <button id="btnAIAssistant"
              label="AI Assistant"
              size="large"
              onAction="OnAIAssistantClick"
              supertip="Open the AI Writing Assistant to help proofread, revise, and write emails."
              imageMso="SmartArtChangeColorsGallery" />
    </group>
  </tab>
  <tab idMso="TabMail">
    <group id="AIAssistantExplorerGroup" label="AI Assistant" insertAfterMso="GroupMailMove">
      <button id="btnAIAssistantExplorer"
              label="AI Assistant"
              size="large"
              onAction="OnAIAssistantClick"
              supertip="Open the AI Assistant to chat with your mailbox - summarize, search, and act on messages."
              imageMso="SmartArtChangeColorsGallery" />
    </group>
  </tab>
</tabs>
```

Both buttons share the same `onAction="OnAIAssistantClick"`; `Ribbon.cs`'s existing callback already calls `Globals.ThisAddIn.ShowTaskPane()`, which Task 15 turned into the context dispatcher.

- [ ] **Step 2: Build**

Expected: clean. Ribbon XML is an embedded resource — no .cs changes needed.

- [ ] **Step 3: Full suite**

Expected: 134 / 134.

- [ ] **Step 4: Commit**

```powershell
git add VSTO2/OutlookAI/Ribbon.xml
git commit -m "Phase 3a Task 16: Ribbon adds AI Assistant group on TabMail (Explorer)"
```

---

## Task 17: Final E2E build + full test sweep + manual smoke + push

**Files:**
- Modify: `docs\superpowers\checklists\phase-2-smoke.md` (rename to `phase-2-and-3a-smoke.md`, append new section).

The previous smoke checklist was Phase-2-specific. Append a Phase 3a section so the next dogfood pass covers both.

- [ ] **Step 1: Append Phase 3a smoke section**

Open `docs\superpowers\checklists\phase-2-smoke.md`. At the very end of the file (after the existing closing line), append:

```markdown

---

# Phase 3a Manual Smoke (Inbox Copilot)

Run on a fresh Outlook session **after** the Phase 2 smoke section above
passes.

## Pre-flight

- [ ] Re-publish + reinstall via the elevated installer one-liner.
- [ ] Verify the AI Assistant button is on both ribbons:
  - Open a compose window -> the AI Assistant group on the compose
    ribbon is unchanged.
  - Close it. Look at the main Outlook Home tab (TabMail) -> an
    AI Assistant group now appears far right, after Move.

## Pane lifecycle

1. [ ] Click AI Assistant from the Inbox view -> a 340-px pane opens on
   the right. Single chat surface, no tabs.
2. [ ] Context strip line 1 shows `In: Inbox (<n> unread)`.
3. [ ] Selection line is hidden when no message is selected.
4. [ ] Click a message in the reading pane -> context strip refreshes,
   shows `Selected: <subject> -- <from>`. Chip row refreshes to add
   `Summarize this thread` and `Draft a reply`.
5. [ ] Ctrl-click a second message -> chip row updates to multi-select
   variants `Summarize all selected` / `Triage selected`.
6. [ ] Switch folders in the navigation pane -> context strip's folder
   line updates.
7. [ ] Open a second Explorer (File -> Open & Export -> New Window) ->
   click AI Assistant -> a second, independent pane opens. Conversations
   are NOT shared between the two.
8. [ ] Close the first Explorer -> its pane disappears. The second
   pane + conversation are untouched.

## Functional checks

9. [ ] Click the `What needs my attention?` chip -> prompt fills the
   textarea AND auto-sends. Streaming response appears. Tool cards
   reflect the actual search/list calls the model made.
10. [ ] Type "Find unread emails from <a known sender> with attachments
    from the last 7 days." -> the model issues exactly ONE
    `outlook_search_messages` call with all four structured fields
    populated. Verify via `%LOCALAPPDATA%\OutlookAI\trace.log` looking
    for the `BuildRunTurnRequest` line and the function_call argument
    payload.
11. [ ] Select one message in the reading pane, click `Summarize this
    thread` -> response mentions the actual subject line and at least
    one detail from the body. The model likely calls
    `outlook_get_current_selection` first (visible in tool cards), then
    `outlook_read_message` for other thread items.
12. [ ] Stop mid-stream -> partial assistant message keeps the
    "stopped" badge. Composer re-enables.
13. [ ] Clear button empties the chat.
14. [ ] Copy button puts the conversation on the clipboard
    (paste into Notepad to verify).

## Settings / model awareness

15. [ ] Open Settings -> change Model to `gpt-5.4` -> save. Close the
    pane (X). Reopen via the ribbon button. The reasoning dropdown now
    includes `Minimal` (gpt-5.4 supports it; gpt-5.5 does not).
16. [ ] Change Model back to `gpt-5.5` and confirm the dropdown drops
    `Minimal` on the next pane open.

## Lifecycle / cleanup

17. [ ] Close all Explorers -> the per-Explorer panes + conversations
    are gone.
18. [ ] Restart Outlook -> no leftover state visible. (Phase 3a is
    deliberately non-persistent.)

If any step fails, capture the trace log + screenshot and file as a
follow-up. Do NOT merge to master until every step passes.
```

- [ ] **Step 2: Rename the checklist file**

Phase 2's smoke checklist was named `phase-2-smoke.md`. Rename it to reflect that it now covers Phase 3a too:

```powershell
git mv docs/superpowers/checklists/phase-2-smoke.md docs/superpowers/checklists/phase-2-and-3a-smoke.md
```

- [ ] **Step 3: Clean rebuild from scratch + nuget restore**

```powershell
& "C:\Users\MDASR\AppData\Local\Temp\opencode\tools\nuget.exe" restore "VSTO2\OutlookAI\packages.config" -PackagesDirectory "VSTO2\OutlookAI\packages"
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /t:Rebuild /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal
```

Expected: clean rebuild, exit 0, zero warnings.

- [ ] **Step 4: Full test suite**

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: 134 / 134 (or whatever final count is — should match the Task 14 number; no tests added in 15-17).

- [ ] **Step 5: Verify all expected file changes are committed**

```powershell
git status --short
```

Expected: clean working tree (no uncommitted changes).

- [ ] **Step 6: Commit the checklist update**

```powershell
git add docs/superpowers/checklists/phase-2-and-3a-smoke.md
git commit -m "Phase 3a Task 17: append Inbox Copilot smoke checklist"
```

- [ ] **Step 7: Publish + reinstall using the existing elevated one-liner**

In an elevated PowerShell:

```powershell
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item "$env:LOCALAPPDATA\OutlookAI\trace.log" -ErrorAction SilentlyContinue

$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
$repo = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-codex-oauth-migration"
Remove-Item $staging -Recurse -Force
New-Item -ItemType Directory -Path $staging -Force | Out-Null

& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    "$repo\VSTO2\OutlookAI\OutlookAI.csproj" `
    /target:Publish /p:Configuration=Release /p:Platform="AnyCPU" `
    /p:PublishDir="$staging\" /v:minimal

Copy-Item "$repo\Deploy\Install-OutlookAI.ps1"           "$staging\" -Force
Copy-Item "$repo\Deploy\Uninstall-OutlookAI.ps1"         "$staging\" -Force
Copy-Item "$repo\Deploy\Enable-OutlookAI-User.ps1"       "$staging\" -Force -ErrorAction SilentlyContinue
Copy-Item "$repo\Deploy\Fetch-WebView2Bootstrapper.ps1"  "$staging\" -Force
Copy-Item "$repo\Deploy\MicrosoftEdgeWebView2Setup.exe"  "$staging\" -Force
Copy-Item "$repo\Deploy\README.txt"                      "$staging\" -Force -ErrorAction SilentlyContinue

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
& "$staging\Install-OutlookAI.ps1" -SourcePath $staging
```

Expected: installer exit 0; all 10 steps complete.

- [ ] **Step 8: Run the manual smoke checklist § Phase 3a** (the one we just appended in Step 1).

For each box you check, if anything fails, capture:
- Last ~50 lines of `%LOCALAPPDATA%\OutlookAI\trace.log`
- A screenshot of the misbehaving pane
- The exact click sequence

Then file a bug task on the branch.

- [ ] **Step 9: Final push**

```powershell
git push origin feature/codex-oauth-migration
```

End of Phase 3a plan. The next phase (3b — Calendar tools) starts with its own brainstorm + spec on top of this one.

---

## Self-Review Checklist (writing-plans skill discipline)

After completing all 17 tasks, the reviewer (you or a subagent) should be able to answer YES to each of:

- [ ] Every spec section in `2026-05-17-phase-3a-inbox-copilot-design.md` is implemented by at least one task.
- [ ] The ribbon button works on both compose and Explorer windows.
- [ ] Multi-Explorer is isolated — two Explorers, two conversations.
- [ ] `outlook_search_messages` accepts the structured fields and the wire body contains the AND-composed DASL filter.
- [ ] `outlook_get_current_selection` returns the user's actual selection.
- [ ] The chip row appears, refreshes on selection change, auto-sends on click.
- [ ] The system prompt includes folder + selection state each turn.
- [ ] Full unit test suite passes — target ~134.
- [ ] Manual smoke § Phase 3a passes end-to-end.
- [ ] Branch tip is published to `origin/feature/codex-oauth-migration`.
- [ ] `master` is unchanged from before this work.

If any answer is NO, return to that task and complete it before declaring Phase 3a done.



