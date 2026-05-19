# Phase 3a Search Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Inbox Copilot email search reliable and faster for oldest/newest, current-folder-first, all-mail broadening, and targeted filters without accidental default boolean constraints.

**Architecture:** Keep the existing tool names and Outlook COM surface. Add a small parser/normalizer for search args, extend `SearchMessagesArgs` with tri-state fields plus scope/sort, update DASL filter generation, and harden `LiveOutlookSurface` execution to support current-folder/all-mail/auto and oldest/newest ordering. No local mailbox index in this phase.

**Tech Stack:** C# .NET Framework 4.7.2, VSTO Outlook COM interop, Newtonsoft.Json, xUnit, MSBuild, Visual Studio Test Platform.

---

## File Structure

- Create `VSTO2/OutlookAI/Services/Tools/SearchMessagesArgsParser.cs`
  - One responsibility: parse JSON tool arguments into normalized `SearchMessagesArgs` for both search and count tools.
  - Owns blank-string normalization, enum normalization, old-field compatibility, max-result clamping, and date parsing.
- Modify `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
  - Extend `SearchMessagesArgs` with `Scope`, `SortOrder`, `AttachmentFilter`, `ReadStatus`, `FlagStatus`, `ImportanceFilter`.
  - Keep existing old fields for hidden compatibility.
- Modify `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`
  - Replace inline parsing with `SearchMessagesArgsParser.ParseSearch(argsJson)`.
- Modify `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs`
  - Replace duplicate inline parsing with `SearchMessagesArgsParser.ParseCount(argsJson)`.
- Modify `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
  - Update `BuildRestrictFilter` for tri-state fields and hidden old-field compatibility.
  - Add current-folder/all-mail/auto scope selection.
  - Add oldest/newest sort direction.
  - Remove diagnostic `items.Count` from `SearchMessages`.
  - Keep `CountMessages` counting but make it scope-aware.
- Modify `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
  - Advertise new fields and hide old booleans from model-facing schema.
- Modify `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPromptBuilder.cs`
  - Add examples for oldest/all-mail/current-folder/EIN behavior and tri-state filters.
- Modify tests under `VSTO2/OutlookAI.Tests/Services/Tools/`
  - Parser tests, filter tests, schema tests.
- Modify `VSTO2/OutlookAI.Tests/TaskPane/InboxCopilot/InboxCopilotPromptBuilderTests.cs`
  - Prompt examples test.

---

### Task 1: Extend SearchMessagesArgs And Add Parser RED Tests

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
- Create: `VSTO2/OutlookAI/Services/Tools/SearchMessagesArgsParser.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/SearchMessagesArgsParserTests.cs`

- [ ] **Step 1: Extend `SearchMessagesArgs` with the new normalized fields**

Open `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`. Replace the existing `SearchMessagesArgs` class with:

```csharp
public sealed class SearchMessagesArgs
{
    public string Query { get; set; }
    public string From { get; set; }
    public string SubjectContains { get; set; }
    public string BodyContains { get; set; }

    // Hidden backward-compat fields. These are accepted by the parser for
    // older conversation history but are no longer advertised to the model.
    public bool? HasAttachment { get; set; }
    public bool? IsUnread { get; set; }
    public bool? IsFlagged { get; set; }
    /// <summary>One of "low" | "normal" | "high"; null = unset.</summary>
    public string Importance { get; set; }

    public string FolderId { get; set; }
    public DateTimeOffset? DateFrom { get; set; }
    public DateTimeOffset? DateTo { get; set; }
    public int MaxResults { get; set; } = 25;

    /// <summary>"current_folder" | "all_mail" | "auto"; default auto.</summary>
    public string Scope { get; set; } = "auto";
    /// <summary>"newest" | "oldest"; default newest.</summary>
    public string SortOrder { get; set; } = "newest";
    /// <summary>"any" | "with" | "without"; default any.</summary>
    public string AttachmentFilter { get; set; } = "any";
    /// <summary>"any" | "read" | "unread"; default any.</summary>
    public string ReadStatus { get; set; } = "any";
    /// <summary>"any" | "flagged" | "unflagged"; default any.</summary>
    public string FlagStatus { get; set; } = "any";
    /// <summary>"any" | "low" | "normal" | "high"; default any.</summary>
    public string ImportanceFilter { get; set; } = "any";
}
```

- [ ] **Step 2: Write the parser test file before creating the parser**

Create `VSTO2/OutlookAI.Tests/Services/Tools/SearchMessagesArgsParserTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchMessagesArgsParserTests
    {
        [Fact]
        public void ParseSearch_DefaultLookingOldBooleans_DoNotPolluteFilters()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"query\":\"\","
                + "\"from\":\"\","
                + "\"subject_contains\":\"\","
                + "\"body_contains\":\"\","
                + "\"has_attachment\":false,"
                + "\"is_unread\":false,"
                + "\"is_flagged\":false,"
                + "\"importance\":\"normal\"}");

            Assert.Null(args.Query);
            Assert.Null(args.From);
            Assert.Null(args.SubjectContains);
            Assert.Null(args.BodyContains);
            Assert.Null(args.HasAttachment);
            Assert.Null(args.IsUnread);
            Assert.Null(args.IsFlagged);
            Assert.Null(args.Importance);
            Assert.Equal("auto", args.Scope);
            Assert.Equal("newest", args.SortOrder);
            Assert.Equal("any", args.AttachmentFilter);
            Assert.Equal("any", args.ReadStatus);
            Assert.Equal("any", args.FlagStatus);
            Assert.Equal("any", args.ImportanceFilter);
        }

        [Fact]
        public void ParseSearch_NewTriStateFields_AreNormalized()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"query\":\" EIN \","
                + "\"scope\":\"ALL_MAIL\","
                + "\"sort_order\":\"OLDEST\","
                + "\"attachment_filter\":\"WITH\","
                + "\"read_status\":\"UNREAD\","
                + "\"flag_status\":\"FLAGGED\","
                + "\"importance_filter\":\"HIGH\","
                + "\"date_to\":\"2020-01-01T00:00:00Z\","
                + "\"max_results\":999}");

            Assert.Equal("EIN", args.Query);
            Assert.Equal("all_mail", args.Scope);
            Assert.Equal("oldest", args.SortOrder);
            Assert.Equal("with", args.AttachmentFilter);
            Assert.Equal("unread", args.ReadStatus);
            Assert.Equal("flagged", args.FlagStatus);
            Assert.Equal("high", args.ImportanceFilter);
            Assert.Equal(new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero), args.DateTo);
            Assert.Equal(100, args.MaxResults);
        }

        [Fact]
        public void ParseSearch_OldBooleanTrueValues_RemainExplicitFilters()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"has_attachment\":true,"
                + "\"is_unread\":true,"
                + "\"is_flagged\":true,"
                + "\"importance\":\"high\"}");

            Assert.Equal(true, args.HasAttachment);
            Assert.Equal(true, args.IsUnread);
            Assert.Equal(true, args.IsFlagged);
            Assert.Equal("high", args.Importance);
        }

        [Fact]
        public void ParseSearch_NewExplicitNegativeTriStates_AreKept()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"attachment_filter\":\"without\","
                + "\"read_status\":\"read\","
                + "\"flag_status\":\"unflagged\","
                + "\"importance_filter\":\"normal\"}");

            Assert.Equal("without", args.AttachmentFilter);
            Assert.Equal("read", args.ReadStatus);
            Assert.Equal("unflagged", args.FlagStatus);
            Assert.Equal("normal", args.ImportanceFilter);
        }

        [Fact]
        public void ParseCount_DefaultsToIntMaxForMaxResults()
        {
            var args = SearchMessagesArgsParser.ParseCount("{\"scope\":\"all_mail\"}");

            Assert.Equal("all_mail", args.Scope);
            Assert.Equal(int.MaxValue, args.MaxResults);
        }
    }
}
```

- [ ] **Step 3: Run parser tests to verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchMessagesArgsParserTests"
```

Expected: build fails because `SearchMessagesArgsParser` does not exist yet. That is the correct RED failure.

- [ ] **Step 4: Create `SearchMessagesArgsParser` implementation**

Create `VSTO2/OutlookAI/Services/Tools/SearchMessagesArgsParser.cs`:

```csharp
using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class SearchMessagesArgsParser
    {
        public static SearchMessagesArgs ParseSearch(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            var maxResults = args["max_results"]?.Value<int>() ?? 25;
            if (maxResults < 1) maxResults = 1;
            if (maxResults > 100) maxResults = 100;
            return Parse(args, maxResults);
        }

        public static SearchMessagesArgs ParseCount(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            return Parse(args, int.MaxValue);
        }

        private static SearchMessagesArgs Parse(JObject args, int maxResults)
        {
            var search = new SearchMessagesArgs
            {
                Query = Clean(args["query"]),
                From = Clean(args["from"]),
                SubjectContains = Clean(args["subject_contains"]),
                BodyContains = Clean(args["body_contains"]),
                FolderId = Clean(args["folder_id"]),
                DateFrom = ParseDate(args["date_from"]),
                DateTo = ParseDate(args["date_to"]),
                MaxResults = maxResults,
                Scope = EnumOrDefault(args["scope"], "auto", "current_folder", "all_mail", "auto"),
                SortOrder = EnumOrDefault(args["sort_order"], "newest", "newest", "oldest"),
                AttachmentFilter = EnumOrDefault(args["attachment_filter"], "any", "any", "with", "without"),
                ReadStatus = EnumOrDefault(args["read_status"], "any", "any", "read", "unread"),
                FlagStatus = EnumOrDefault(args["flag_status"], "any", "any", "flagged", "unflagged"),
                ImportanceFilter = EnumOrDefault(args["importance_filter"], "any", "any", "low", "normal", "high"),
            };

            // Hidden old-shape compatibility: preserve true values only.
            // False was a model default in real traces and must not mean
            // "without/read/unflagged" unless the new tri-state says so.
            if (args["has_attachment"]?.Type == JTokenType.Boolean
                && args["has_attachment"].Value<bool>())
            {
                search.HasAttachment = true;
            }
            if (args["is_unread"]?.Type == JTokenType.Boolean
                && args["is_unread"].Value<bool>())
            {
                search.IsUnread = true;
            }
            if (args["is_flagged"]?.Type == JTokenType.Boolean
                && args["is_flagged"].Value<bool>())
            {
                search.IsFlagged = true;
            }
            var oldImportance = Clean(args["importance"]);
            if (oldImportance == "low" || oldImportance == "high")
            {
                search.Importance = oldImportance;
            }

            return search;
        }

        private static string Clean(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            var value = ((string)token)?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string EnumOrDefault(JToken token, string fallback, params string[] allowed)
        {
            var value = Clean(token);
            if (value == null) return fallback;
            value = value.ToLowerInvariant();
            foreach (var allowedValue in allowed)
            {
                if (value == allowedValue) return value;
            }
            return fallback;
        }

        private static DateTimeOffset? ParseDate(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
            return DateTimeOffset.TryParse(
                (string)token,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out value) ? value : (DateTimeOffset?)null;
        }
    }
}
```

- [ ] **Step 5: Run parser tests to verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchMessagesArgsParserTests"
```

Expected: all `SearchMessagesArgsParserTests` pass.

- [ ] **Step 6: Commit Task 1**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs VSTO2/OutlookAI/Services/Tools/SearchMessagesArgsParser.cs VSTO2/OutlookAI.Tests/Services/Tools/SearchMessagesArgsParserTests.cs
git commit -m "feat(search): normalize mailbox search args"
```

---

### Task 2: Replace Duplicate Tool Parsing With Shared Parser

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`
- Modify: `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs`

- [ ] **Step 1: Add failing tests that the tools use normalized fields**

Append to `OutlookSearchMessagesToolTests.cs` before the nested `Surface` class:

```csharp
[Fact]
public async Task Execute_UsesSharedParser_NewTriStateAndScopeFields()
{
    SearchMessagesArgs observed = null;
    var surface = new Surface
    {
        OnSearch = a => { observed = a; return new MessageSummary[0]; }
    };
    var tool = new OutlookSearchMessagesTool();
    var argsJson = "{"
        + "\"scope\":\"all_mail\","
        + "\"sort_order\":\"oldest\","
        + "\"read_status\":\"unread\","
        + "\"attachment_filter\":\"with\","
        + "\"flag_status\":\"flagged\","
        + "\"importance_filter\":\"high\"}";

    await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

    Assert.Equal("all_mail", observed.Scope);
    Assert.Equal("oldest", observed.SortOrder);
    Assert.Equal("unread", observed.ReadStatus);
    Assert.Equal("with", observed.AttachmentFilter);
    Assert.Equal("flagged", observed.FlagStatus);
    Assert.Equal("high", observed.ImportanceFilter);
}
```

Append to `OutlookCountMessagesToolTests.cs` before the nested `Surface` class:

```csharp
[Fact]
public async Task Execute_UsesSharedParser_ForScopeAndTriStates()
{
    SearchMessagesArgs observed = null;
    var surface = new Surface { OnCount = a => { observed = a; return 5; } };
    var tool = new OutlookCountMessagesTool();

    await tool.ExecuteAsync(
        "{\"scope\":\"all_mail\",\"read_status\":\"read\"}",
        surface, CancellationToken.None);

    Assert.Equal("all_mail", observed.Scope);
    Assert.Equal("read", observed.ReadStatus);
    Assert.Equal(int.MaxValue, observed.MaxResults);
}
```

- [ ] **Step 2: Run tool tests to verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookSearchMessagesToolTests|FullyQualifiedName~OutlookCountMessagesToolTests"
```

Expected: new tests fail because the tools still parse inline and do not populate new fields.

- [ ] **Step 3: Replace inline parsing in `OutlookSearchMessagesTool`**

In `OutlookSearchMessagesTool.ExecuteAsync`, replace lines that parse `JObject` and construct `SearchMessagesArgs` with:

```csharp
var search = SearchMessagesArgsParser.ParseSearch(argsJson);
```

Remove the now-unused `ParseDate` helper and unused `System.Globalization` import from `OutlookSearchMessagesTool.cs`.

- [ ] **Step 4: Replace inline parsing in `OutlookCountMessagesTool`**

In `OutlookCountMessagesTool.ExecuteAsync`, replace lines that parse `JObject` and construct `SearchMessagesArgs` with:

```csharp
var search = SearchMessagesArgsParser.ParseCount(argsJson);
```

Remove the now-unused `ParseDate` helper and unused `System.Globalization` import from `OutlookCountMessagesTool.cs`.

- [ ] **Step 5: Run tool tests to verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~OutlookSearchMessagesToolTests|FullyQualifiedName~OutlookCountMessagesToolTests|FullyQualifiedName~SearchMessagesArgsParserTests"
```

Expected: all filtered tests pass.

- [ ] **Step 6: Commit Task 2**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookSearchMessagesToolTests.cs VSTO2/OutlookAI.Tests/Services/Tools/OutlookCountMessagesToolTests.cs
git commit -m "refactor(search): share mailbox search arg parsing"
```

---

### Task 3: Update DASL Filter Builder For Tri-State Fields

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/BuildRestrictFilterTests.cs`

- [ ] **Step 1: Add failing filter tests**

Append these facts/theories to `BuildRestrictFilterTests.cs`:

```csharp
[Theory]
[InlineData("any", null)]
[InlineData("with", "urn:schemas:httpmail:hasattachment = 1")]
[InlineData("without", "urn:schemas:httpmail:hasattachment = 0")]
public void AttachmentFilter_MapsTriState(string value, string expected)
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { AttachmentFilter = value });
    if (expected == null)
        Assert.Null(f);
    else
        Assert.Contains(expected, f);
}

[Theory]
[InlineData("any", null)]
[InlineData("unread", "urn:schemas:httpmail:read = 0")]
[InlineData("read", "urn:schemas:httpmail:read = 1")]
public void ReadStatus_MapsTriState(string value, string expected)
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { ReadStatus = value });
    if (expected == null)
        Assert.Null(f);
    else
        Assert.Contains(expected, f);
}

[Theory]
[InlineData("any", null)]
[InlineData("flagged", "= 2")]
[InlineData("unflagged", "<> 2")]
public void FlagStatus_MapsTriState(string value, string expected)
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { FlagStatus = value });
    if (expected == null)
        Assert.Null(f);
    else
    {
        Assert.Contains("0x10900003", f);
        Assert.Contains(expected, f);
    }
}

[Theory]
[InlineData("any", null)]
[InlineData("low", "= 0")]
[InlineData("normal", "= 1")]
[InlineData("high", "= 2")]
public void ImportanceFilter_MapsTriState(string value, string expected)
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { ImportanceFilter = value });
    if (expected == null)
        Assert.Null(f);
    else
    {
        Assert.Contains("0x00170003", f);
        Assert.Contains(expected, f);
    }
}

[Fact]
public void OldDefaultFalseFields_DoNotCreateClauses()
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
    {
        HasAttachment = null,
        IsUnread = null,
        IsFlagged = null,
        Importance = null,
        AttachmentFilter = "any",
        ReadStatus = "any",
        FlagStatus = "any",
        ImportanceFilter = "any",
    });
    Assert.Null(f);
}

[Fact]
public void FirstEmailEverArgs_DoNotContainDefaultFilterClauses()
{
    var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
    {
        Scope = "all_mail",
        SortOrder = "oldest",
        MaxResults = 1,
        AttachmentFilter = "any",
        ReadStatus = "any",
        FlagStatus = "any",
        ImportanceFilter = "any",
    });

    Assert.Null(f);
}
```

- [ ] **Step 2: Run filter tests to verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~BuildRestrictFilterTests"
```

Expected: new tri-state tests fail because `BuildRestrictFilter` does not use new fields yet.

- [ ] **Step 3: Update `BuildRestrictFilter` bool/importance section**

In `LiveOutlookSurface.BuildRestrictFilter`, replace the attachment/read/flag/importance block with:

```csharp
var attachmentFilter = (args.AttachmentFilter ?? "any").Trim().ToLowerInvariant();
if (attachmentFilter == "with" || args.HasAttachment == true)
{
    clauses.Add("urn:schemas:httpmail:hasattachment = 1");
}
else if (attachmentFilter == "without")
{
    clauses.Add("urn:schemas:httpmail:hasattachment = 0");
}

var readStatus = (args.ReadStatus ?? "any").Trim().ToLowerInvariant();
if (readStatus == "unread" || args.IsUnread == true)
{
    clauses.Add("urn:schemas:httpmail:read = 0");
}
else if (readStatus == "read")
{
    clauses.Add("urn:schemas:httpmail:read = 1");
}

var flagStatus = (args.FlagStatus ?? "any").Trim().ToLowerInvariant();
if (flagStatus == "flagged" || args.IsFlagged == true)
{
    clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" = 2");
}
else if (flagStatus == "unflagged")
{
    clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" <> 2");
}

var importanceFilter = (args.ImportanceFilter ?? "any").Trim().ToLowerInvariant();
if (importanceFilter == "any" && !string.IsNullOrEmpty(args.Importance))
{
    importanceFilter = args.Importance.Trim().ToLowerInvariant();
}
if (importanceFilter == "low" || importanceFilter == "normal" || importanceFilter == "high")
{
    var val = importanceFilter == "low" ? "0" : importanceFilter == "normal" ? "1" : "2";
    clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x00170003\" = " + val);
}
```

Keep the existing query/from/subject/body/date clauses unchanged.

- [ ] **Step 4: Run filter tests to verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~BuildRestrictFilterTests|FullyQualifiedName~SearchMessagesArgsParserTests"
```

Expected: all filtered tests pass.

- [ ] **Step 5: Commit Task 3**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/BuildRestrictFilterTests.cs
git commit -m "fix(search): prevent default boolean filters"
```

---

### Task 4: Update Model-Facing Schema And Prompt Guidance

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs`
- Modify: `VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPromptBuilder.cs`
- Modify: `VSTO2/OutlookAI.Tests/TaskPane/InboxCopilot/InboxCopilotPromptBuilderTests.cs`

- [ ] **Step 1: Add failing schema tests for new fields and hidden old booleans**

Append to `ToolCatalogSchemaTests.cs`:

```csharp
[Fact]
public void SearchMessages_Schema_AdvertisesScopeSortAndTriStates_NotOldBooleans()
{
    var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var search = FindTool(tools, "outlook_search_messages");
    var props = (JObject)search["parameters"]["properties"];

    Assert.NotNull(props["scope"]);
    Assert.NotNull(props["sort_order"]);
    Assert.NotNull(props["attachment_filter"]);
    Assert.NotNull(props["read_status"]);
    Assert.NotNull(props["flag_status"]);
    Assert.NotNull(props["importance_filter"]);
    Assert.Null(props["has_attachment"]);
    Assert.Null(props["is_unread"]);
    Assert.Null(props["is_flagged"]);
    Assert.Null(props["importance"]);
}

[Fact]
public void SearchMessages_Description_IncludesOldestAndAllMailExamples()
{
    var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var search = FindTool(tools, "outlook_search_messages");
    var desc = (string)search["description"];

    Assert.Contains("first email ever", desc);
    Assert.Contains("scope:'all_mail'", desc);
    Assert.Contains("sort_order:'oldest'", desc);
    Assert.Contains("EIN", desc);
}

[Fact]
public void CountMessages_Schema_AdvertisesScopeAndTriStates()
{
    var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
    var count = FindTool(tools, "outlook_count_messages");
    var props = (JObject)count["parameters"]["properties"];

    Assert.NotNull(props["scope"]);
    Assert.NotNull(props["read_status"]);
    Assert.Null(props["is_unread"]);
}
```

Update existing description tests if their old assertions mention `is_unread:true` or `has_attachment:true`; the new expected strings should mention `read_status:'unread'` and `attachment_filter:'with'`.

- [ ] **Step 2: Add failing prompt test for oldest/all-mail guidance**

Append to `InboxCopilotPromptBuilderTests.cs`:

```csharp
[Fact]
public void Prompt_IncludesOldestAllMailAndTriStateSearchGuidance()
{
    var prompt = InboxCopilotPromptBuilder.Build("Inbox", 0, 0, null);

    Assert.Contains("first email ever", prompt);
    Assert.Contains("scope=all_mail", prompt);
    Assert.Contains("sort_order=oldest", prompt);
    Assert.Contains("read_status", prompt);
    Assert.Contains("attachment_filter", prompt);
    Assert.Contains("Do not use has_attachment:false", prompt);
}
```

- [ ] **Step 3: Run schema/prompt tests to verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ToolCatalogSchemaTests|FullyQualifiedName~InboxCopilotPromptBuilderTests"
```

Expected: new tests fail because schema/prompt still advertise old booleans and lack oldest/all-mail examples.

- [ ] **Step 4: Replace search/count schema fields**

In `ToolCatalogSchema.cs`, in both `outlook_search_messages` and `outlook_count_messages` property maps, remove:

```csharp
new JProperty("has_attachment",   new JObject(new JProperty("type","boolean"))),
new JProperty("is_unread",        new JObject(new JProperty("type","boolean"))),
new JProperty("is_flagged",       new JObject(new JProperty("type","boolean"))),
new JProperty("importance",       new JObject(new JProperty("type","string"),
                                  new JProperty("enum", new JArray("low","normal","high")))),
```

Add these fields near the other filters:

```csharp
new JProperty("scope", new JObject(new JProperty("type","string"),
                                  new JProperty("enum", new JArray("current_folder","all_mail","auto")),
                                  new JProperty("description","Default auto. Use all_mail directly for 'ever', 'any email', 'anywhere', 'everything', or 'all mail'. folder_id overrides scope."))),
new JProperty("sort_order", new JObject(new JProperty("type","string"),
                                       new JProperty("enum", new JArray("newest","oldest")),
                                       new JProperty("description","Default newest. Use oldest for 'first', 'earliest', or 'oldest'."))),
new JProperty("attachment_filter", new JObject(new JProperty("type","string"),
                                              new JProperty("enum", new JArray("any","with","without")),
                                              new JProperty("description","Use with only when the user asks for attachments; without only when they ask for no attachments; otherwise any."))),
new JProperty("read_status", new JObject(new JProperty("type","string"),
                                        new JProperty("enum", new JArray("any","read","unread")),
                                        new JProperty("description","Use unread/read only when explicitly requested; otherwise any."))),
new JProperty("flag_status", new JObject(new JProperty("type","string"),
                                        new JProperty("enum", new JArray("any","flagged","unflagged")),
                                        new JProperty("description","Use flagged/unflagged only when explicitly requested; otherwise any."))),
new JProperty("importance_filter", new JObject(new JProperty("type","string"),
                                              new JProperty("enum", new JArray("any","low","normal","high")),
                                              new JProperty("description","Use low/normal/high only when explicitly requested; otherwise any."))),
```

Update the search description to include:

```csharp
+ "User says 'what was my first email ever' -> {scope:'all_mail', sort_order:'oldest', max_results:1}. "
+ "User says 'latest email from Jane with an attachment' -> {from:'Jane', attachment_filter:'with', sort_order:'newest', max_results:1}. "
+ "User says 'find an email with EIN' -> search query:'EIN' with scope:'auto'; if zero, try query:'Employer Identification Number'. "
+ "Never send default false filters such as has_attachment:false or is_unread:false; use attachment_filter/read_status/flag_status/importance_filter and use 'any' when not requested. "
```

- [ ] **Step 5: Update Inbox Copilot prompt tips**

In `InboxCopilotPromptBuilder.Build`, update the Tips block with these lines:

```csharp
sb.AppendLine("- For 'first email ever', 'earliest', or 'oldest', call");
sb.AppendLine("  outlook_search_messages with scope=all_mail, sort_order=oldest,");
sb.AppendLine("  max_results=1, and no filters unless the user named one.");
sb.AppendLine("- For broad words like 'ever', 'any email', 'anywhere', 'everything',");
sb.AppendLine("  or 'all mail', use scope=all_mail. Otherwise use scope=auto so the");
sb.AppendLine("  tool searches the current folder first and broadens only if needed.");
sb.AppendLine("- Use read_status, attachment_filter, flag_status, and");
sb.AppendLine("  importance_filter. Do not use has_attachment:false,");
sb.AppendLine("  is_unread:false, is_flagged:false, or importance=normal as defaults.");
```

Keep existing concise-reply guidance.

- [ ] **Step 6: Run schema/prompt tests to verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~ToolCatalogSchemaTests|FullyQualifiedName~InboxCopilotPromptBuilderTests"
```

Expected: all filtered tests pass.

- [ ] **Step 7: Commit Task 4**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs VSTO2/OutlookAI.Tests/Services/Tools/ToolCatalogSchemaTests.cs VSTO2/OutlookAI/TaskPane/InboxCopilot/InboxCopilotPromptBuilder.cs VSTO2/OutlookAI.Tests/TaskPane/InboxCopilot/InboxCopilotPromptBuilderTests.cs
git commit -m "feat(search): advertise safe scope and filter fields"
```

---

### Task 5: Add Pure Search Execution Helpers And Tests

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
- Create: `VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs`

- [ ] **Step 1: Add failing helper tests**

Create `VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs`:

```csharp
using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchExecutionHelperTests
    {
        [Theory]
        [InlineData("oldest", false)]
        [InlineData("newest", true)]
        [InlineData(null, true)]
        [InlineData("bogus", true)]
        public void SortDescending_ReturnsExpectedDirection(string sortOrder, bool expected)
        {
            Assert.Equal(expected, LiveOutlookSurface.SortDescending(new SearchMessagesArgs { SortOrder = sortOrder }));
        }

        [Theory]
        [InlineData("Deleted Items", true)]
        [InlineData("Junk Email", true)]
        [InlineData("Drafts", true)]
        [InlineData("Outbox", true)]
        [InlineData("Sync Issues", true)]
        [InlineData("RSS Feeds", true)]
        [InlineData("Inbox", false)]
        [InlineData("Sent Items", false)]
        [InlineData("Archive", false)]
        [InlineData("Projects", false)]
        public void ShouldSkipAllMailFolder_ExcludesNoisyFolders(string name, bool expected)
        {
            Assert.Equal(expected, LiveOutlookSurface.ShouldSkipAllMailFolder(name));
        }

        [Fact]
        public void MergeAndSortSearchResults_HonorsOldestAndMaxResults()
        {
            var hits = new[]
            {
                new MessageSummary { Id = "new", ReceivedAt = DateTimeOffset.Parse("2024-01-01T00:00:00Z") },
                new MessageSummary { Id = "old", ReceivedAt = DateTimeOffset.Parse("2010-01-01T00:00:00Z") },
                new MessageSummary { Id = "mid", ReceivedAt = DateTimeOffset.Parse("2020-01-01T00:00:00Z") },
            };

            var merged = LiveOutlookSurface.MergeAndSortSearchResults(
                hits, new SearchMessagesArgs { SortOrder = "oldest", MaxResults = 2 });

            Assert.Equal(2, merged.Count);
            Assert.Equal("old", merged[0].Id);
            Assert.Equal("mid", merged[1].Id);
        }
    }
}
```

- [ ] **Step 2: Run helper tests to verify RED**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchExecutionHelperTests"
```

Expected: build fails because helper methods do not exist.

- [ ] **Step 3: Add helper methods to `LiveOutlookSurface`**

Add these internal static helpers near `BuildRestrictFilter` in `LiveOutlookSurface.cs`:

```csharp
internal static bool SortDescending(SearchMessagesArgs args)
{
    return !string.Equals(args?.SortOrder, "oldest", StringComparison.OrdinalIgnoreCase);
}

internal static bool ShouldSkipAllMailFolder(string folderName)
{
    if (string.IsNullOrWhiteSpace(folderName)) return true;
    var name = folderName.Trim();
    return string.Equals(name, "Deleted Items", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "Junk Email", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "Drafts", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "Outbox", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "Sync Issues", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "RSS Feeds", StringComparison.OrdinalIgnoreCase);
}

internal static IReadOnlyList<MessageSummary> MergeAndSortSearchResults(
    IEnumerable<MessageSummary> hits,
    SearchMessagesArgs args)
{
    var ordered = SortDescending(args)
        ? hits.OrderByDescending(m => m.ReceivedAt)
        : hits.OrderBy(m => m.ReceivedAt);
    return ordered.Take(Math.Max(1, args?.MaxResults ?? 25)).ToList();
}
```

Ensure `LiveOutlookSurface.cs` already has `using System.Linq;` and `using System.Collections.Generic;`.

- [ ] **Step 4: Run helper tests to verify GREEN**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~SearchExecutionHelperTests"
```

Expected: helper tests pass.

- [ ] **Step 5: Commit Task 5**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs VSTO2/OutlookAI.Tests/Services/Tools/SearchExecutionHelperTests.cs
git commit -m "test(search): pin scope sorting helpers"
```

---

### Task 6: Implement Scope-Aware SearchMessages And CountMessages

**Files:**
- Modify: `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`

- [ ] **Step 1: Add current-folder resolver helper**

In `LiveOutlookSurface.cs`, add these private helpers near `ResolveFolder`:

```csharp
private Outlook.MAPIFolder ResolveCurrentFolder()
{
    try
    {
        var folder = _explorer?.CurrentFolder as Outlook.MAPIFolder;
        if (folder != null) return folder;
    }
    catch (COMException) { }

    try { return _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); }
    catch (COMException) { return null; }
}

private IReadOnlyList<Outlook.MAPIFolder> ResolveSearchFolders(SearchMessagesArgs args, bool allMail)
{
    var folders = new List<Outlook.MAPIFolder>();
    if (!string.IsNullOrEmpty(args?.FolderId))
    {
        var folder = ResolveFolder(args.FolderId);
        if (folder != null) folders.Add(folder);
        return folders;
    }

    if (!allMail)
    {
        var folder = ResolveCurrentFolder();
        if (folder != null) folders.Add(folder);
        return folders;
    }

    try
    {
        foreach (Outlook.Store store in _application.Session.Stores)
        {
            WalkMailFolders(store.GetRootFolder(), folders, depth: 0);
        }
    }
    catch (COMException) { }
    return folders;
}

private void WalkMailFolders(Outlook.MAPIFolder folder, List<Outlook.MAPIFolder> results, int depth)
{
    if (folder == null || depth > MaxFolderDepth) return;
    string name = "";
    try { name = folder.Name ?? ""; } catch (COMException) { }
    if (!ShouldSkipAllMailFolder(name))
    {
        try
        {
            if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
            {
                results.Add(folder);
            }
        }
        catch (COMException) { }
    }

    try
    {
        foreach (Outlook.MAPIFolder child in folder.Folders)
        {
            WalkMailFolders(child, results, depth + 1);
        }
    }
    catch (COMException) { }
}
```

- [ ] **Step 2: Add single-folder search helper**

Add this private helper in `LiveOutlookSurface.cs`:

```csharp
private void SearchOneFolder(
    Outlook.MAPIFolder folder,
    SearchMessagesArgs args,
    string filter,
    List<MessageSummary> summaries)
{
    if (folder == null) return;
    Outlook.Items items;
    try
    {
        items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
        try { items.Sort("[ReceivedTime]", SortDescending(args)); } catch (COMException) { }

        int taken = 0;
        foreach (var obj in items)
        {
            if (taken >= args.MaxResults) break;
            var mi = obj as Outlook.MailItem;
            if (mi == null) continue;
            summaries.Add(new MessageSummary
            {
                Id = _ids.Shorten(mi.EntryID),
                Subject = mi.Subject ?? "",
                From = mi.SenderName ?? mi.SenderEmailAddress ?? "",
                To = SplitAddresses(mi.To),
                ReceivedAt = ToOffset(mi.ReceivedTime),
                Snippet = SnippetOf(mi.Body),
                HasAttachments = mi.Attachments?.Count > 0,
            });
            taken++;
        }
    }
    catch (COMException ex)
    {
        OutlookAI.Diagnostics.TraceLog.Write(
            "SearchMessages folder=" + (folder.Name ?? "") + " failed: " + ex.Message,
            "LiveOutlookSurface");
    }
}
```

- [ ] **Step 3: Replace `SearchMessages` implementation body**

Replace the current `SearchMessages` body with scope-aware logic:

```csharp
public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args) =>
    Run(() =>
    {
        args = args ?? new SearchMessagesArgs();
        var filter = BuildRestrictFilter(args);
        var scope = (args.Scope ?? "auto").Trim().ToLowerInvariant();
        var searchedFolders = 0;
        var broadened = false;

        try
        {
            var currentHits = new List<MessageSummary>();
            var searchAllMail = scope == "all_mail";
            if (scope == "auto" && string.IsNullOrEmpty(args.FolderId))
            {
                foreach (var folder in ResolveSearchFolders(args, allMail: false))
                {
                    searchedFolders++;
                    SearchOneFolder(folder, args, filter, currentHits);
                }
                if (currentHits.Count > 0)
                {
                    var result = MergeAndSortSearchResults(currentHits, args);
                    TraceSearch(args, filter, searchedFolders, result.Count, broadened);
                    return result;
                }
                searchAllMail = true;
                broadened = true;
            }

            var allHits = new List<MessageSummary>();
            foreach (var folder in ResolveSearchFolders(args, allMail: searchAllMail))
            {
                searchedFolders++;
                SearchOneFolder(folder, args, filter, allHits);
            }

            var merged = MergeAndSortSearchResults(allHits, args);
            TraceSearch(args, filter, searchedFolders, merged.Count, broadened);
            return merged;
        }
        catch (COMException ex)
        {
            OutlookAI.Diagnostics.TraceLog.Write(
                "SearchMessages outer COMException: " + ex.Message,
                "LiveOutlookSurface");
            return (IReadOnlyList<MessageSummary>)new List<MessageSummary>();
        }
    }) ?? (IReadOnlyList<MessageSummary>)new List<MessageSummary>();
```

- [ ] **Step 4: Add `TraceSearch` helper**

Add:

```csharp
private static void TraceSearch(SearchMessagesArgs args, string filter, int foldersSearched, int returned, bool broadened)
{
    try
    {
        OutlookAI.Diagnostics.TraceLog.Write(
            "SearchMessages scope=" + (args.Scope ?? "auto")
            + " sort_order=" + (args.SortOrder ?? "newest")
            + " filter=" + (filter ?? "<none>")
            + " folders_searched=" + foldersSearched
            + " returned=" + returned
            + " broadened=" + broadened
            + " maxResults=" + args.MaxResults,
            "LiveOutlookSurface");
    }
    catch { }
}
```

- [ ] **Step 5: Replace `CountMessages` implementation body**

Replace `CountMessages` with:

```csharp
public int CountMessages(SearchMessagesArgs args) =>
    Run(() =>
    {
        args = args ?? new SearchMessagesArgs();
        var filter = BuildRestrictFilter(args);
        var scope = (args.Scope ?? "auto").Trim().ToLowerInvariant();
        var countAllMail = scope == "all_mail";
        var total = 0;
        try
        {
            foreach (var folder in ResolveSearchFolders(args, allMail: countAllMail))
            {
                Outlook.Items items = string.IsNullOrEmpty(filter)
                    ? folder.Items
                    : folder.Items.Restrict(filter);
                total += items.Count;
            }
        }
        catch (COMException) { }
        return total;
    });
```

Note: `scope=auto` for count stays current-folder only. Broadening counts automatically can produce surprising huge totals and is not needed for the “first email” path because the prompt/schema will steer that to `outlook_search_messages` with `scope=all_mail`.

- [ ] **Step 6: Build and run all unit tests**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: all tests pass.

- [ ] **Step 7: Commit Task 6**

Run:

```powershell
git add VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs
git commit -m "feat(search): support scope and oldest sorting"
```

---

### Task 7: Improve Dispatch-Time Search Diagnostics

**Files:**
- Modify: `VSTO2/OutlookAI/Services/CodexChatService.cs`
- Modify: `VSTO2/OutlookAI.Tests/Services/CodexChatServiceMultiRoundTests.cs`

- [ ] **Step 1: Add regression test that long search args are preserved in dispatch history**

Append to `CodexChatServiceMultiRoundTests.cs` near the SSE delta tests:

```csharp
[Fact]
public async Task RunTurnAsync_LongSearchArgs_AreEchoedFullyInFunctionCallHistory()
{
    var fixt = MakeAuth();
    var fake = new FakeHttpMessageHandler();
    fake.QueueSse(HttpStatusCode.OK,
        "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"item_long\",\"call_id\":\"call_long\",\"name\":\"outlook_search_messages\",\"arguments\":\"\"}}\n\n"
        + "data: {\"type\":\"response.function_call_arguments.done\",\"item_id\":\"item_long\",\"arguments\":\"{\\\"query\\\":\\\"Employer Identification Number\\\",\\\"scope\\\":\\\"all_mail\\\",\\\"sort_order\\\":\\\"oldest\\\",\\\"attachment_filter\\\":\\\"any\\\",\\\"read_status\\\":\\\"any\\\",\\\"flag_status\\\":\\\"any\\\",\\\"importance_filter\\\":\\\"any\\\",\\\"date_to\\\":\\\"2020-01-01T00:00:00Z\\\",\\\"max_results\\\":1}\"}\n\n"
        + "data: {\"type\":\"response.completed\"}\n\n");
    fake.QueueSse(HttpStatusCode.OK,
        "data: {\"type\":\"response.output_text.delta\",\"delta\":\"done\"}\n\n"
        + "data: {\"type\":\"response.completed\"}\n\n");
    try
    {
        using (fixt.AuthHttp)
        using (fixt.Auth)
        using (var chatHttp = new HttpClient(fake))
        using (var chat = new CodexChatService(fixt.Auth, chatHttp))
        {
            var ctx = new ConversationContext();
            var sink = new CapturingChatEventSink();
            var tools = new FakeToolHost();
            tools.Queue("outlook_search_messages", "{\"messages\":[]}");

            await chat.RunTurnAsync(ctx, "search", tools, sink, CancellationToken.None);

            var fc = ctx.History.Find(it => (string)it["type"] == "function_call");
            Assert.Contains("Employer Identification Number", (string)fc["arguments"]);
            Assert.Contains("\"scope\":\"all_mail\"", (string)fc["arguments"]);
            Assert.Contains("\"sort_order\":\"oldest\"", (string)fc["arguments"]);
        }
    }
    finally
    {
        try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
    }
}
```

- [ ] **Step 2: Expand trace truncation in `DispatchOneAsync`**

In `CodexChatService.DispatchOneAsync`, replace the hard-coded 200-char truncation with 500:

```csharp
const int maxTraceArgs = 500;
OutlookAI.Diagnostics.TraceLog.Write(
    "Dispatch " + name + " call_id=" + callId
    + " args=" + (args.Length > maxTraceArgs ? args.Substring(0, maxTraceArgs) + "..." : args),
    "CodexChat");
```

- [ ] **Step 3: Run Codex chat tests**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj" /p:Configuration=Debug /v:minimal /nologo
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll" /TestCaseFilter:"FullyQualifiedName~CodexChatServiceMultiRoundTests"
```

Expected: all Codex chat tests pass.

- [ ] **Step 4: Commit Task 7**

Run:

```powershell
git add VSTO2/OutlookAI/Services/CodexChatService.cs VSTO2/OutlookAI.Tests/Services/CodexChatServiceMultiRoundTests.cs
git commit -m "diag(search): expand tool argument tracing"
```

---

### Task 8: Full Verification, Publish, Install, And Smoke Test

**Files:**
- No source modifications expected.
- Use: `docs/superpowers/checklists/phase-2-and-3a-smoke.md`

- [ ] **Step 1: Run full Debug build**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
```

Expected: build succeeds with `OutlookAI.dll` and `OutlookAI.Tests.dll` produced under `bin\Debug`.

- [ ] **Step 2: Run full unit test suite**

Run:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" "VSTO2\OutlookAI.Tests\bin\Debug\net472\OutlookAI.Tests.dll"
```

Expected: all tests pass. Expected count is existing `149` plus new tests from Tasks 1-7.

- [ ] **Step 3: Publish Release build**

Run:

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
if (-not (Test-Path $staging)) { New-Item -ItemType Directory -Path $staging | Out-Null }
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" /p:PublishDir="$staging\" /v:minimal /nologo
```

Expected: fresh `OutlookAI.dll`, `OutlookAI.vsto`, and `setup.exe` under `$staging`.

- [ ] **Step 4: Install elevated with correct SourcePath**

Confirm Outlook is closed:

```powershell
Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
```

Run elevated installer:

```powershell
$staging = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2"
$installer = Join-Path $staging "Install-OutlookAI.ps1"
Start-Process powershell.exe -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","`"$installer`"","-SourcePath","`"$staging`"" -Verb RunAs -Wait -PassThru
```

Expected: accept UAC; installer prints `Installation Complete!`.

Verify installed DLL hash matches staged DLL:

```powershell
$installed = "C:\Program Files\OutlookAI\OutlookAI.dll"
$staged = "C:\Users\MDASR\AppData\Local\Temp\opencode\OutlookAI-publish-phase2\OutlookAI.dll"
(Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash
```

Expected: `True`.

- [ ] **Step 5: Manual smoke tests in Outlook**

Open Outlook, then use Inbox Copilot with these queries:

```text
What was my first email ever?
Find an email with EIN
Find invoices from before 2020
Latest email from Jane with an attachment
Show unread emails from last week
Search all mail for Cisco UC560
```

Expected trace checks in `%LOCALAPPDATA%\OutlookAI\trace.log`:

```text
(CodexChat) Dispatch outlook_search_messages ... "scope":"all_mail" ... "sort_order":"oldest" ... "max_results":1
(LiveOutlookSurface) SearchMessages scope=all_mail sort_order=oldest filter=<none> ... returned=1
```

For “first email ever,” trace must NOT contain these clauses unless explicitly requested:

```text
hasattachment = 0
urn:schemas:httpmail:read = 1
0x10900003" <> 2
0x00170003" = 1
```

- [ ] **Step 6: Push all commits**

Run:

```powershell
git status -sb
git push origin feature/codex-oauth-migration
```

Expected: branch is pushed and clean.

---

## Plan Self-Review

- Spec coverage: covered parser/schema/filter/scope/sort/diagnostics/tests/rollout from `2026-05-17-phase-3a-search-hardening-design.md`.
- Placeholder scan: no unfinished-marker steps; each code-changing step includes exact code or replacement text.
- Type consistency: field names match the design: `Scope`, `SortOrder`, `AttachmentFilter`, `ReadStatus`, `FlagStatus`, `ImportanceFilter`; schema names are snake_case counterparts.
- Scope check: no local index, attachment content indexing, delete/send, or UI redesign included.
