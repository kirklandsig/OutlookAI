# Phase 3b: Non-Blocking Mailbox Search — Design

Branch: `feature/codex-oauth-migration`
Author session: 2026-05-17
Status: Proposed (pending user review)

## Why

Phase 3a fixed search *correctness* (structured args, sentinel-date normalization,
explicit scope/sort/tri-state filters). It did not fix *responsiveness*. A live
trace of `What is my oldest email?` shows every folder being walked on the
Outlook UI thread (`T 1 UI`) inside a single `OutlookThreadMarshaller.RunAsync<T>`
call. Outlook cannot pump UI messages until the entire search finishes:

```
[68316ms UI] folder_done=6 Month Reviews elapsed_ms=25
...
[71263ms UI] folder_done=Sent Items     elapsed_ms=2082
[80296ms UI] folder_done=Inbox          elapsed_ms=5557
...
```

12+ seconds of UI thread held in one call, with the search still not finished.
From the user's perspective Outlook is frozen.

The root cause is not "search is slow"; it is "search holds the UI thread
continuously." Any tool that does a long sequence of Outlook OOM calls on the
UI thread will have the same problem. We must change the architecture so that
long mailbox work no longer blocks the UI thread continuously.

## Goal

1. Outlook UI never freezes for the duration of a `outlook_search_messages`
   or `outlook_count_messages` call.
2. Searches use Outlook's native search engine where possible (same engine the
   Outlook search box uses), with the indexed-search performance benefit.
3. Searches are cancellable: there is a real Stop affordance, and cancellation
   actually halts the in-progress work.
4. Results stay consistent with Outlook's own search behavior so the model's
   "find emails about X" matches the user's intuition.
5. The fallback path keeps the tool functional on stores where the native
   engine fails, stalls, or returns nothing.

## Non-goals

- Building our own local mailbox index.
- Replacing `OutlookThreadMarshaller`. Other tools still need UI-thread
  marshalling for their short COM calls.
- Changing the Codex tool schema. `outlook_search_messages` keeps its current
  args (scope, sort_order, tri-states, etc.) and its current JSON output.
- Changing other Outlook tools (selection, read, draft, flag, category, etc.).
- Streaming partial results to the model mid-search. AdvancedSearch completes
  once and we return once.

## Approach (recommended)

**Primary engine: Outlook `Application.AdvancedSearch`.**
`AdvancedSearch` is the same API Outlook's UI uses. It runs asynchronously in
Outlook's own search infrastructure, fires `Application.AdvancedSearchComplete`
when done, and supports cancellation via `Search.Stop()`. The UI thread is only
held for the brief invocation and the brief result projection.

**Fallback engine: the existing iterative per-folder path, refactored to yield.**
If AdvancedSearch fails to fire (`COMException` on invoke, timeout, null
results, etc.), we fall back to walking folders, but with one folder per
marshalled call and a real `CancellationToken`. The UI thread is released
between folders, so Outlook can pump messages.

Both paths feed the same projection (`MessageSummary[]`) and the same skip
list, so behavior stays consistent and only the engine swaps.

## Architecture

### New components

**`OutlookAdvancedSearchRunner`** (new, in `OutlookAI/Services/Tools/`)
Encapsulates the COM event plumbing. Public surface:

```csharp
public interface IOutlookAdvancedSearchRunner : IDisposable
{
    Task<AdvancedSearchResult> RunAsync(
        string scope,        // DASL-compatible scope string
        string filter,       // DASL filter, may be null/empty
        bool searchSubFolders,
        TimeSpan timeout,
        CancellationToken ct);
}

public sealed class AdvancedSearchResult
{
    public bool Completed { get; set; }   // true if AdvancedSearchComplete fired
    public bool TimedOut { get; set; }
    public bool Cancelled { get; set; }
    public Exception Error { get; set; }  // non-null if invoke threw
    // Raw Outlook Search.Results, or null on failure. Caller projects.
    public Outlook.Results Results { get; set; }
}
```

Implementation details:

- Subscribes once to `Application.AdvancedSearchComplete`.
- Dispatches by `Tag` (GUID per call).
- Maintains a `ConcurrentDictionary<string, TaskCompletionSource<AdvancedSearchResult>>`
  keyed by Tag.
- On `RunAsync`:
  - Marshals to UI thread.
  - Calls `_application.AdvancedSearch(scope, filter, searchSubFolders, tag)`.
  - Stores TCS keyed by tag.
  - Returns a task that completes when the event fires, the timeout elapses, or
    the CT cancels (whichever first).
  - On timeout / cancellation: marshals to UI thread and calls
    `search.Stop()`, then completes the TCS.
- Concurrency: serialised behind an `AsyncLock` (or `SemaphoreSlim(1,1)`) so
  Outlook never sees more than one in-flight AdvancedSearch from this add-in
  at a time. This avoids the documented "too many in-flight searches" failure
  mode and keeps event dispatch simple.

**`SearchScopeBuilder`** (new)
Converts `SearchMessagesArgs` + the Outlook session into an AdvancedSearch
scope string (comma-separated, each path single-quoted) plus a
`searchSubFolders` flag.

```csharp
public static class SearchScopeBuilder
{
    public static SearchScope Build(
        Outlook.Application application,
        Outlook.Explorer explorer,
        SearchMessagesArgs args,
        IFolderClassifier classifier);
}

public sealed class SearchScope
{
    public string ScopeString { get; set; }
    public bool SearchSubFolders { get; set; }
    public IReadOnlyList<string> ResolvedFolderPaths { get; set; } // for tracing
}
```

Rules:

- `args.FolderId` set → resolve to that folder's `FolderPath`. SearchSubFolders
  = true.
- `Scope == "current_folder"` → explorer's current folder path.
  SearchSubFolders = false (the user is asking about this folder, not its tree).
- `Scope == "all_mail"` (or `"auto"` after broaden) →
  - Enumerate all stores.
  - For each store, take its root folder.
  - Build a single comma-separated scope of all store roots, with
    SearchSubFolders = true.
  - The `IFolderClassifier` is *not* applied to the scope itself; instead,
    skip-list filtering happens during result projection (see Filtering below).
  - Reason: AdvancedSearch performs better with a single scope entry plus
    subfolder recursion than with hundreds of individual leaf-folder entries.

**`IFolderClassifier`** (new)
Decides whether a folder is "noise" (a system folder we never want to surface
results from). Replaces the current `ShouldSkipAllMailFolder` helper.

Real folder names observed in the trace and from Outlook documentation:

- `Deleted Items` (any locale variation), `Junk E-mail`, `Junk Email`
- `Drafts`, `Outbox`
- `Conflicts`, `Local Failures`, `Server Failures`, `Sync Issues`
  (and `Sync Issues (This computer only)`)
- `RSS Feeds`, `RSS Subscriptions`
- `Conversation Action Settings`, `Conversation History`
- `Quick Step Settings`, `Quick Step History`
- `News Feed`, `Feeds`, `Files`, `Detected Items`
- `Working Set`, `Yammer Root`
- Anything with `DefaultItemType != olMailItem` (Calendar, Contacts, etc.)

Classifier returns one of `{Include, ExcludeAlways}`. The list is centralised
so both the AdvancedSearch projection and the fallback iterative path see the
same skip behavior.

**`SearchResultProjector`** (new)
Converts an `Outlook.Results` (or a sequence of `Outlook.MailItem`) into
`MessageSummary[]` with:

- Sort by `[ReceivedTime]` desc/asc per `args.SortOrder`.
- Drop items whose parent folder is `ExcludeAlways`.
- Take top `args.MaxResults`.
- Defer `mi.Body` snippet to ONLY top-N items (avoid paying body-load cost
  for items we drop).
- Resilient: `COMException` per item → log, skip, continue.

### Changes to existing components

**`IOutlookSurface.SearchMessages`**
Add an optional `CancellationToken`. Default `CancellationToken.None`
preserves all existing callers.

```csharp
IReadOnlyList<MessageSummary> SearchMessages(
    SearchMessagesArgs args,
    CancellationToken ct = default);
```

Same change to `CountMessages` for symmetry. Count uses the same engine but
projects to a count of matching items (post-skip-list).

**`LiveOutlookSurface.SearchMessages`**
Re-implemented as:

```text
1. Build filter (existing BuildRestrictFilter, unchanged).
2. Build scope via SearchScopeBuilder.
3. Try primary:
     result = advancedSearchRunner.RunAsync(scope, filter, searchSubFolders,
                                             timeout, ct);
     if result.Completed && result.Results != null:
         return projector.Project(result.Results, args, classifier);
4. Log fallback reason. If cancelled by user (not timeout), return
   projector.Project(<empty>) and exit.
5. Fallback: per-folder iterative path:
     - resolve folder list (same rules as today).
     - for each folder:
         marshaller.RunAsync(() => SearchOneFolder(folder, ...), ct).
         ct.ThrowIfCancellationRequested between folders.
     - merge + sort + project (deferred snippet).
6. Return.
```

The yielding fallback is the *only* place the iterative code remains, and it
uses one `RunAsync` per folder, so the UI thread is released between folders.

**`OutlookSearchMessagesTool` / `OutlookCountMessagesTool`**
Plumb the existing tool `CancellationToken` through to `IOutlookSurface`.
No JSON schema change.

**`OutlookThreadMarshaller`**
No behavior change. We use it as-is for both per-folder calls and for the
brief AdvancedSearch invocation.

### Cancellation chain

```
User clicks "Stop" in compact status UI
    ↓
ChatSession.CancelCurrent()  (existing CTS)
    ↓
ToolDispatcher CT  (existing)
    ↓
OutlookSearchMessagesTool.ExecuteAsync(ct)
    ↓
LiveOutlookSurface.SearchMessages(args, ct)
    ├─ Primary: OutlookAdvancedSearchRunner observes ct
    │           → Marshals search.Stop() to UI thread
    │           → Completes TCS as Cancelled
    └─ Fallback: per-folder loop observes ct between folders
                 → Breaks out, returns accumulated partial results
```

### Status / progress UI

In scope for this phase. The existing compact tool-status UI gains:

- A small "(Stop)" link/button rendered inline with the status text. Clicking
  it calls `ChatSession.CancelCurrent()`.
- A time counter in the existing status text ("Searching mailbox (12s)…")
  updated from a `Progress<TimeSpan>` callback the runner pulses every
  ~500ms while a search is in flight.

The engine change is the responsiveness fix; the Stop affordance is the
user-control fix. We commit to both in this phase because the cancellation
chain already terminates at `ChatSession.CancelCurrent()` and the WebView2
change to surface the link is small.

## Filter (DASL) syntax

AdvancedSearch's `Filter` parameter uses the same DASL we already build in
`BuildRestrictFilter` for `Items.Restrict`. No filter-language change. The
DASL builder is reused unchanged.

For the "oldest with no model-provided filter" case (today's failing
scenario), filter is `null`/empty and AdvancedSearch returns every item in
scope. We sort+top-N on `Results`, taking only the first item. The
projector reads `Body` only for that one item.

## Scope syntax

Outlook expects scope as a comma-separated list of single-quoted folder
paths:

```
"'\\Mailbox - User\\Inbox','\\Archive PST\\Sent Items'"
```

`SearchScopeBuilder` builds exactly that and never accepts user input
verbatim — the model never sees a raw scope string, only the enum
`current_folder | all_mail | auto`.

## Error handling

| Failure | Behavior |
|---|---|
| `COMException` from invoking `AdvancedSearch()` | Log, fall back. |
| `AdvancedSearchComplete` never fires within timeout | Log, `search.Stop()`, fall back. |
| `Complete` fires with `Search.Results == null` | Log, fall back. |
| Tag mismatch / event for unknown tag | Log, ignore (defensive). |
| User CT cancellation during primary | `search.Stop()`, surface returns empty results, tool emits `{"error":{"code":"cancelled","message":"Search cancelled by user."}}`, do NOT fall back. |
| `COMException` per item during projection | Log, skip item, continue. |
| Fallback per-folder `COMException` | Log, skip folder, continue. |
| Fallback per-folder CT cancellation | Break, surface returns empty results, tool emits the same `{"error":{"code":"cancelled",...}}` envelope. |

Default timeout: 30 seconds. Configurable in `config.xml` later if needed;
hardcoded for the first ship.

## Tracing

Add structured entries (replace ad-hoc traces from today's diag commit):

```
SearchMessages primary=AdvancedSearch start scope_count=N filter=<...>
SearchMessages primary=AdvancedSearch complete elapsed_ms=... raw_count=...
SearchMessages primary=AdvancedSearch fallback reason={timeout|com|null_results|...}
SearchMessages fallback=Iterative folder=... taken=... elapsed_ms=...
SearchMessages return count=... sort=... scope=...
```

Existing per-folder progress traces stay in the fallback path (they were
useful evidence; they remain useful for fallback observability).

## Testing strategy

All tests run without a live Outlook. Existing tests for parser, filter
builder, schema, prompt builder, etc. continue to pass.

New tests:

**`OutlookAdvancedSearchRunnerTests`** (using an `IAdvancedSearchHost`
test seam):

- Happy path: invoke → host raises `Complete` with a fake Results → task
  returns Results.
- Timeout: invoke → no event → after timeout, task returns `TimedOut=true`,
  host's `Stop` was called.
- Cancellation: invoke → CT triggered → task returns `Cancelled=true`,
  host's `Stop` was called.
- COMException on invoke: task returns `Error != null`, no event subscribed.
- Multiple concurrent calls correctly serialise (semaphore behaviour).
- Tag mismatch event is ignored without breaking pending tasks.

**`SearchScopeBuilderTests`** (using fake Outlook tree):

- `current_folder` returns single quoted path, SearchSubFolders=false.
- `all_mail` enumerates stores, builds combined scope, SearchSubFolders=true.
- `folder_id` set wins over scope value.

**`FolderClassifierTests`**:

- All known system folder names → ExcludeAlways.
- Non-mail-item folders → ExcludeAlways.
- "Junk E-mail" with hyphen → ExcludeAlways (today's bug).
- User folders ("Inbox", "Archive", "Projects") → Include.

**`SearchResultProjectorTests`** (in-memory fake messages):

- Sort newest / oldest correct.
- Top-N respected.
- Snippet only computed for top-N (assert via a fake whose `Body` getter
  increments a counter).
- Items inside ExcludeAlways folders are dropped.
- COMException on one item doesn't kill the batch.

**`LiveOutlookSurface.SearchMessages` integration** (using fakes for
runner, scope builder, classifier, projector, marshaller):

- Primary success → returns projected results, no fallback called.
- Primary timeout → fallback called → fallback results returned.
- Primary COMException → fallback called.
- User CT cancellation during primary → primary stops, fallback NOT called,
  returns empty.
- User CT cancellation during fallback → loop breaks, returns accumulated.

## Configuration

For the first ship, hardcode:

- AdvancedSearch enabled: `true`.
- Timeout: `30s`.

Optional follow-up: expose `<SearchAdvancedSearchEnabled>` and
`<SearchAdvancedSearchTimeoutSeconds>` in `config.xml` for emergency
disable / tuning. Not required for this phase.

## Migration / compatibility

- `IOutlookSurface.SearchMessages` adds optional `CancellationToken` param.
  All existing callers compile unchanged.
- `OutlookSearchMessagesTool` and `OutlookCountMessagesTool` already receive
  `CancellationToken` from the dispatcher and now pass it through.
- Old conversation history's tool arguments still parse identically.
- No on-disk format changes; no config changes for the first ship.
- Trace log gains new entries; old log readers still work.

## Risk register

| Risk | Mitigation |
|---|---|
| AdvancedSearch can be slow on PSTs without an index. | Timeout + automatic fallback. |
| Some store types reject AdvancedSearch (delegate / shared mailbox edge cases). | Fallback path. |
| Event handler lifetime bugs (leaked handlers, duplicate dispatch). | Singleton runner with explicit `IDisposable`; tag-keyed dispatch; subscribe once, unsubscribe on dispose. |
| Outlook can swap `SynchronizationContext` instances. | Reuse existing `OutlookThreadMarshaller` which already handles re-capture. |
| AdvancedSearch may run concurrently with other COM calls from our add-in. | Serialise our AdvancedSearch invocations with a semaphore; other COM calls are short and run on UI thread already. |
| Results enumeration on UI thread still costs UI-thread time. | Project lazily, defer `mi.Body`, cap by `MaxResults` — typical projection cost is sub-100ms. |
| User cancels mid-search; partial results misleading. | Cancellation returns empty for primary; fallback returns whatever it had finalised. We do not surface "partial" — we return what we believe is correct or nothing. |

## Acceptance criteria

The phase is done when ALL of:

1. Running `What is my oldest email?` does NOT freeze Outlook for any
   observable duration. The Outlook UI stays responsive (mouse hover,
   ribbon click, folder switch) throughout the search.
2. Same query returns the actual oldest email across all mail stores within
   the configured timeout, or returns a clear "search timed out" error to
   the model (not a silent hang).
3. Running `Find an email with EIN` returns matching messages and does not
   freeze Outlook.
4. The compact tool-status UI shows a Stop affordance; clicking it cancels
   the in-flight search within ~1s; the tool returns a structured cancel
   result to the model.
5. All existing tests still pass; new tests above all pass.
6. Trace log shows `primary=AdvancedSearch complete` for normal searches
   and `fallback=Iterative` only in failure scenarios.

## Out of scope (explicit deferrals)

- Search ranking / relevance scoring.
- Streaming partial results to the model mid-search.
- A local mailbox index.
- Configurable timeout / disable in `config.xml`.
- Multi-mailbox / shared-mailbox-aware scope rules beyond what
  AdvancedSearch supports natively.

## Files touched (anticipated)

- New:
  - `VSTO2/OutlookAI/Services/Tools/OutlookAdvancedSearchRunner.cs`
  - `VSTO2/OutlookAI/Services/Tools/SearchScopeBuilder.cs`
  - `VSTO2/OutlookAI/Services/Tools/FolderClassifier.cs`
  - `VSTO2/OutlookAI/Services/Tools/SearchResultProjector.cs`
  - Matching test files under `VSTO2/OutlookAI.Tests/Services/Tools/`.
- Modified:
  - `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs`
    (signature: add `CancellationToken`).
  - `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs`
    (re-implement `SearchMessages` and `CountMessages`).
  - `VSTO2/OutlookAI/Services/Tools/OutlookSearchMessagesTool.cs`
    (plumb CT).
  - `VSTO2/OutlookAI/Services/Tools/OutlookCountMessagesTool.cs`
    (plumb CT).
  - `VSTO2/OutlookAI/Services/Tools/SearchExecutionHelper.cs`
    (move shared helpers, expand skip list).
  - `VSTO2/OutlookAI/ThisAddIn.cs` / DI wiring
    (construct the runner, dispose on shutdown).
  - Chat UI: compact status Stop affordance (small JS/CSS change in the
    existing `inbox-copilot` WebView2 bundle).
- Tests updated:
  - `SearchExecutionHelperTests.cs` — broaden skip-list coverage with real
    folder names.
  - `OutlookSearchMessagesToolTests.cs` — plumb CT.
  - `OutlookCountMessagesToolTests.cs` — plumb CT.

## Open questions (resolved here, callable out for review)

- **Concurrency.** Resolved: serialise via semaphore. One AdvancedSearch
  in-flight per process at a time. Keeps event dispatch trivial. If we
  ever need parallelism we can lift this later.
- **Stop button in this phase or next.** Recommend in this phase. Tiny
  amount of WebView2 work, big UX value, and we already have the
  cancellation chain to wire it to.
- **Streaming results.** Out of scope.
- **Default timeout.** 30s. Empirically larger than indexed AdvancedSearch
  ever takes; well under "user gave up" patience for an unindexed PST.
