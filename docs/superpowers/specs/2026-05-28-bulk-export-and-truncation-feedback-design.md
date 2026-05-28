# Bulk Export + Search Truncation Feedback — Design

**Date:** 2026-05-28
**Target release:** v2.1.2
**Status:** Approved (design phase)

## Problem

`outlook_search_messages` silently truncates large result sets to `max_results`
(hard cap 100, default 25) and returns **no signal** to the model that more
matches existed. An attorney asked for what amounted to a complete list of
matching emails; the model issued one search, received ≤100 rows with no
indication of truncation, built an export from those rows, and reported done —
silently omitting the rest. The user described this as "still hitting the cap."

### Root cause (verified by tracing the code)

1. `SearchMessagesArgsParser.ParseSearch` hard-caps `max_results` at 100,
   default 25 (`SearchMessagesArgsParser.cs:12-14`).
2. For a `from:`/`subject:`/`body:` filter over `all_mail`/`auto`, the search
   runs through `FallbackIterativeSearch`, which collects **every** match into
   `allInputs` (no early-stop fires for those filters —
   `SearchFallbackBudget.ShouldStop*` only trigger for `to:` or no-filter broad
   scans).
3. `SearchResultProjector.Project` then **silently clamps** `allInputs` down to
   `max_results` and discards the true count (`SearchResultProjector.cs:33-34`).
4. `OutlookSearchMessagesTool` returns `{ messages: [...] }` with **no
   `total_matches` and no `truncated` flag** (`OutlookSearchMessagesTool.cs:26-36`).

The model is therefore blind to truncation. The existing schema guidance to
"page by date window" (`ToolCatalogSchema.cs:56-57`) cannot help, because the
model never learns that a given search *was* truncated. The schema also
actively discourages large pulls ("use 25 / never 100" —
`ToolCatalogSchema.cs:53,100`), which makes truncation worse for targeted
lookups.

## Key insight that shapes the design

The 100-result cap is **only** in the model-facing parser
(`SearchMessagesArgsParser.ParseSearch`). It is **not** a limit of the
underlying surface. `FallbackIterativeSearch` already collects all matches
(up to its per-folder budget) before projection. Therefore a **server-side
tool can request up to a configurable ceiling directly — no date-window
walking required.** "Chunk by date window" was the *model's* workaround for the
parser cap; server-side code is not bound by it.

## Goals

1. Tell the model when a search was truncated (feedback loop) so it can
   self-correct for synthesis-style tasks.
2. Provide a deterministic, server-side path that produces a **complete**
   mechanical list export (up to a bounded ceiling) regardless of model
   behavior.
3. Stay within the project's hard-won performance guardrails — never reintroduce
   the multi-minute Outlook freeze that the `SearchFallbackBudget` machinery
   exists to prevent.

## Non-goals

- AI-synthesized bulk exports ("summarize each of 400 emails"). That inherently
  requires the model to read and reason per message; no single server-side tool
  call can synthesize thousands of bodies. Covered by Mechanism 1 + existing
  tools, not by the new tool.
- Full-message-body columns in the bulk tool. The bulk tool projects only fields
  available from the fast Table-API search path (incl. the ~200-char snippet).
  Body-derived columns remain the synthesis case.
- PDF output from the bulk tool. Excel only. A 2,000-row mechanical table is not
  a useful PDF.
- Server-side date-window walking. Unnecessary given the key insight above.

## Architecture

Two complementary mechanisms.

### Mechanism 1 — truncation feedback on `outlook_search_messages`

The surface contract changes so the search result carries the pre-clamp count.

- New value type `SearchResult`:
  ```
  SearchResult {
      IReadOnlyList<MessageSummary> Messages;   // clamped to max_results
      int TotalMatches;                          // pre-clamp count
      bool Truncated;                            // TotalMatches > Messages.Count
  }
  ```
  `TotalMatches` is the accurate pre-clamp count where the collection path saw
  all matches (the common `from:`/`subject:`/`body:` case). Where an early-stop
  fired (`to:` or broad no-filter scans), the collection is itself a floor; in
  that case `TotalMatches` is reported as the floor and `Truncated` is `true`
  (honest "at least this many, more exist"). The distinction is recorded in the
  trace log; the model-facing contract is simply "truncated means you did not
  get everything."

- `SearchResultProjector.Project` returns `SearchResult` (count + truncated +
  clamped list) instead of just the list.

- `IOutlookSurface.SearchMessages` returns `SearchResult`.
  `LiveOutlookSurface.SearchMessages` populates the count from
  `allInputs.Count` (fallback path) or `primary.Items.Count` (AdvancedSearch
  path) — both already computed and currently only logged.

- `OutlookSearchMessagesTool` adds `total_matches` and `truncated` to its JSON.

- `ToolCatalogSchema` steering for `outlook_search_messages` is rewritten:
  - Explain that `truncated: true` means the result is incomplete.
  - On truncation: escalate to `outlook_export_search_results` for a complete
    list export, OR (for synthesis tasks) narrow by date window and accumulate
    across turns.
  - Remove the contradictory "use 25 / never 100" phrasing; replace with
    guidance to size with `outlook_count_messages` when completeness matters.

### Mechanism 2 — `outlook_export_search_results` (deterministic backstop)

A new model-callable tool. Entirely server-side, no per-message body reads.

Flow:
1. Parse a search filter (same filter fields as `outlook_search_messages`) plus
   a column selection from a fixed allowed set.
2. `CountMessages(filter)` → **M** (true total; count-mode is already bounded
   per folder by `SearchFallbackBudget.CountModeCap`).
3. `SearchMessages(filter, max_results = min(ceiling, M))` → up to **N** rows via
   the fast Table-API fallback path.
4. Project the chosen mechanical columns and build a typed `.xlsx` via the
   existing `ExcelWorkbookBuilder`.
5. Return `{ file, exported: N, total_matches: M, truncated: (N < M) }`.

Allowed columns (mechanical, from search projection): `subject`, `from`, `to`,
`received_at`, `snippet`, `has_attachments`, `folder`. Unknown column names are
dropped (consistent with existing `Config.EnabledWriteTools` intersection
pattern). At least one valid column required.

Ceiling: `Config.MaxBulkExportRows`, default **2,000**, loaded from the
server/global `config.xml` (layer-2 global path; not user-overridable, like
`CodexAuthPath`/`VoiceModel`). Clamped to a sane floor/cap on load.

Reuses existing export infrastructure: `ExportPathResolver` (inherits the
v2.1.1 UNC→LocalAppData fallback), `ExportFilenameSanitizer`,
`ExcelWorkbookBuilder`, `FileSavedResult`, and the `IExportPathPolicy` gate.
Results surface as the same file card UI as the existing export tools.

## Components touched

| Component | Change |
|---|---|
| `Services/Tools/SearchResult.cs` | **New.** Value type: Messages + TotalMatches + Truncated. |
| `Services/Tools/SearchResultProjector.cs` | Return `SearchResult` (count + truncated + clamped list). |
| `Services/Tools/IOutlookSurface.cs` | `SearchMessages` returns `SearchResult`. |
| `Services/Tools/LiveOutlookSurface.cs` | Populate count/truncated from existing raw counts. |
| `Services/Tools/OutlookSearchMessagesTool.cs` | Emit `total_matches` + `truncated`. |
| `Services/Tools/ExportSearchResultsArgs.cs` + parser | **New.** Filter + column selection. |
| `Services/Tools/OutlookExportSearchResultsTool.cs` | **New.** The bulk tool. |
| `Services/Tools/ToolCatalogSchema.cs` | Register new tool; rewrite truncation/chunking steering. |
| `Services/OutlookToolHost.cs` | Register the new tool. |
| `Config.cs` | New `MaxBulkExportRows` (default 2000), global-config load + clamp. |
| `IOutlookSurface` consumers | Any other caller of `SearchMessages` updated for the new return type. |

## Data flow (bulk tool)

```
model -> outlook_export_search_results { filter…, columns:[…], filename_hint }
  -> ExportSearchResultsArgsParser (validate filter + columns)
  -> surface.CountMessages(filter)            => M
  -> surface.SearchMessages(filter, min(ceiling, M))  => SearchResult (N rows)
  -> project rows -> ExcelWorkbookBuilder -> ExportPathResolver/Policy -> save
  -> { file:{path,name}, exported:N, total_matches:M, truncated:(N<M) }
```

## Error / partial handling

- Zero matches → still produce an (empty-but-valid) workbook? No — return a
  structured "no matches" result without writing a file, so the model can tell
  the user plainly. (Matches existing zero-result conventions.)
- Over ceiling → export the ceiling's worth, set `truncated:true`,
  `total_matches:M`; the model relays "exported 2,000 of M; narrow the date
  range for the rest."
- Path unavailable / IO / sharing violation → reuse the existing
  `ExportException` envelope and file-card error path.
- Cancellation → `OperationCanceledException` → structured `cancelled` error,
  consistent with the other tools.

## Testing

- `SearchResultProjectorTests`: accurate `TotalMatches`/`Truncated` for
  clamped, exactly-at-limit, and under-limit inputs.
- `ExportSearchResultsArgsParserTests`: filter parsing reuse, column validation,
  unknown-column drop, empty-columns rejection.
- `OutlookExportSearchResultsToolTests`: fake `IOutlookSurface` returning 0 /
  under-ceiling / over-ceiling sets → asserts row count, `exported`,
  `total_matches`, `truncated`, ceiling honored, filename/path via fakes.
- `OutlookSearchMessagesToolTests`: result JSON includes `total_matches` +
  `truncated`.
- `ConfigTests`: `MaxBulkExportRows` default, global override, clamp.
- `ToolCatalogSchema` snapshot/shape test updated for the new tool + tool count.

## Rollout

- Single focused plan; subagent-driven TDD with spec-compliance + code-quality
  review per task, matching the workflow used for the in-app updater and v2.1.1.
- Ships as v2.1.2. Release build is local (CI workflow still blocked by issue
  #9); publish via `Make-ReleaseZip.ps1` + `gh release create`, same as v2.1.0/1.
- Smoke: on the RDS (or dev box), ask for a known-large list export
  ("Excel of every email from <frequent sender>") and confirm the row count
  matches `outlook_count_messages` (up to the 2,000 ceiling) and the
  partial-result message appears when over ceiling.
