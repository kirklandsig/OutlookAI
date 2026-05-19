# Phase 4: Inbox Reports — Design

Branch: `feature/codex-oauth-migration`
Author session: 2026-05-18
Status: Proposed (pending user review)

## Why

After Phase 3b made mailbox search responsive, the user asked: "could I have a generated report based on email content?" The chat is already capable of producing one-off summaries when asked, but six report scenarios came up repeatedly:

1. Daily / weekly digest of new mail
2. Conversation summary with a specific person
3. Action items extracted from messages
4. Topic / project status
5. Statistical breakdown (sender counts, busiest days, folder breakdown)
6. Out-of-office catchup

Each is recurring enough to deserve a one-click affordance instead of "the user has to remember how to phrase the prompt."

## Goal

Introduce a dedicated **Inbox Reports** task pane that mirrors the InboxCopilot architecture, surfaces these six report scenarios as quick-action chips, and ships two new bulk-friendly tools so the model can produce reports without N×round-trips.

## Non-goals

- Persistent report history (save to disk + history list).
- Scheduled auto-runs (daily digest at 9am etc.).
- Email-to-self with the report content.
- Raw CSV / JSON / Excel export of email data.
- Cross-pane state sharing (Reports and Copilot remain independent).
- Rich UI parameter forms in place of placeholder text in chip templates.

These are explicit follow-up phases. Phase 4 ships an ephemeral, templated, chat-driven reports pane and the supporting tools.

## Architecture

Reports is a **parallel** task pane to InboxCopilot. Both can be open at once. Each owns its own conversation, tool host, and prompt. No shared state.

```
ribbon button "Reports"
   ↓
ThisAddIn.ShowReportsTaskPane(Explorer)
   ↓
new InboxReportsPane                  (parallel to InboxCopilotPane)
   ├ WebView2 surface                  (reused)
   ├ InboxReportsController            (parallel to InboxCopilotController)
   │    ├ ConversationStore            (per-pane instance)
   │    ├ CodexChatService             (reused, same singleton)
   │    └ OutlookToolHost              (reused, with extended tool catalog)
   ├ LiveOutlookSurface                (reused, with two new methods)
   ├ InboxReportsPromptBuilder         (parallel to InboxCopilotPromptBuilder)
   └ ReportQuickActionChips            (six templates)
```

**Threading:** unchanged from Phase 3b. Every Outlook OOM call goes through `OutlookThreadMarshaller`. The Table-API path with `YieldUi(ct)` between stores / folders / items keeps Outlook responsive.

**State lifecycle:** each pane instance has its own `ConversationStore`. Closing the pane drops the conversation. Copy-to-clipboard uses the existing `ExportForClipboard()`.

## Components — new

### `InboxReportsPane` (`VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.cs`)

Mirrors `InboxCopilotPane`:

- Constructed by `ThisAddIn.ShowReportsTaskPane(Explorer)`.
- `Bind(Outlook.Explorer explorer)` resolves marshaller, IdResolver, application, AdvancedSearchRunner, FolderClassifier from `Globals.ThisAddIn`.
- Constructs `LiveOutlookSurface`, `OutlookToolHost` with the extended tool catalog, `InboxReportsController`.
- Hosts the WebView2 control. Same WebView2 bootstrap as InboxCopilot.

### `InboxReportsController` (`VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs`)

Mirrors `InboxCopilotController`. Differences:

- Builds prompt via `InboxReportsPromptBuilder`, not `InboxCopilotPromptBuilder`.
- Initializes the WebView with the chips JS payload (six chip definitions, not the InboxCopilot chips).
- Otherwise identical: `WebMessageReceived` handler, `_activeCts`, streaming sink, conversation store, copy/clear handling.

### `InboxReportsPromptBuilder` (`VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs`)

Builds the system prompt for the Reports pane. Returns plain text. Asserts in tests:

- Mentions `outlook_read_messages` (bulk) and recommends it over many single reads.
- Mentions `outlook_aggregate_messages` and recommends it for counts/grouping.
- Mentions `outlook_search_messages` and recommends date_from/date_to use.
- Instructs markdown structure, header-with-scope-and-count, concise tone.
- Instructs: if a `[placeholder]` is still in the user's text, ask for clarification first.

### `ReportQuickActionChip` (`VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs`)

A small immutable POCO: `{ string Label, string TemplateText, string IconUnicode }`. A static `Defaults()` method returns the six chips below in order.

Templates:

| # | Label | Template prefilled into input |
|---|---|---|
| 1 | `📅 This week's digest` | `Summarize what came into my Inbox over the past 7 days. Group by sender or topic. Highlight urgent items and emails I'm directly addressed in.` |
| 2 | `💬 Conversation summary` | `Summarize my recent email conversations with [name or email]. Show the chronological flow and key decisions/topics.` |
| 3 | `✓ Action items` | `Find action items I need to do based on emails from the past 7 days. Read the relevant messages, extract TODOs/deadlines/asks. Group by who's waiting on what.` |
| 4 | `📁 Project status` | `Summarize the status of [topic/project name]. Find relevant emails, read the most recent ones, and give me: latest update, open questions, action items, key participants.` |
| 5 | `📊 Email stats` | `Give me email statistics for the past 30 days: top 10 senders, busiest days, breakdown by folder. Use outlook_aggregate_messages.` |
| 6 | `🏖️ While I was out` | `I was out from [start date] to [end date]. Show me what's important from that timeframe: urgent items, direct asks, replies needed. De-prioritize newsletters and automated mail.` |

Clicking a chip **prefills** the chat input with the template text and **does not auto-submit**, so the user can edit `[placeholders]` before sending. Identical UX to InboxCopilot's existing chips.

### Ribbon button

Add a new ribbon button "Reports" in `Ribbon.xml` (sibling to "AI Assistant"). The click handler calls `ThisAddIn.ShowReportsTaskPane(Application.ActiveExplorer())`. The pane toggles visibility on subsequent clicks, same pattern as `ShowExplorerTaskPane`.

## Components — new tools

### `outlook_read_messages` — bulk read

```jsonc
// Args
{
  "ids": ["abc123", "def456"],
  "include_body": true,                 // default true
  "max_items": 25                       // default 25, hard cap 100
}

// Returns
{
  "messages": [
    {
      "id": "abc123",
      "subject": "...",
      "from": "Jane Doe",
      "to": ["bob@example.com"],
      "cc": [],
      "received_at": "2026-05-14T18:32:00Z",
      "body_plaintext": "...",          // only when include_body=true
      "body_truncated": false,
      "attachments": [{"filename":"x.pdf","size_bytes":1234}],
      "in_reply_to_message_id": null,
      "conversation_topic": "Q4 plan"
    }
  ]
}
```

**Surface contract:**
```csharp
IReadOnlyList<MessageDetail> ReadMessages(
    string[] ids,
    bool includeBody,
    int maxItems,
    CancellationToken ct = default);
```

**Implementation:** for each short ID, call `_ids.Resolve(...)` to get the EntryID, then `_application.Session.GetItemFromID(entryId)`. Build a `MessageDetail` matching the existing `ReadMessage` shape. Marshalled to UI thread. `YieldUi(ct)` between items. Unknown / unreadable IDs are dropped silently and traced. `max_items` is clamped to `[1, 100]` and acts as a hard cap on the array length passed to the surface.

**Tool class:** `OutlookReadMessagesTool : IOutlookTool` in `VSTO2/OutlookAI/Services/Tools/`, named `"outlook_read_messages"`. Parses args, calls surface, projects JSON. Same `OperationCanceledException → cancel envelope` pattern as Phase 3b.

### `outlook_aggregate_messages` — group & count

```jsonc
// Args
{
  "scope": "all_mail",                  // current_folder | all_mail | auto
  "folder_id": "",                      // optional explicit folder
  "date_from": "2026-05-01T00:00:00Z",
  "date_to":   "2026-05-31T23:59:59Z",
  "from": "",                           // optional substring
  "subject_contains": "",
  "body_contains": "",
  "group_by": "sender",                 // sender | day | folder
  "top_n": 10                           // default 10, hard cap 100
}

// Returns
{
  "buckets": [
    {"label": "Jane Doe", "count": 47},
    {"label": "Bob Smith", "count": 31}
  ],
  "total": 312,                         // pre-cap sum across all matching messages
  "scanned_folders": 35
}
```

**Surface contract:**
```csharp
IReadOnlyList<AggregationBucket> AggregateMessages(
    AggregateMessagesArgs args,
    CancellationToken ct = default);
```

Plus a result-shaping pair:
```csharp
public sealed class AggregationBucket { public string Label { get; set; } public int Count { get; set; } }
public sealed class AggregateMessagesArgs { /* fields above */ }
```

**Implementation:** use the same Table API path we shipped for `SearchMessages`. Per resolved folder, call `folder.GetTable(filter, olUserItems)` with columns `[SenderName, SenderEmailAddress, ReceivedTime, MessageClass]`. `TableMessageClassFilter` drops non-mail rows. Bucket key derives from `group_by`:

- `sender` → `SenderName` if non-empty, else `SenderEmailAddress`, normalized (`Trim()`).
- `day`    → `ReceivedTime.Date.ToString("yyyy-MM-dd")`.
- `folder` → folder display name.

After all folders, group in C# (`GroupBy(key).Select(g => new Bucket(g.Key, g.Count()))`), sort descending by `count`, take `top_n`. `total` is the sum across **all** buckets (not just top_n) and represents matched messages. Folder enumeration uses the existing `ResolveSearchFolders` with `YieldUi(ct)` between folders.

**Tool class:** `OutlookAggregateMessagesTool : IOutlookTool` in `VSTO2/OutlookAI/Services/Tools/`, named `"outlook_aggregate_messages"`.

### Tool catalog registration

Both tools are added to `ToolCatalogSchema` (and tested there) so the model sees them in the catalog. Schemas include `description` text steering the model toward when to prefer each — examples:

- `outlook_read_messages`: "Use this instead of multiple outlook_read_message calls when you have an array of IDs from a prior search and you need bodies. Faster by 5-10×."
- `outlook_aggregate_messages`: "Use this for counts and groupings ('top 10 senders this month', 'busiest days last week', 'breakdown by folder'). Avoid running outlook_count_messages many times."

## Data flow (action items report, illustrative)

```
User clicks ✓ Action items chip
   ↓ chip's template prefilled into input
User clicks Send
   ↓
InboxReportsController.StartTurnAsync(text)
   ├ ConversationStore.Append(userMessage)
   ├ InboxReportsPromptBuilder.Build(scope) → system prompt
   └ CodexChatService.RunTurnAsync(prompt, toolCatalog, CT)
        ↓ model emits:
        ├ outlook_search_messages {date_from: now-7d, date_to: now, scope: auto}
        │     → LiveOutlookSurface.SearchMessages(args, CT)   [Phase 3b path]
        │     → returns N message summaries
        ├ outlook_read_messages {ids:[...], include_body:true}
        │     → LiveOutlookSurface.ReadMessages(ids, true, 25, CT)
        │     → returns N message details with bodies
        └ assistant message: markdown action-items report
             → streamed to WebView via WebViewSink
             → rendered in chat scrollback
```

## Threading & cancellation

Unchanged from Phase 3b. Each tool call marshals to the Outlook UI thread, calls `YieldUi(ct)` between items / folders, and respects the chat session's CT. The `Stop` button on the composer cancels the active turn via `_activeCts.Cancel()`, which cascades through dispatcher → tool → surface → `ct.ThrowIfCancellationRequested()` → tool emits cancel envelope.

The `OutlookAdvancedSearchRunner` semaphore (one in-flight `AdvancedSearch` per process) is preserved; we don't add new `AdvancedSearch` invocations.

## Error handling

| Failure | Behavior |
|---|---|
| Tool `COMException` | Trace + return safe default (empty list, null detail). Model receives the empty result and can respond accordingly. |
| Tool `OperationCanceledException` | Tool emits `{"error":{"code":"cancelled","message":"Cancelled by user."}}`; chat shows cancelled state. |
| Surface receives empty `ids[]` | Return empty list, no COM calls. |
| Unknown short ID in `ReadMessages` | Drop silently, trace it. Continue with remaining IDs. |
| Per-folder Table API failure in `AggregateMessages` | Skip folder, trace, continue. Partial result returned with `scanned_folders` count reflecting only successful folders. |
| Model produces malformed markdown | Rendered as plain text by WebView. No crash. |
| Two panes open, both querying Outlook | Existing `OutlookAdvancedSearchRunner` semaphore + per-call marshalling serialise correctly. Each pane has its own `ConversationStore`. |
| Placeholder `[name]` etc. left in user's prompt | Model is instructed to ask for clarification before tool calls. |

## Testing strategy

All tests are TDD'd before implementation, in the order listed.

| Layer | Tests |
|---|---|
| **Helpers (pure)** | `SenderKeyNormalizerTests` (sender name vs email priority, trim, null/empty). `DateBucketFormatterTests` (UTC date format `yyyy-MM-dd`, time-of-day stripped). `TopNBucketSelectorTests` (sort by count desc, take N, ties stable). `AggregateMessagesArgsParserTests` (defaults, validation, group_by enum). |
| **`OutlookReadMessagesTool`** | `Execute_PassesIdsThrough`, `Execute_ClampsMaxItemsToHardCap`, `Execute_ProjectsMessageDetails`, `Execute_PassesCancellationTokenThroughToSurface`, `Execute_OnCancellation_EmitsStructuredCancelEnvelope`, `Execute_EmptyIds_ReturnsEmptyArray`. |
| **`OutlookAggregateMessagesTool`** | `Execute_PassesAllArgsThroughToSurface`, `Execute_ClampsTopNToHardCap`, `Execute_ProjectsBucketsAndTotal`, `Execute_PassesCancellationTokenThroughToSurface`, `Execute_OnCancellation_EmitsStructuredCancelEnvelope`, `Execute_DefaultsMatchSpec` (group_by required, top_n=10, etc.). |
| **`InboxReportsPromptBuilder`** | `Prompt_AlwaysIncludesRolePreamble`, `Prompt_MentionsBulkReadTool`, `Prompt_MentionsAggregateTool`, `Prompt_TellsModelToAskBeforeUnresolvedPlaceholders`, `Prompt_DescribesMarkdownAndConciseFormat`. |
| **`ReportQuickActionChip`** | `Defaults_ReturnsSixChips`, `Defaults_EachChipHasLabelAndTemplate`, `Defaults_TemplatesMentionPlaceholdersWhereAppropriate` (chips 2/4/6). |
| **`ToolCatalogSchema`** | `OutlookReadMessages_Schema_HasIdsArrayAndBodyToggle`, `OutlookAggregateMessages_Schema_HasGroupByEnumAndTopN`, `OutlookAggregateMessages_Description_HintsAtUseInsteadOfManyCountCalls`. |
| **`MinimalSurface` + tests' nested `Surface : MinimalSurface`** | Add overridable `ReadMessages` and `AggregateMessages` with `OnReadMessages` / `OnAggregateMessages` delegates following the existing pattern. |
| **Integration / smoke** | Manual on the real mailbox: each of 6 chips, fill placeholders, send, verify report renders. Trace inspected for expected tool-call sequence. |

No tests require live Outlook OOM — the surface methods themselves are smoke-verified only, same as `LiveAdvancedSearchHost` in Phase 3b.

## Implementation phasing

Although Phase 4 ships as one design, it is committed in slices for safer rollout. Each slice is independently buildable, testable, and (where useful) deployable.

| Slice | What ships | Verification gate |
|---|---|---|
| 1 | Ribbon button + `ThisAddIn.ShowReportsTaskPane` + empty `InboxReportsPane` skeleton + `InboxReportsController` wiring + DI | Pane opens, toggles, no crash. Existing tests pass. |
| 2 | `InboxReportsPromptBuilder` + `ReportQuickActionChip.Defaults()` + chip rendering in pane | Click each chip, see template prefilled. Send works, model responds using existing tools (action items will be slow without bulk-read). |
| 3 | `IOutlookSurface.ReadMessages` + `LiveOutlookSurface.ReadMessages` (Table-API-friendly via GetItemFromID) + `OutlookReadMessagesTool` + schema | Action items / topic status reports become noticeably faster (bulk read replaces N×read_message). |
| 4 | `IOutlookSurface.AggregateMessages` + `LiveOutlookSurface.AggregateMessages` + `OutlookAggregateMessagesTool` + schema | Stats / OOO reports use grouping in one tool call. |
| 5 | Full real-Outlook smoke for all 6 reports + Release publish + elevated install + push to origin | All chips produce reasonable reports on the real mailbox. Trace shows expected tool-call patterns. |

Total: roughly **6-12 commits**, similar shape to Phase 3b.

## Files touched (anticipated)

**New:**
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.cs`
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPane.Designer.cs`
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsController.cs`
- `VSTO2/OutlookAI/TaskPane/InboxReports/InboxReportsPromptBuilder.cs`
- `VSTO2/OutlookAI/TaskPane/InboxReports/ReportQuickActionChip.cs`
- `VSTO2/OutlookAI/Services/Tools/OutlookReadMessagesTool.cs`
- `VSTO2/OutlookAI/Services/Tools/OutlookAggregateMessagesTool.cs`
- `VSTO2/OutlookAI/Services/Tools/AggregateMessagesArgsParser.cs`
- `VSTO2/OutlookAI/Services/Tools/SenderKeyNormalizer.cs`
- `VSTO2/OutlookAI/Services/Tools/DateBucketFormatter.cs`
- `VSTO2/OutlookAI/Services/Tools/TopNBucketSelector.cs`
- Matching test files under `VSTO2/OutlookAI.Tests/`.

**Modified:**
- `VSTO2/OutlookAI/Services/Tools/IOutlookSurface.cs` — add `ReadMessages` + `AggregateMessages` signatures, `AggregateMessagesArgs`, `AggregationBucket`.
- `VSTO2/OutlookAI/Services/Tools/LiveOutlookSurface.cs` — implement the two new methods.
- `VSTO2/OutlookAI/Services/Tools/ToolCatalogSchema.cs` — register the two new tools' schemas.
- `VSTO2/OutlookAI/Services/OutlookToolHost.cs` — register the two new tools in the catalog.
- `VSTO2/OutlookAI/ThisAddIn.cs` — add `ShowReportsTaskPane(Explorer)`.
- `VSTO2/OutlookAI/Ribbon.xml` — add Reports button.
- `VSTO2/OutlookAI/Ribbon.cs` — wire the new button's `OnAction` to `ThisAddIn.ShowReportsTaskPane`.
- `VSTO2/OutlookAI.Tests/Services/Tools/MinimalSurface.cs` — add overridable `ReadMessages` and `AggregateMessages`.

## Risks

| Risk | Mitigation |
|---|---|
| Reports pane + Copilot pane both active, both calling tools | Marshaller + AdvancedSearchRunner semaphore from Phase 3b already serialise correctly. |
| Bulk `ReadMessages` may still be slow on a huge `ids[]` (e.g., 100 messages) | `max_items` hard cap (100) + `YieldUi` between items keeps the UI responsive. The model is steered to keep `ids[]` reasonably sized via prompt + schema description. |
| Model overuses `aggregate_messages` when `count_messages` is cheaper | Prompt + schema description disambiguate. Acceptable cost if it occasionally misroutes. |
| Two `ConversationStore` instances in memory | Tiny; each is a list of conversation items. No issue. |
| Webview2 second instance is heavy | We already run one per pane (Inspector + Explorer can coexist with InboxCopilot). Adding a third instance for Reports is the same shape. |

## Acceptance criteria

Phase 4 is complete when ALL of:

1. Clicking the ribbon "Reports" button opens the Inbox Reports pane (and toggles it off on re-click).
2. The pane shows all six chips. Clicking each prefills the chat input with the matching template, does not auto-submit, and the user can edit `[placeholders]` before sending.
3. Sending any of the six prompts produces a markdown-formatted report rendered in the pane's chat scrollback within a reasonable time (under ~30s for digest/stats on this user's 200-folder mailbox).
4. The model uses `outlook_read_messages` for action items and topic-status reports (verifiable in the trace).
5. The model uses `outlook_aggregate_messages` for the stats report (verifiable in the trace).
6. Outlook UI remains responsive throughout (per Phase 3b acceptance — DoEvents yielding, marshalled chunks).
7. Cancellation via Stop button works within ~1s and the model receives a cancel envelope.
8. All new tests pass; all existing tests still pass.
9. Each slice in the implementation plan is independently committable and verifiable.

## Open questions (resolved here, called out for review)

- **Should the chip row scroll if more chips are added in the future?** Resolved: not yet. Six fits in a 2×3 grid in 340px-wide pane. Revisit if we add more.
- **Should report results include any raw structured data (JSON beside the markdown) for future export?** Resolved: not yet. Phase 4 is markdown-only. Raw-data export is the explicit follow-up phase.
- **Should the pane share the InboxCopilot's `ConversationStore`?** Resolved: no, separate per pane. Reports and Copilot are different mental models; mixing their conversation histories would be confusing.
- **Should the ribbon button be in a new "AI Assistant" group with the existing button, or its own group?** Resolved: same group, sibling button. Logical pairing.
