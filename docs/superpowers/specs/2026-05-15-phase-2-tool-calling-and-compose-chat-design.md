# Phase 2 Design: Tool Calling + Compose-Window Chat

**Status:** Approved design, pre-implementation
**Branch:** `feature/codex-oauth-migration`
**Date:** 2026-05-15
**Depends on:** Phase 1 (`2026-05-14-codex-oauth-migration-design.md`)
**Followed by:** Phase 3 (Inbox Copilot UI on the Explorer ribbon — separate spec)

---

## Executive Summary

Phase 2 turns OutlookAI from a single-shot text rewriter into a tool-using assistant scoped to the compose window. It adds:

1. A **generic multi-round tool-calling loop** to `CodexChatService` (matching Codex CLI's client-side accumulation pattern, no `previous_response_id`).
2. A **10-tool Outlook catalog** — 6 read, 4 safe writes — implemented behind an `IToolHost` seam so the chat service has no compile-time COM dependency.
3. A **tabbed compose task pane** — Actions / Chat / Variants — preserving the existing Actions UX while adding two new surfaces.
4. A **WebView2 chat surface** with streaming output, per-tool inline cards, action bar per assistant message, and a small JS↔C# message bridge.
5. A **Variants tab** producing 1–5 alternative drafts with tone tags and per-card actions.
6. A **first real test project** at `VSTO2\OutlookAI.Tests\` covering services, tools, dispatcher, conversation store, and config — the testing groundwork Phase 1 deferred.

The compose-window experience remains the default entry point in Phase 2. Phase 3 will extend this same tool catalog and chat infrastructure to an Explorer-ribbon entry point so the user can chat with the whole mailbox without an open compose window.

The user-stated build directive: **thorough, thoughtful, robust, no cut corners.** Every design choice in this document was selected against that bar.

---

## Confirmed Decisions

| Item | Decision |
|---|---|
| Phase 2 surface | Compose-window task pane only; Explorer ribbon is Phase 3. |
| Mailbox read access | Full read (folders, search, read, count, recipient-thread history). |
| Mailbox write access | Safe writes only: `create_draft`, `mark_as_read`, `flag_message`, `set_category`. No send / delete / move-to-deleted. |
| Confirmation gate on writes | None per call; every write is logged to a visible audit row in chat. |
| Compose pane layout | Tabs at the top: **Actions / Chat / Variants**. Existing Actions UX preserved. |
| Chat conversation state | Per-`Inspector`, in-memory, cleared on close. Manual Clear + Copy-to-clipboard. |
| Message rendering | **WebView2** (Microsoft Edge Chromium-based). HTML/CSS/JS chat shell. Markdown via vendored `marked.min.js`. |
| Streaming | Token-by-token (`response.output_text.delta`); rendered progressively in the chat bubble. |
| Tool-call transparency | Inline tool cards: name, abbreviated args, status (running/ok/error), expandable JSON detail. |
| Variants count | Configurable 1–5, default 3. Card grid with tone tag + rationale + per-card Insert / Replace / Regenerate / Discard. |
| Tone enum | Fixed: `Formal`, `Brief`, `Persuasive`, `Friendly`, `Technical`, `Apologetic`, `Direct`, `Diplomatic`, `Enthusiastic`. |
| Stop button | Cancel HTTP, finalize partial assistant message with a `· stopped` badge, drop pending tool calls. |
| Round cap | 16 model ↔ tool round-trips per turn; `MaxRoundsReached` surfaces as an inline error card. |
| Conversation transport | Client-side history accumulation. `store: false`. No `previous_response_id`. |
| Reasoning effort | Admin-set default (Phase 1.x roadmap) + per-turn override dropdown in Chat and Variants tabs. |
| Voice input | Existing realtime WebSocket flow reused in Variants intent box. Chat-tab mic is Phase 3. |
| Test project | New `VSTO2\OutlookAI.Tests\` SDK-style net472 xUnit project, registered in a new `VSTO2\OutlookAI.sln`. |

---

## 1. Architecture Overview

### Component map

```
ThisAddIn (per-process singletons)
  ├── CodexAuthService           (Phase 1 — unchanged)
  ├── CodexChatService           (extended: RunTurnAsync + multi-round loop + streaming)
  ├── RealtimeVoiceService       (Phase 1 — unchanged)
  └── OutlookToolHost            (NEW — implements IToolHost via Outlook OOM, STA-marshaled)

AITaskPane (per-Inspector)        (rebuilt around a TabControl)
  ├── Tab "Actions"               (existing 6 buttons + Draft section + Result panel)
  ├── Tab "Chat"                  (WebView2 + JS↔C# bridge + ConversationController)
  └── Tab "Variants"              (card grid + intent input + mic + VariantController)

Each AITaskPane owns (per-Inspector state):
  ├── ConversationStore           (List<ResponseItem>, cleared on close)
  ├── VariantStore                (current variant set, cleared on regenerate / close)
  └── ToolDispatcher              (delegates to OutlookToolHost; carries owning Inspector ref)
```

### Cross-cutting rules

- The chat service never imports `Microsoft.Office.Interop.Outlook`. Outlook access happens through `IToolHost`.
- The tool host marshals every COM access onto the Outlook STA UI thread via `OutlookThreadMarshaller`. Phase 2 tools never touch COM from a thread-pool thread.
- `ToolDispatcher` carries a reference to its owning `Inspector`. The model's `outlook_get_current_compose_state` reads from that specific Inspector, **not** `Application.ActiveInspector()` — protecting against the user clicking into a different window mid-turn.
- Conversation history is a list of Codex-shape `ResponseItem`s (`Message`, `FunctionCall`, `FunctionCallOutput`). The list is the canonical state — UI rendering is a projection of it.
- Reasoning-effort override at turn time is a property on `ConversationContext`. Setting it to `null` defers to `Config.ReasoningEffort` (admin-set).

### What is NOT in Phase 2

- Explorer-window task pane / new Explorer ribbon button (Phase 3).
- Cross-Inspector chat continuity.
- Disk persistence of chat history.
- `outlook_send_message`, `outlook_delete_message`, `outlook_move_to_deleted` tools.
- File-attachment content extraction (metadata only).
- Multi-language UI (English only).
- File drag-and-drop into chat; image/screenshot paste.

---

## 2. Tool Catalog

All tools are namespaced `outlook_*`. Each tool definition lives in a single static schema file and is used both as the model-facing `tools[]` entry (in the Responses request) and as the runtime arg-validation schema before dispatch.

### Read tools (6)

#### `outlook_get_current_compose_state`

Read the compose-window state for the Inspector that owns this chat.

- Parameters: `{ include_full_body?: bool = false }`
- Returns:
  ```json
  {
    "subject": "string",
    "recipients": { "to": ["…"], "cc": ["…"], "bcc": ["…"] },
    "sender_name": "string",
    "sender_email": "string",
    "body_plaintext": "string",
    "body_truncated": false,
    "in_reply_to": {
      "thread_topic": "string",
      "last_n_messages": [
        { "from": "…", "received_at": "ISO-8601", "snippet": "string" }
      ]
    },
    "attachments": [{ "filename": "string", "size_bytes": 0 }]
  }
  ```
- Notes: when `include_full_body=false`, `body_plaintext` is the first 1000 chars and `body_truncated=true`. Saves prompt tokens for turns that only need summary context. `in_reply_to` is omitted if the compose is not a reply.

#### `outlook_list_folders`

Return the user's store folder tree.

- Parameters: `()`
- Returns: `[{ "id": "string", "name": "string", "parent_id": "string|null", "item_count": 0 }]`
- Hard caps: max depth 6, max 200 nodes. Shallow Inbox/Sent/Drafts/etc. always returned; deeper branches truncated with a `{"truncated": true}` marker.

#### `outlook_search_messages`

Server-side mailbox search via DASL `Items.Restrict` (faster than `.Find`).

- Parameters: `{ query: string, folder_id?: string, date_from?: ISO-8601, date_to?: ISO-8601, max_results?: int }`
- Returns: `[{ "id": "string", "subject": "string", "from": "string", "to": ["string"], "received_at": "ISO-8601", "snippet": "string", "has_attachments": bool }]`
- `max_results` default 25, hard cap 100.
- `query` semantics: simple full-text against subject + body. Special tokens `from:`, `to:`, `has:attachment` translated to DASL clauses.

#### `outlook_read_message`

Fetch one message by ID.

- Parameters: `{ message_id: string, include_full_body?: bool = true }`
- Returns: `{ id, subject, from, to[], cc[], received_at, body_plaintext, body_truncated: bool, attachments: [{ filename, size_bytes }], in_reply_to?: { thread_topic, message_id }, conversation_topic }`
- Body always plaintext (`MailItem.Body`; Outlook converts HTML→plain). Truncated at 32 KB with `body_truncated=true`.

#### `outlook_count_messages`

Counts without returning bodies. Same `query` syntax as `search_messages`. Much cheaper for "how many emails matching X" intents.

- Parameters: `{ query: string, folder_id?: string, date_from?: ISO-8601, date_to?: ISO-8601 }`
- Returns: `{ count: int }`

#### `outlook_list_recent_threads_with`

Group Inbox+Sent conversations by `ConversationID` for one recipient.

- Parameters: `{ recipient_email: string, max_threads?: int }`
- Returns: `[{ "thread_topic": "string", "last_message_at": "ISO-8601", "message_count": int, "snippet": "string", "thread_id": "string" }]`
- `max_threads` default 5, hard cap 20.

### Write tools (4)

#### `outlook_create_draft`

Creates a new draft in the Drafts folder. Never sends.

- Parameters: `{ subject: string, body_plaintext: string, to?: string[], cc?: string[], in_reply_to_message_id?: string }`
- Returns: `{ draft_id: string, location: "Drafts" }`
- If `in_reply_to_message_id` is set, uses `MailItem.Reply()` to seed quoted body + recipient + subject, then overwrites Subject/Body with the parameters.

#### `outlook_mark_as_read`

Toggle the `UnRead` flag on a message.

- Parameters: `{ message_id: string, read: bool }`
- Returns: `{ ok: true }`

#### `outlook_flag_message`

Set follow-up flag status.

- Parameters: `{ message_id: string, flag: "none"|"todo"|"complete" }`
- Returns: `{ ok: true }`

#### `outlook_set_category`

Replace `Categories` (single value; matches Outlook UI behavior).

- Parameters: `{ message_id: string, category: string }`
- Returns: `{ ok: true }`

### Catalog-wide rules

1. **All tools return JSON.** Errors are structured `{ "error": { "code": "string", "message": "string" } }` and become `function_call_output` items in history so the model can recover.
2. **IDs are short opaque strings**, produced by `IdResolver` from Outlook `EntryID`s. Stable within a session, not across sessions. Forged or unknown IDs return `{ "error": { "code": "not_found" } }`.
3. **Hard caps** prevent context-window blowups: `search_messages` ≤ 100 hits, `read_message` body ≤ 32 KB, `list_folders` ≤ 200 nodes, `list_recent_threads_with` ≤ 20.
4. **Write audit row** — every successful write tool call posts a permanent audit row in the Chat history: `✓ Marked 'Re: Q4 review' as read` etc.
5. **Permission posture** — admin SettingsForm gains a checklist of the 4 write tools (all checked by default). Unchecking removes that tool from the catalog server-side at chat time. Read tools are not toggleable.
6. **One class per tool**: `OutlookToolHost` is a registry mapping name → `IOutlookTool` implementation. Adding a tool in Phase 3 is just adding a class.

---

## 3. Chat Service + Multi-Round Tool Loop

### New surface on `CodexChatService`

Phase 1's `ProcessEmailAsync(action, content, prompt)` remains in place. Phase 2 adds:

```csharp
public async Task<TurnResult> RunTurnAsync(
    ConversationContext context,
    string userMessage,
    IToolHost toolHost,
    ChatEventSink sink,
    CancellationToken cancellationToken);
```

- `ConversationContext` — system instructions, accumulated `List<ResponseItem>` history, reasoning-effort override (nullable; falls back to `Config.ReasoningEffort`), tool catalog filter (which tools are enabled), the owning `Inspector` reference for scope.
- `ChatEventSink` — streaming callback interface: `OnTokenDelta(string)`, `OnToolCallStart(callId, name, argsJson)`, `OnToolCallResult(callId, ok, summary, resultJson)`, `OnAssistantMessageComplete(text)`, `OnError(error)`, `OnRoundBoundary()`.
- `TurnResult` — `{ StopReason: Completed|Cancelled|MaxRoundsReached|Error, AppendedItems: ResponseItem[], FinalAssistantText: string, RoundsUsed: int }`.

### The loop

```
function RunTurnAsync(ctx, userMsg, toolHost, sink, ct):
    ctx.History.Append(Message{role:user, content:userMsg})
    rounds = 0
    while rounds < MAX_ROUNDS (16):
        rounds++
        request  = BuildResponsesRequest(
                       model: Config.Model,
                       instructions: ctx.SystemInstructions,
                       input: ctx.History,
                       tools: tool catalog filtered by ctx.EnabledTools,
                       tool_choice: "auto",
                       parallel_tool_calls: true,
                       reasoning: effectiveReasoning(ctx),
                       store: false,
                       stream: true)
        stream    = await codexBackend.PostResponses(request, ct)
        assistantText = ""
        pendingCalls  = []
        foreach event in stream:
            switch event.type:
                "response.output_text.delta":
                    assistantText += event.delta
                    sink.OnTokenDelta(event.delta)
                "response.output_item.added" (function_call):
                    pendingCalls.Add(event.item)
                    sink.OnToolCallStart(event.item.id, event.item.name, event.item.arguments)
                "response.output_item.added" (message):
                    /* placeholder; final text already captured via deltas */
                "response.completed":
                    break event loop
                "error":
                    sink.OnError(event); throw
        if assistantText.Length > 0:
            ctx.History.Append(Message{role:assistant, content:assistantText})
            sink.OnAssistantMessageComplete(assistantText)
        if pendingCalls.IsEmpty:
            sink.OnRoundBoundary()
            return TurnResult(Completed, …)
        // Run all pending function calls in parallel (each marshals to STA internally).
        results = await Task.WhenAll(pendingCalls.Select(call =>
            DispatchOne(toolHost, call, sink, ct)))
        foreach (call, result) in (pendingCalls zip results):
            ctx.History.Append(FunctionCall{call_id, name, arguments})
            ctx.History.Append(FunctionCallOutput{call_id, output:result.json})
        if ct.IsCancellationRequested:
            return TurnResult(Cancelled, …)
        sink.OnRoundBoundary()
    return TurnResult(MaxRoundsReached, …)
```

`DispatchOne` catches all exceptions from the tool host and maps them to a structured `{ error }` JSON output so the model can react instead of the turn dying.

### Critical behaviors

- **Streaming + tool calls coexist.** SSE can interleave `output_text.delta` and `function_call` items. Both are collected; deltas render immediately, tool calls run after `response.completed`.
- **Parallel tool dispatch.** Multiple function calls in one round run via `Task.WhenAll`. Each call still serializes onto the Outlook STA thread internally via `OutlookThreadMarshaller`.
- **Bounded rounds.** 16 round cap. Surfaces `MaxRoundsReached` as an inline error card: *"Stopped after 16 tool-calling rounds. Try refining the question."*
- **Cancellation.** `CancellationToken` propagated through HTTP and tool dispatch. Stop = HTTP request aborted, pending tool calls cancelled, partial assistant message preserved with `· stopped` badge, history updated, future turns can continue.
- **Tool errors → history.** Failed dispatch becomes a `function_call_output` with `{ "error": { code, message } }`. The model sees the failure and can adapt.
- **Client-side history only.** `store: false` on every request. No `previous_response_id`. Matches Codex CLI's canonical pattern.

### Actions tab integration

Existing buttons (Proofread / Revise / etc.) now route through `RunTurnAsync` with:

- Each action's existing system prompt + a small addendum: *"You may call mailbox tools if you need additional context. Most quick edits don't require any tools."*
- Empty starting `History`.
- Same `IToolHost` instance.
- Same `ChatEventSink` (but bound to the Result panel — tool cards show as a compact strip above the result text).
- A Cancel button appears next to "Processing…" while a turn is active.

User experience is unchanged for trivial actions (no tool calls fire). For Draft Reply, the model now reliably reads the actual thread via `outlook_get_current_compose_state` instead of receiving 4000 chars of pre-truncated text.

---

## 4. WebView2 Chat Surface

### Embedding model

- Dependency: `Microsoft.Web.WebView2` (`Microsoft.Web.WebView2.WinForms` host control), `net472` target.
- Runtime: Microsoft Edge WebView2 Evergreen Runtime (preinstalled on Win 10 21H2+, Win 11, Server 2022, Server 2025). Install script detects via `HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}\pv`; if missing, runs the Evergreen Bootstrap installer (`MicrosoftEdgeWebView2Setup.exe`) shipped alongside `OutlookAI.vsto`.
- Files served from a virtual host: `SetVirtualHostNameToFolderMapping("outlookai.local", appDir, CoreWebView2HostResourceAccessKind.Allow)`. Initial URL: `https://outlookai.local/index.html`. No on-disk extraction, no temp files, no AV false positives, no DevTools in production builds.

### Files

```
VSTO2\OutlookAI\WebUI\
  ├── index.html
  ├── styles.css
  ├── chat.js
  ├── marked.min.js          (vendored, MIT)
  └── highlight.min.js       (vendored, MIT — code blocks)
```

All five files compile as embedded resources and are extracted at startup to `%LOCALAPPDATA%\OutlookAI\WebUI\` (idempotent; hash-verified). The WebView host points at that directory via the virtual host mapping.

### JS ↔ C# bridge

Bidirectional via `CoreWebView2.PostWebMessageAsJson` (JS→C#) and `ExecuteScriptAsync` (C#→JS). All payloads JSON.

**C# → JS** (rendering):

- `chat.appendUserMessage({ id, text, ts })`
- `chat.appendAssistantMessage({ id, ts })` — empty bubble; fills via deltas
- `chat.appendTextDelta({ id, delta })`
- `chat.appendToolCallCard({ id, name, args_brief })`
- `chat.updateToolCallCard({ id, status: "running"|"ok"|"error", summary?, result_json? })`
- `chat.finalizeAssistantMessage({ id, text, stopped: bool })`
- `chat.showError({ message })`
- `chat.setComposerEnabled(bool)`
- `chat.clear()`
- `chat.applyTheme({ accent: "#…", dark: bool })`
- `chat.setContextStrip({ subject, sender, thread_message_count })`

**JS → C#** (user intent):

- `send({ text, reasoning_override?: "minimal"|"low"|"medium"|"high" })`
- `stop()`
- `clear()`
- `insertMessage({ id })`
- `replaceMessage({ id })`
- `copyMessage({ id })`
- `regenerate({ id })`
- `expandToolCard({ id })`

### Visual treatment

- **User bubbles** right-aligned, accent fill (matches Outlook's User Defined Color Scheme accent read via `Office.ColorSchemeBeginColor`).
- **Assistant bubbles** left-aligned, neutral fill, hover-revealed action bar (Insert / Replace / Copy / Regenerate).
- **Tool cards** between bubbles: monospace pill with icon + tool name + abbreviated args + status indicator. Click expands to a JSON-styled detail block (raw args + raw result). While running, animated dots.
- **Streaming text** trailing caret while the bubble is being filled; caret hides on finalize.
- **Stopped messages** small `· stopped` badge inline.
- **Errors** red-tinted system cards.
- **System cards** appear for write-tool audit rows: `✓ Marked 'Re: Q4 review' as read`.

### Composer (input area)

- Multi-line `<textarea>`, auto-grows up to 5 lines, scrollbar after.
- Send/Stop button toggles based on turn state.
- Reasoning override dropdown (None / Minimal / Low / Medium / High) — defaults to admin setting, shows a `·` indicator when not at default.
- Slash commands: `/clear`, `/copy`. Also accessible via toolbar.
- Context strip below the composer: *"Email: 'Q4 review' from John Doe · 3 messages in thread"*. Clicking it switches to the Actions tab.

### Accessibility + integration polish

- Keyboard: arrow keys / PgUp / PgDn scroll naturally; Enter sends, Shift+Enter newline.
- High-contrast theme detected via `Microsoft.Win32.SystemEvents.UserPreferenceChanged`, bridged to JS.
- Outlook accent color read once at startup and on theme change.
- Right-click context menu in chat: Copy / Select All only — no DevTools, no "Inspect element" in production.

### Failure handling

- **WebView2 runtime missing post-install** → Chat tab shows: *"Microsoft Edge WebView2 Runtime is required. [Install]"* — clicking runs the bundled bootstrapper.
- **WebView2 process crashes** (`CoreWebView2.ProcessFailed`) → inline reload button; conversation state survives in C#.
- **Slow startup** → skeleton state for up to ~800 ms while WebView initializes.

---

## 5. Variants Tab

### Layout (260 px wide)

```
Variants
─────────────────────────────
Intent: [ multiline textbox ] [🎤]

Count: [▾ 3]   [ Generate Variants ]
Reasoning: [▾ default]
─────────────────────────────
┌──────────────────────────────┐
│ [Formal · 412 chars]         │
│ "Dear Mr. Smith, ..."        │
│ rationale: "Mirrors prior…"  │
│ [Expand] [Insert] [Replace] [↻] │
└──────────────────────────────┘
┌──────────────────────────────┐
│ [Brief · 188 chars]          │
│ "Hi John, ..."               │
│ rationale: "Saves time"      │
│ [Expand] [Insert] [Replace] [↻] │
└──────────────────────────────┘
┌──────────────────────────────┐
│ [Persuasive · 624 chars]     │
│ ...                          │
└──────────────────────────────┘

[ Regenerate all ]
```

### Behavior

1. **Intent input** — same TextBox+mic pattern as the Phase 1 Draft prompt. Mic uses the existing `RealtimeVoiceService`.
2. **Count picker** — NumericUpDown 1–5, default 3.
3. **Generate Variants** — runs a single chat turn with:
   - System prompt instructing exactly *N* variants in a constrained JSON envelope.
   - Tool catalog available (model may call `outlook_get_current_compose_state` and `outlook_list_recent_threads_with` to inform tone).
   - System prompt explicitly forbids calling `outlook_create_draft` from variant generation — variants are previewed-only.
4. **Constrained output shape** (model must emit, single fenced JSON code block):
   ```json
   {
     "variants": [
       {
         "tone": "Formal",
         "rationale": "Mirrors recipient's prior tone",
         "subject": "string",
         "body": "string"
       }
     ]
   }
   ```
5. **Tone enum** — fixed list (see Confirmed Decisions §). Free-form model tones are post-processed to the nearest enum value.
6. **Per-card actions:**
   - **Expand** — modal preview with full body + full rationale.
   - **Insert** — adds the body at the top of the compose body (Phase 1 semantics).
   - **Replace** — overwrites the compose body.
   - **↻ Regenerate this one** — 1-variant turn with the same intent + an instruction to preserve the same tone tag but vary phrasing.
7. **Regenerate all** — discards the current set and runs the original turn fresh with a different seed.
8. **VariantStore** — per-`Inspector`, in-memory, cleared on inspector close or on next Generate Variants click.

### Visual states

- **Idle** — empty list with hint text: *"Describe the email you want. I'll draft a few options for you to compare."*
- **Generating** — skeleton cards; tone/rationale arrive first via streaming, body fills last.
- **Generated** — cards in steady state.
- **Error** — red banner + retry button.

---

## 6. Testing, Quality Bar, and Rollout

### Test project

`VSTO2\OutlookAI.Tests\` — SDK-style project targeting `net472`, registered in a new `VSTO2\OutlookAI.sln`.

Packages: `xunit`, `xunit.runner.visualstudio`, `Microsoft.NET.Test.Sdk`, `Newtonsoft.Json`, `Moq` (for `IToolHost` fakes only — production code remains POCO).

Verification:

```
nuget restore VSTO2\OutlookAI.sln
msbuild     VSTO2\OutlookAI.sln /p:Configuration=Debug /p:Platform="Any CPU"
vstest.console VSTO2\OutlookAI.Tests\bin\Debug\OutlookAI.Tests.dll
```

### Coverage matrix

| Layer | Tests |
|---|---|
| `CodexAuthService` | PKCE shape; code/refresh exchange request bodies; atomic token write; expiry detection; single-flight semaphore. `FakeHttpMessageHandler` for OAuth endpoints. (Phase 1 owed; written now.) |
| `CodexChatService` single-shot | Request body shape; SSE delta accumulation; final-text override; error mapping; reasoning-effort propagation. Fixture-driven `FakeHttpMessageHandler`. |
| `CodexChatService` multi-round | Tool calls dispatched; output appended; second-round request shape; max-round cap; cancellation mid-round preserves partial output; parallel tool dispatch. Scripted multi-response sequences + `FakeToolHost`. |
| Each `IOutlookTool` | Argument validation; result shape; error mapping for missing/invalid IDs. `FakeOutlookSurface` mocks just the OOM methods each tool uses. Tools never touch real COM in unit tests. |
| `ToolDispatcher` | Name routing; unknown-tool rejection; JSON-schema arg validation; validation errors surface as structured tool errors. |
| `ConversationStore` | Add / clear / copy-to-clipboard formatting; per-Inspector isolation. |
| `VariantStore` + variant JSON parsing | Well-formed and malformed JSON handling; tone enum clamping; count clamping; round-trip. |
| `IdResolver` | Round-trip EntryID ↔ opaque ID; rejection of forged short IDs. |
| `OutlookThreadMarshaller` | Calls execute on UI thread; cancellation honored; exceptions propagate. `SynchronizationContext` test seam. |
| `Config` v2 + v1-ignore | Per the Phase 1 plan (deferred from Phase 1). Temp-dir paths. |

### Deliberately not unit tested

- WebView2 host — manual smoke only.
- Real Outlook OOM — manual smoke checklist.
- Realtime WebSocket — Phase 1 left this manual; same here.

### Manual smoke checklist (`docs/superpowers/checklists/phase-2-smoke.md`)

1. Open Outlook → new email → AI Assistant.
2. Actions tab works (Proofread / Revise / Draft Reply).
3. Chat tab: "Summarize this thread" → tool cards: `outlook_get_current_compose_state` → `outlook_read_message` → final summary.
4. Chat tab: "Has John replied to my proposal?" → `outlook_list_recent_threads_with` + `outlook_search_messages`.
5. Chat tab: "Mark all newsletter emails as read" → repeated `outlook_search_messages` + `outlook_mark_as_read`; mailbox state matches.
6. Chat tab: Stop mid-stream → partial message preserved with `· stopped`; follow-up works.
7. Variants tab: count 3, "Polite follow-up about Q4 report" → 3 distinct-tone cards.
8. Variants tab: regenerate one card → only that card changes.
9. Settings: change model to `gpt-5.5-pro` + reasoning to High → next turn reflects them.
10. Close and reopen compose window → chat history cleared (by design); Model/Reasoning persist.

### Performance + robustness gates

| Gate | Target |
|---|---|
| Cold WebView2 init (first chat tab activation) | < 800 ms on warm boxes |
| First streamed token latency (Chat) | < 2.0 s p50, < 4.0 s p95 |
| Streaming render throughput | ≥ 30 tokens/sec without frame drops |
| `outlook_search_messages(max=25)` on 50k-message Inbox | < 1.5 s |
| `outlook_read_message` on 200 KB body | < 200 ms |
| Multi-round turn (3 tool calls + final text) | < 8 s p50 end-to-end |
| Memory delta after 100 chat turns | < 50 MB sustained growth (GC settles) |
| Main-thread block duration during a tool call | < 100 ms |

### Rollout

1. All Phase 2 commits land on `feature/codex-oauth-migration`, no merge to master.
2. Continuous push to GitHub on that branch.
3. Manual smoke pass on dev machine.
4. Canary install on one RDS server with 2–3 trusted users for 1-week dogfood.
5. Defect triage + iterate.
6. Go/no-go decision for Phase 3 based on Phase 2 stability.

### Risk register

| Risk | Mitigation |
|---|---|
| WebView2 runtime missing on a target | Install script bootstraps via bundled Evergreen installer. Fallback panel offers manual install. |
| Outlook OOM throws `COMException` intermittently on RDS | Each tool catches `COMException`, returns structured `{ error }`. Retries once with backoff for transient codes (e.g., `RPC_E_CALL_REJECTED`). |
| Model hallucinates an EntryID | `IdResolver` rejects forged IDs; tool returns `not_found`; model can recover. |
| Model emits malformed Variants JSON | Variant parser falls back to plain text + error card; user can regenerate. |
| `outlook_search_messages` slow on huge mailboxes | Hard `max_results` cap; `Items.Restrict` (DASL) over `.Find`; STA-marshaled background work. |
| Chat backend latency spikes | UI shows `· thinking` after 3 s, `· still working` after 10 s. Cancellation always honored. |
| Multi-window concurrent turns saturate backend | Each pane has its own context; `CodexAuthService` already single-flights token access. |
| WebView2 process crash mid-turn | `CoreWebView2.ProcessFailed` handler shows inline reload; conversation state survives in C#. |
| Prompt injection via inbound mail content | All tool outputs containing untrusted text (snippets, bodies) are clearly fenced in the model's view (`<<UNTRUSTED MAILBOX CONTENT>>` markers around each retrieved body). System prompt instructs the model never to follow instructions embedded in mailbox content. |
| `outlook_create_draft` populating Drafts unexpectedly | Variants explicitly forbid the tool; Chat allows it but every call posts a visible audit row. |

---

## 7. Implementation Boundaries

These are the units the writing-plans phase will turn into bite-sized tasks. Each is independently testable.

| Unit | New / Modified | Purpose |
|---|---|---|
| `Services\IToolHost.cs` | New | Interface between chat service and Outlook access. |
| `Services\Tools\IOutlookTool.cs` | New | Per-tool implementation contract. |
| `Services\Tools\OutlookToolRegistry.cs` | New | Maps tool name → `IOutlookTool` impl. |
| `Services\Tools\OutlookGetCurrentComposeStateTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookListFoldersTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookSearchMessagesTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookReadMessageTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookCountMessagesTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookListRecentThreadsWithTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookCreateDraftTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookMarkAsReadTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookFlagMessageTool.cs` | New | Tool implementation. |
| `Services\Tools\OutlookSetCategoryTool.cs` | New | Tool implementation. |
| `Services\OutlookToolHost.cs` | New | Aggregates tools, owns `OutlookThreadMarshaller`, implements `IToolHost`. |
| `Services\OutlookThreadMarshaller.cs` | New | Marshals all COM access onto the Outlook STA UI thread. |
| `Services\IdResolver.cs` | New | EntryID ↔ short opaque ID. |
| `Services\Chat\ConversationContext.cs` | New | Per-pane chat context (history + system + reasoning override). |
| `Services\Chat\ConversationStore.cs` | New | Per-pane chat state. |
| `Services\Chat\ChatEventSink.cs` | New | Streaming callback interface. |
| `Services\Chat\TurnResult.cs` | New | DTO. |
| `Services\Variants\VariantStore.cs` | New | Per-pane variant set. |
| `Services\Variants\VariantParser.cs` | New | Parses Variants JSON envelope. |
| `Services\CodexChatService.cs` | Modified | Adds `RunTurnAsync`, multi-round loop, tool-call event handling, parallel dispatch. |
| `Services\Tools\ToolCatalogSchema.cs` | New | Static JSON schema definitions (one schema per tool). |
| `TaskPane\AITaskPane.cs` + `AITaskPane.Designer.cs` | Modified | Replaces single-column layout with `TabControl`. Actions tab inherits current controls. Chat tab embeds `WebView2`. Variants tab is a card list. |
| `TaskPane\Chat\ChatController.cs` | New | Mediates between `CodexChatService` events and the WebView2 surface. |
| `TaskPane\Chat\WebView2Bootstrap.cs` | New | Creates the WebView2, sets virtual host mapping, registers the JS message handler. |
| `TaskPane\Variants\VariantsController.cs` | New | Drives Variants tab UX. |
| `WebUI\index.html` + `styles.css` + `chat.js` + `marked.min.js` + `highlight.min.js` | New | Chat surface assets (embedded resources). |
| `SettingsForm.cs` | Modified | Adds the tool-permission checklist + the Phase 1.x model + reasoning effort dropdowns. |
| `Config.cs` | Modified | Adds `ReasoningEffort` and `WriteToolsEnabled` flags. |
| `OutlookAI.csproj` | Modified | Adds `Microsoft.Web.WebView2`; embeds WebUI files. |
| `packages.config` | Modified | Adds WebView2. |
| `Deploy\Install-OutlookAI.ps1` | Modified | WebView2 Evergreen detection + bootstrap; ships `MicrosoftEdgeWebView2Setup.exe`. |
| `Deploy\MicrosoftEdgeWebView2Setup.exe` | New (vendored) | Microsoft-redistributable Evergreen bootstrap installer. |
| `VSTO2\OutlookAI.sln` | New | Solution containing product + test project. |
| `VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj` | New | xUnit test project. |
| `VSTO2\OutlookAI.Tests\Helpers\FakeHttpMessageHandler.cs` | New | Already designed in Phase 1's deferred test scaffold. |
| `VSTO2\OutlookAI.Tests\Helpers\FakeToolHost.cs` | New | Scripted tool dispatch for chat-service tests. |
| `VSTO2\OutlookAI.Tests\Helpers\FakeOutlookSurface.cs` | New | Per-method mockable Outlook OOM surface. |
| Per-tool test class (10×) | New | One test class per `IOutlookTool` implementation. |
| `VSTO2\OutlookAI.Tests\Services\CodexAuthServiceTests.cs` | New | Phase 1 owed. |
| `VSTO2\OutlookAI.Tests\Services\CodexChatServiceTests.cs` | New | Single-shot + multi-round. |
| `VSTO2\OutlookAI.Tests\Services\ToolDispatcherTests.cs` | New | Routing + validation. |
| `VSTO2\OutlookAI.Tests\Services\ConversationStoreTests.cs` | New | State + isolation. |
| `VSTO2\OutlookAI.Tests\Services\VariantParserTests.cs` | New | Envelope parsing. |
| `VSTO2\OutlookAI.Tests\Services\IdResolverTests.cs` | New | Round-trip + forgery. |
| `VSTO2\OutlookAI.Tests\Services\OutlookThreadMarshallerTests.cs` | New | Threading. |
| `VSTO2\OutlookAI.Tests\ConfigTests.cs` | New | Phase 1 owed. |

---

## 8. Open Follow-ups (Phase 2.x)

These extend the spec after the initial Phase 2 cut lands and the dogfood feedback is in:

- **Custom tone enum.** Letting admins extend the Variants tone list via Settings.
- **Per-turn token-budget display.** Surfacing ChatGPT-plan usage hints.
- **Conversation export to .md file.** Beyond clipboard copy.
- **Pre-built prompt templates.** Slash commands like `/summarize-thread`, `/find-action-items`, `/draft-reply-cordial`.
- **Tool call retries.** Specific transient COM errors (e.g., `RPC_E_CALL_REJECTED` during user typing) auto-retry once with backoff before surfacing.

These are intentionally not in the initial Phase 2 cut. Anything broader belongs in Phase 3 (Inbox Copilot UI on the Explorer ribbon).
