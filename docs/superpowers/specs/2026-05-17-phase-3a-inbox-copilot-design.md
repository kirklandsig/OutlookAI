# Phase 3a Design: Inbox Copilot UI

**Status:** Approved design, pre-implementation
**Branch:** `feature/codex-oauth-migration`
**Date:** 2026-05-17
**Depends on:** Phase 2 (`2026-05-15-phase-2-tool-calling-and-compose-chat-design.md`)
**Followed by:** Phase 3b (Calendar tools), 3c (Contacts), 3d (Tasks), 3e (Send — separate spec each)

---

## Executive Summary

Phase 3a brings the Outlook AI Assistant out of the compose window and onto the main Explorer view. The user clicks the same "AI Assistant" ribbon button from the Inbox, gets a per-Explorer chat pane, and can ask mailbox-wide questions — "what needs my attention?", "summarize unread", "find emails from Jane about Q4" — using the existing 10-tool catalog (plus one small new tool for the current selection) and the existing Phase 2 chat infrastructure.

The pane is **chat-only** (no tabs), with a context strip showing the user's current folder + selected message and a row of quick-action chips above the composer that auto-send common prompts. Conversations are **per-Explorer, in-memory, cleared on close** — the same lifecycle model as Phase 2's compose pane.

Phase 3a also makes `outlook_search_messages` and `outlook_count_messages` substantially smarter by exposing structured filter fields (`from`, `subject_contains`, `body_contains`, `has_attachment`, `is_unread`, `is_flagged`, `importance`) so the model can express precise mailbox queries without writing DASL syntax. This is the key unlock that makes Inbox Copilot actually useful — Phase 2's `query`-only search was naive `LIKE '%foo%'` against subject + body, which forced the model to make multiple sequential search calls to express anything specific.

No new write tools. No send. No delete. The Phase 2 safe-write set (`create_draft`, `mark_as_read`, `flag_message`, `set_category`) is unchanged. Future high-risk tools (`outlook_send_message`) will need their own design pass with a per-call confirmation gate (Phase 3e).

---

## Confirmed Decisions

| Item | Decision |
|---|---|
| Pane shape | Chat-only. No tabs. Full-pane WebView2 surface reusing the Phase 2 chat WebUI bundle. |
| Quick-action chips | Row above the composer. Auto-send on click. Static set + dynamic chips that depend on selection state. |
| Ribbon entry | Unified — same "AI Assistant" group on both `TabNewMailMessage` (compose, Phase 2) and `TabMail` (Explorer, new). One button, contextually routed. |
| Click routing | `Application.ActiveWindow` tells us Inspector vs Explorer. Inspector → existing `ShowTaskPane`. Explorer → new `ShowExplorerTaskPane`. |
| Current-state awareness | Folder name + unread count + currently-selected message(s) baked into every turn's system prompt. New `outlook_get_current_selection` tool lets the model drill in further. |
| Conversation lifecycle | Per-Explorer, in-memory. Cleared when the Explorer closes. Manual Clear + Copy-to-clipboard preserved from Phase 2. |
| Multi-Explorer | Each open Explorer window gets its own pane + independent conversation. Same pattern as multi-Inspector today. |
| Voice | Deferred to Phase 3.x. Chat-tab mic was already Phase 3 work in the Phase 2 spec; treating it as a separate follow-up. |
| Send / delete tools | Still not present. Phase 3e (separate spec) will add `outlook_send_message` with a per-call confirmation dialog. |
| Search enhancement | `outlook_search_messages` + `outlook_count_messages` gain structured fields. Backward-compatible (old `query`-only calls still work). |
| Body-text search performance | Stays on DASL `LIKE`. `AdvancedSearch` (Windows Search index) is a Phase 3.x perf follow-up if real mailboxes get slow. |
| Quick-action chip config | Hardcoded for v1. Admin-configurable chips via `Config` XML is a Phase 3.x follow-up. |

---

## 1. Architecture Overview

### Component map (additions only — everything from Phase 2 stays)

```
ThisAddIn (per-process singletons)
  ├── existing Phase 1 + Phase 2 services unchanged
  └── (no new singletons in Phase 3a)

CustomTaskPanes
  ├── per-Inspector  AITaskPane         (Phase 2 — unchanged)
  └── per-Explorer   InboxCopilotPane   (NEW)

InboxCopilotPane (per-Explorer)
  ├── WebView2 chat surface             (reuses Phase 2 chat.js / styles.css / index.html)
  └── InboxCopilotController            (NEW — parallel to ChatController)

Each InboxCopilotPane owns (per-Explorer state):
  ├── ConversationStore                 (List<ResponseItem>, cleared on Explorer close)
  ├── LiveOutlookSurface                (constructed with composeInspector=null)
  └── OutlookToolHost                   (registers all 10 + new selection tool)

Tools (count: 11)
  ├── outlook_get_current_compose_state   (Phase 2 — unchanged; returns "no compose" when no Inspector)
  ├── outlook_get_current_selection       (NEW — Phase 3a)
  ├── outlook_list_folders                (Phase 2 — unchanged)
  ├── outlook_search_messages             (ENHANCED — structured fields)
  ├── outlook_read_message                (Phase 2 — unchanged)
  ├── outlook_count_messages              (ENHANCED — same args as search)
  ├── outlook_list_recent_threads_with    (Phase 2 — unchanged)
  ├── outlook_create_draft                (Phase 2 — unchanged)
  ├── outlook_mark_as_read                (Phase 2 — unchanged)
  ├── outlook_flag_message                (Phase 2 — unchanged)
  └── outlook_set_category                (Phase 2 — unchanged)
```

### Cross-cutting rules

- `Application.ActiveWindow` is read at ribbon-click time only; per-pane code does not re-query it. Each `InboxCopilotController` holds a reference to its owning `Explorer` and operates strictly against that reference.
- Same STA-marshaling discipline as Phase 2: every Outlook OOM call goes through `OutlookThreadMarshaller`. The new selection tool is no exception.
- `LiveOutlookSurface` already accepts a nullable `composeInspector` in its constructor (Phase 2 design). Phase 3a passes `null` for that parameter when constructing the surface for an Explorer-bound pane — `GetCurrentComposeState` will return `EmptyCompose()` in that case, which is the right answer.
- The Inbox Copilot pane carries its own `ConversationStore`. Compose-pane conversations and Inbox-pane conversations never mix — even if the user has both open simultaneously.

---

## 2. Ribbon + Click Routing

### `Ribbon.xml` changes

Current (Phase 2):

```xml
<tab idMso="TabNewMailMessage">
  <group id="AIAssistantGroup" label="AI Assistant" insertAfterMso="GroupClipboard">
    <button id="btnAIAssistant" ... onAction="OnAIAssistantClick" .../>
  </group>
</tab>
```

Phase 3a adds an identical group on `TabMail` (Outlook's Home tab in the Explorer):

```xml
<tab idMso="TabMail">
  <group id="AIAssistantExplorerGroup" label="AI Assistant" insertAfterMso="GroupMailMove">
    <button id="btnAIAssistantExplorer"
            label="AI Assistant"
            size="large"
            onAction="OnAIAssistantClick"
            supertip="Open the AI Assistant to chat with your mailbox..."
            imageMso="SmartArtChangeColorsGallery" />
  </group>
</tab>
```

Note: `id`s must be distinct (`btnAIAssistant` vs `btnAIAssistantExplorer`), but both buttons call the same `onAction="OnAIAssistantClick"`. Outlook ignores the group ID for navigation purposes; the *button ID* uniqueness is what matters.

### `Ribbon.cs` callback (unchanged)

`OnAIAssistantClick` already calls `Globals.ThisAddIn.ShowTaskPane()`. We change `ShowTaskPane` to context-route:

```csharp
public void ShowTaskPane()
{
    var activeWindow = this.Application.ActiveWindow();
    if (activeWindow is Outlook.Inspector insp)
    {
        ShowInspectorTaskPane(insp);   // existing Phase 2 path, renamed for clarity
    }
    else if (activeWindow is Outlook.Explorer expl)
    {
        ShowExplorerTaskPane(expl);    // NEW
    }
    else
    {
        // Defensive: shouldn't happen since the ribbon button only renders
        // when an Inspector or Explorer is the active window.
        ShowGenericNoActiveWindowMessage();
    }
}
```

`ShowInspectorTaskPane` is the existing per-Inspector logic from Phase 2. `ShowExplorerTaskPane` is the new path:

```csharp
private void ShowExplorerTaskPane(Outlook.Explorer explorer)
{
    foreach (CustomTaskPane pane in this.CustomTaskPanes)
    {
        if (pane.Window == explorer)
        {
            pane.Visible = !pane.Visible;
            return;
        }
    }
    var paneControl = new InboxCopilotPane();
    paneControl.Bind(explorer);
    var ctp = this.CustomTaskPanes.Add(paneControl, "AI Assistant", explorer);
    ctp.Width = 340;
    ctp.Visible = true;
}
```

### Explorer lifecycle

`ThisAddIn_Startup` hooks `Application.Explorers.NewExplorer` so we can attach `Close` handlers per Explorer. Same pattern Phase 2 uses for Inspectors, just on the other collection. On `Explorer.Close`, dispose the matching pane's controller (which disposes the WebView2 and clears the ConversationStore).

Multiple simultaneous Explorers (a common Outlook pattern — second window for a different folder) each get their own pane and conversation. No shared state.

---

## 3. InboxCopilotPane + InboxCopilotController

### `InboxCopilotPane` (WinForms UserControl)

Mirrors `AITaskPane` in shape but simpler — no TabControl. The control is a single root container holding the WebView2 surface.

```
InboxCopilotPane (UserControl)
  ├── lblTitle (small header, optional — could omit)
  ├── btnSettings (gear icon, opens existing SettingsForm)
  └── WebView2 (Dock = Fill)  ← managed by InboxCopilotController
```

`Bind(Outlook.Explorer)` is the entry point called by `ShowExplorerTaskPane`:

1. Store the Explorer reference.
2. Construct `LiveOutlookSurface(app, marshaller, idResolver, composeInspector: null, explorer: explorer)` — note the new constructor parameter (see § 4).
3. Construct `OutlookToolHost(surface, Config.WriteToolsEnabled)`.
4. Construct `InboxCopilotController(tabHostContainer, ChatService, toolHost, surface, conversationStore, explorer)`.
5. Fire `controller.InitializeAsync()` fire-and-forget (with `ContinueWith` to surface faulted-task exceptions to TraceLog, same as `ChatController`).

`DisposeCustomResources` disposes the controller, which in turn disposes the WebView2 and cancels any in-flight turn.

### `InboxCopilotController`

Parallel to `ChatController` from Phase 2, with these differences:

| Concern | ChatController (Phase 2) | InboxCopilotController (Phase 3a) |
|---|---|---|
| Anchor | `Outlook.Inspector` (compose) | `Outlook.Explorer` (mailbox view) |
| System prompt | Compose helper, knows the draft body | Mailbox helper, knows the folder + selection |
| Context strip | Subject, To, thread topic | Folder name + unread count + selected message |
| Quick-action chips | Not present | Static + dynamic chips rendered above composer |
| Selection awareness | N/A | Tracks `Explorer.SelectionChange` to update chips + system prompt |
| Conversation store | Per-Inspector | Per-Explorer |

Initialization sequence (when WebView is ready):

1. `outlookai.applyTheme('light')` — same as Phase 2.
2. `outlookai.setReasoningOptions(efforts, '')` — same as Phase 2.
3. `outlookai.setContextStrip(...)` — folder + selection summary (see § 5).
4. `outlookai.setQuickActions(chips)` — see § 6.
5. Subscribe to `Explorer.SelectionChange` so chips + context strip refresh when the user clicks a different message.

Turn handling (`StartTurnAsync`):

- Same pattern as `ChatController.StartTurnAsync`: append user message, set composer disabled, append empty assistant message, await `RunTurnAsync`, marshal final assistant message back to the WebUI, finalize.
- `ConversationContext.SystemInstructions` is rebuilt fresh each turn from the current `Explorer.CurrentFolder` + `Explorer.Selection` — model always sees up-to-date state.
- `ConversationContext.IncludeWriteTools` honors `Config.WriteToolsEnabled` exactly as Phase 2.
- ChatEventSink callbacks marshal through `OutlookThreadMarshaller` (existing `RunScript` plumbing — unchanged from the Phase 2 fix).

---

## 4. `outlook_get_current_selection` Tool

### Schema

```json
{
  "type": "object",
  "properties": {
    "include_full_bodies": {
      "type": "boolean",
      "description": "If true, returns full message body for each selected item (up to 32 KB). Default false: 200-char snippet only."
    },
    "max_items": {
      "type": "integer",
      "minimum": 1,
      "maximum": 20,
      "description": "Hard cap on how many selected items to return. Default 5."
    }
  },
  "additionalProperties": false
}
```

### Result shape

```json
{
  "folder": "Inbox",
  "folder_id": "AAMkAGI2...",
  "count": 1,
  "messages": [
    {
      "id": "<short_opaque_id>",
      "subject": "Re: Q4 plan",
      "from": "Jane Doe <jane@acme.com>",
      "received_at": "2026-05-17T09:14:00Z",
      "snippet": "Hi team — Thanks for the early look. A few thoughts on the regional split...",
      "has_attachments": true,
      "is_unread": false,
      "conversation_topic": "Q4 plan"
    }
  ]
}
```

When `include_full_bodies: true`, each message gets an additional `body_plaintext` field (truncated at 32 KB with `body_truncated` flag, same conventions as `outlook_read_message`).

### `IOutlookSurface` extension

```csharp
public sealed class CurrentSelectionResult
{
    public string Folder { get; set; }
    public string FolderId { get; set; }
    public int Count { get; set; }
    public IReadOnlyList<MessageSummary> Messages { get; set; }   // or MessageDetail when include_full_bodies
}

public interface IOutlookSurface
{
    // ... existing methods unchanged ...
    CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems);
}
```

`LiveOutlookSurface.GetCurrentSelection` reads from the Explorer reference passed at construction time:

```csharp
public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) =>
    Run(() =>
    {
        if (_explorer == null) return EmptySelection();   // compose-anchored pane
        var sel = _explorer.Selection;
        var folder = _explorer.CurrentFolder;
        // ... iterate sel up to maxItems, project each MailItem to MessageSummary/MessageDetail ...
    });
```

Non-`MailItem` selection entries (meeting requests, contacts, tasks) are filtered out for Phase 3a — they need different surface methods (Phase 3b/c/d).

### Tool registration

`OutlookToolHost` registers `OutlookGetCurrentSelectionTool` unconditionally (it's a read tool — not gated on `WriteToolsEnabled`). `ToolCatalogSchema` adds the schema entry above. No Settings UI change needed.

---

## 5. Context Strip

### What the UI shows

```
┌──────────────────────────────────────┐
│ In: Inbox (47 unread)                │
│ Selected: Re: Q4 plan — Jane Doe     │
└──────────────────────────────────────┘
```

Two lines:
- Folder line: `In: <folder name> (<unread count> unread)`. Always present.
- Selection line: `Selected: <subject> — <from display name>`. Present only when ≥1 message is selected. When 2+ are selected, shows `Selected: 3 messages` instead.

The third line (`context-thread` in the compose pane) is repurposed: in Inbox context it's hidden, or could show a long-form folder path for nested folders ("Inbox / Vendors / Acme").

### What the model sees in `SystemInstructions`

The system prompt is rebuilt each turn from the live Explorer state. Skeleton:

```
You are the Outlook Inbox Copilot. The user is viewing their mailbox.
Help them search, summarize, triage, and act on messages. You have
mailbox tools available; prefer one well-targeted tool call over many.

Current context:
- Folder: Inbox (47 unread, 1284 total)
- Selected: Re: Q4 plan
  From: Jane Doe <jane@acme.com>
  Received: 2026-05-17T09:14:00Z
  Snippet: Hi team — Thanks for the early look. A few thoughts...

Reply concisely; the user is busy.
```

The "Selected:" block is included only when something is selected. If nothing is selected, that part of the prompt is omitted entirely.

### Update triggers

- On WebView ready (once at pane open).
- On `Explorer.SelectionChange` event.
- On `Explorer.FolderSwitch` event.

Each update pushes a fresh `outlookai.setContextStrip(...)` + (if chips changed) `outlookai.setQuickActions(...)` to the WebUI.

---

## 6. Quick-Action Chips

### Chip data shape

A chip is an object with: visible label, the prompt to fire when clicked, and an optional condition predicate (in C# — chips are computed server-side and pushed to the WebUI as a flat list).

```csharp
public sealed class QuickActionChip
{
    public string Label { get; set; }    // shown on the button
    public string Prompt { get; set; }   // pre-filled into the textarea
}
```

### Default chip set

**Static chips** (always shown):

| Label | Prompt |
|---|---|
| What needs my attention? | "Look at my inbox and tell me what needs attention. Prioritize by recency, importance, and sender. Be concise." |
| Summarize unread | "Summarize all my unread messages. Group by sender or topic. Be concise." |
| Today's emails | "Show me everything I received today, grouped by sender. Highlight anything that looks urgent." |

**Dynamic chips** (rendered only when ≥1 message is selected in the Explorer):

| Label | Prompt |
|---|---|
| Summarize this thread | "Summarize the selected message and the rest of its conversation thread." |
| Draft a reply | "Draft a reply to the selected message. Match the tone of the sender." |

When 2+ messages are selected, the dynamic chips change to:

| Label | Prompt |
|---|---|
| Summarize all selected | "Summarize all the selected messages." |
| Triage selected | "Triage the selected messages — which need action, which can be archived, which can be marked read?" |

### Click handling

The WebUI's `outlookai.setQuickActions(chips)` API renders the chip row. On click, the JS:

1. Sets the textarea value to the chip's `prompt`.
2. Immediately triggers the existing send path (auto-send per the design).

Auto-send is the locked-in behavior; the user explicitly requested completing the task as quickly as possible.

### Refresh on selection change

`Explorer.SelectionChange` fires on the UI thread. The controller's handler:

1. Recomputes the chip list based on `Explorer.Selection.Count`.
2. Calls `outlookai.setQuickActions(newChips)` to repaint.

### Admin-configurable chips (deferred)

Phase 3.x: `Config` gains a list-of-chips field persisted to XML; Settings UI gets a chip editor (add/remove/edit). For v1 (Phase 3a), chips are hard-coded in `InboxCopilotController.BuildDefaultChips()`. Adjusting them requires a code change + redeploy.

---

## 7. Enhanced `outlook_search_messages` and `outlook_count_messages`

### Motivation

The Phase 2 schema's `query` field only supports a free-form string. The C# side builds a DASL `@SQL=` filter that does:

```
urn:schemas:httpmail:subject LIKE '%<query>%' OR urn:schemas:httpmail:textdescription LIKE '%<query>%'
```

That's it. No way for the model to express "from Jane" or "with attachment" or "unread" except by hoping the substring matches somewhere. Real-world consequence: the model makes multiple sequential search calls trying to intersect results, each one costing a round-trip and consuming the round budget. Pre-Phase-3a this is mostly hidden because compose-pane chat rarely uses search aggressively. Inbox Copilot will use it constantly.

### New schema

```json
{
  "type": "object",
  "properties": {
    "query":             { "type": "string",  "description": "Free-form text matched against subject + body. Combine with structured filters via AND. Leave empty if filtering only by structured fields." },
    "from":              { "type": "string",  "description": "Sender substring. Matches display name OR email address (case-insensitive)." },
    "subject_contains":  { "type": "string" },
    "body_contains":     { "type": "string" },
    "has_attachment":    { "type": "boolean" },
    "is_unread":         { "type": "boolean" },
    "is_flagged":        { "type": "boolean" },
    "importance":        { "type": "string", "enum": ["low","normal","high"] },
    "folder_id":         { "type": "string", "description": "Default: Inbox" },
    "date_from":         { "type": "string", "format": "date-time" },
    "date_to":           { "type": "string", "format": "date-time" },
    "max_results":       { "type": "integer", "minimum": 1, "maximum": 100, "description": "Default 25. Hard cap 100." }
  },
  "additionalProperties": false
}
```

Applies to both `outlook_search_messages` (returns hits) and `outlook_count_messages` (returns count only). Same `SearchMessagesArgs` POCO; same `BuildRestrictFilter` builder.

### `SearchMessagesArgs` POCO update

```csharp
public sealed class SearchMessagesArgs
{
    public string Query { get; set; }
    public string From { get; set; }                    // NEW
    public string SubjectContains { get; set; }         // NEW
    public string BodyContains { get; set; }            // NEW
    public bool? HasAttachment { get; set; }            // NEW
    public bool? IsUnread { get; set; }                 // NEW
    public bool? IsFlagged { get; set; }                // NEW
    public string Importance { get; set; }              // NEW — "low"|"normal"|"high"
    public string FolderId { get; set; }
    public DateTimeOffset? DateFrom { get; set; }
    public DateTimeOffset? DateTo { get; set; }
    public int MaxResults { get; set; } = 25;
}
```

### DASL filter construction

`LiveOutlookSurface.BuildRestrictFilter(args)` builds the `@SQL=` AND-of-clauses:

| Arg | DASL clause |
|---|---|
| `query` (non-empty) | `(urn:schemas:httpmail:subject LIKE '%<q>%' OR urn:schemas:httpmail:textdescription LIKE '%<q>%')` |
| `from` (non-empty) | `(urn:schemas:httpmail:fromname LIKE '%<f>%' OR urn:schemas:httpmail:fromemail LIKE '%<f>%')` |
| `subject_contains` | `urn:schemas:httpmail:subject LIKE '%<s>%'` |
| `body_contains` | `urn:schemas:httpmail:textdescription LIKE '%<b>%'` |
| `has_attachment: true` | `urn:schemas:httpmail:hasattachment = 1` |
| `has_attachment: false` | `urn:schemas:httpmail:hasattachment = 0` |
| `is_unread: true` | `urn:schemas:httpmail:read = 0` |
| `is_unread: false` | `urn:schemas:httpmail:read = 1` |
| `is_flagged: true` | `http://schemas.microsoft.com/mapi/proptag/0x10900003 = 2` (PR_FLAG_STATUS = flagged) |
| `is_flagged: false` | `http://schemas.microsoft.com/mapi/proptag/0x10900003 <> 2` |
| `importance: low` | `http://schemas.microsoft.com/mapi/proptag/0x00170003 = 0` |
| `importance: normal` | `http://schemas.microsoft.com/mapi/proptag/0x00170003 = 1` |
| `importance: high` | `http://schemas.microsoft.com/mapi/proptag/0x00170003 = 2` |
| `date_from` | `urn:schemas:httpmail:datereceived >= '<iso>'` |
| `date_to` | `urn:schemas:httpmail:datereceived <= '<iso>'` |

All non-null clauses are joined with `AND`. If zero clauses are present, the filter is `null` and we enumerate the folder unfiltered (capped by `max_results`).

`Escape()` (the existing helper) protects against quotes in user-supplied strings.

### Backward compatibility

Existing Phase 2 calls that only supply `query` keep working unchanged — the new fields default to null/unset and contribute no clauses. The two new test fixtures in `OutlookSearchMessagesToolTests` cover both the old shape and the new structured shape.

### Performance

DASL `LIKE` on `textdescription` (body) is slow on huge mailboxes — Outlook reads each message individually to apply the LIKE. For inboxes under ~5k items this is fine. For larger:
- Header-only filters (`from`, `subject_contains`, dates, flags, importance, attachments) are fast — they're MAPI properties indexed by Outlook.
- Body-text filters (`query`, `body_contains`) are the slow ones.

A Phase 3.x perf follow-up will optionally route body-text queries through `Application.AdvancedSearch` (Windows Search indexed search, async, ~10× faster on big mailboxes). For Phase 3a we accept the DASL cost.

---

## 8. ConversationStore + Lifecycle

Reuses the Phase 2 `ConversationStore` class verbatim. Each `InboxCopilotPane` constructs its own instance and disposes it on Explorer close.

Manual Clear and Copy-to-clipboard already work in the WebUI (Phase 2). No changes needed.

---

## 9. WebUI Changes

### `chat.js` additions

```js
// Push the quick-action chip row above the composer. Each chip is
// { label: "Display text", prompt: "Pre-fill into textarea" }.
// Auto-send is hard-coded: clicking a chip both fills and sends.
outlookai.setQuickActions(chips) { ... }

// Existing setContextStrip enhanced to also accept folder + selection
// when present. Backward-compatible with the compose-pane shape.
outlookai.setContextStrip({
  subject?, recipients?, thread?,    // compose shape (Phase 2)
  folder?, unread_count?, selection? // inbox shape (Phase 3a)
})
```

### `index.html` additions

A new `<div id="quickActions" class="quick-actions"></div>` between `#messages` and `#composer`. `chat.js`'s `setQuickActions` rebuilds its children on every call.

### `styles.css` additions

Small CSS for the chip row:
- Horizontal flex, wrapping, gap 4px.
- Chips styled like small pill buttons (background `--tool-bg`, border `--tool-border`, hover `--accent`).
- Hidden when empty (CSS `:empty` selector).

### Theming

Same three themes as Phase 2 (light / dark / high-contrast). Chips use CSS variables so they pick up the theme automatically.

---

## 10. Testing

### Unit tests (xUnit, existing test project)

| Area | New tests |
|---|---|
| `OutlookGetCurrentSelectionTool` | `Execute_ProjectsSelectionToJson`, `Execute_RespectsMaxItems`, `Execute_EmptySelection_ReturnsEmptyList`, `Execute_IncludeFullBodies_AddsBodyPlaintext` |
| `BuildRestrictFilter` (enhanced) | `BuildsCompositeAndFilter_FromAllStructuredFields`, `EmptyArgs_ProducesNullFilter`, `QueryOnly_StillUsesSubjectOrBodyClause`, `EscapesQuotesInValues`, `ImportanceLow_UsesCorrectMapiProperty` |
| `OutlookSearchMessagesTool` (extended schema) | `Execute_PassesStructuredFieldsToSurface`, `Execute_BackwardCompat_QueryOnlyStillWorks` |
| `OutlookCountMessagesTool` (extended schema) | `Execute_PassesStructuredFieldsToSurface` |
| `InboxCopilotController` (where unit-testable) | Static helper `BuildSystemPromptForExplorerState(folderName, unreadCount, selection)` returns the expected prompt shape. Quick-action chip computation pure function `ComputeChipsForSelectionCount(int n)` returns the expected list. |
| `Ribbon.cs` routing | `OnAIAssistantClick_WithInspectorActive_RoutesToInspectorPane`, `OnAIAssistantClick_WithExplorerActive_RoutesToExplorerPane`. Fake `Application` exposes a settable `ActiveWindow`. |

Estimated +12-15 tests. Brings the suite from 109 → ~125.

### Deliberately not unit-tested

- WebView2 pane behavior (manual smoke).
- Real Outlook Explorer + Selection events (manual smoke).
- Multi-Explorer isolation (manual smoke).

### Manual smoke (new section in the checklist)

1. Open Outlook → AI Assistant button visible on TabMail (Home tab) in addition to TabNewMailMessage in compose windows.
2. Click AI Assistant from the Inbox view → pane opens on the right. Width 340 px. Single chat surface, no tabs.
3. Context strip shows `In: <folder> (<n> unread)`. Selection line hidden.
4. Click a message in the reading pane → context strip refreshes to add `Selected: <subject> — <sender>`. Quick-action chip row also refreshes to show "Summarize this thread" + "Draft a reply".
5. Click "What needs my attention?" chip → prompt fills the textarea AND auto-sends. Streaming response appears with tool cards.
6. Type "Find unread emails from Jane about Q4 with attachments." The model should make ONE search call with the structured fields populated, not multiple. Verify in `%LOCALAPPDATA%\OutlookAI\trace.log`.
7. Open a second Explorer (`Ctrl+N` from Outlook's File menu, "Open in New Window") → click AI Assistant → second pane opens, independent of the first. Conversations don't bleed across.
8. Switch folders in one Explorer → context strip updates.
9. Compose a new email (open an Inspector) → AI Assistant button on the compose ribbon still opens the existing Phase 2 pane. Both surfaces coexist.
10. Close Explorer → pane + conversation disappears.

---

## 11. Risk Register

| Risk | Mitigation |
|---|---|
| `Application.ActiveWindow()` returns something other than Inspector/Explorer | Defensive fallback to a "no active window" message dialog. Should be impossible in practice (ribbon only renders in those contexts), but cheap to guard. |
| `Explorer.SelectionChange` fires on a fast-clicking user; chip-refresh debounce | Throttle chip refresh to one update per 200 ms via a single-shot timer. The model never sees stale state because the system prompt is rebuilt at turn-send time, not at chip-refresh time. |
| Non-MailItem in Selection (meeting / contact / task) | Filtered out for Phase 3a. The model can't act on them — Phase 3b/c/d will add tools that surface them properly. |
| DASL filter syntax errors crash the Restrict call | Defensive: `try/catch (COMException)` around `folder.Items.Restrict(filter)`; on failure, fall back to unfiltered enumeration with a tool-error response so the model knows search failed. |
| Slow body-text search on huge mailboxes | Document as known. Phase 3.x perf work will add `AdvancedSearch` path for body queries on inboxes > 5k items. |
| User has both compose pane and Explorer pane open with the same model | Each has its own conversation. Reasoning effort, write tools, and model selection are read from `Config` at turn-build time, so admin changes apply to both panes' next turn. No state leaks across. |
| WebView2 fails to initialize in the Explorer pane (e.g., runtime uninstalled mid-session) | Existing Phase 2 fallback label rendering inherited automatically. |
| `Explorer.Close` doesn't always fire on Outlook shutdown | The pane is disposed by VSTO's CTP collection when the host shuts down anyway. Defensive cleanup is best-effort. |

---

## 12. Implementation Boundaries

### What this spec produces

- A working Inbox Copilot pane opened from the Explorer ribbon.
- Folder + selection awareness in the system prompt and the UI.
- Quick-action chips with auto-send.
- The new `outlook_get_current_selection` tool.
- Enhanced structured search across both `outlook_search_messages` and `outlook_count_messages`.
- ~12-15 new unit tests.
- Manual smoke checklist updates.

### What this spec explicitly does NOT include

- Calendar tools (Phase 3b).
- Contacts tools (Phase 3c).
- Tasks tools (Phase 3d).
- Send tool (Phase 3e — different design entirely because of the confirmation gate).
- Voice in either pane (Phase 3.x).
- Admin-configurable chip list (Phase 3.x).
- Persistent conversations across Outlook restarts (deliberate v1 choice; revisit if users miss it).
- `AdvancedSearch`-based body-text indexing for performance (Phase 3.x perf work).
- Refresh-on-model-change for the in-pane reasoning dropdowns (Phase 2 limitation carried forward).

---

## 13. Done Criteria

Phase 3a is "done" when:

1. Solution rebuilds cleanly with the new files; no warnings.
2. All unit tests pass (target ~125 total).
3. Manual smoke checklist § 10 passes in a real Outlook session.
4. The `feature/codex-oauth-migration` branch on `origin` includes the spec, the implementation plan, every implementation commit, and an updated README section.
5. A trace-log run of "Find unread emails from Jane about Q4 with attachments" shows exactly one `outlook_search_messages` call with the structured fields populated — confirming the model uses the new schema instead of multiple sequential calls.

No merge to `master` is required for Phase 3a to count as done — same branch policy as Phases 1 and 2.

---

## Notes

This spec is intentionally additive — almost nothing in Phase 2 changes. The Inbox Copilot pane reuses Phase 2's WebView2 surface, chat.js, ConversationStore, ChatEventSink, CodexChatService.RunTurnAsync, OutlookThreadMarshaller, and OutlookToolHost. The only Phase 2 file with non-additive changes is `LiveOutlookSurface.cs` (new selection method + enhanced search-filter builder).

This keeps Phase 3a small and reversible: if the Inbox Copilot UX doesn't pan out, we can revert by removing the new files + the `TabMail` ribbon entry without disturbing the compose-pane experience.
