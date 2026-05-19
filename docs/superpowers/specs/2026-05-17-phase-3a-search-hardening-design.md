# Phase 3a Search Hardening Design

## Problem

The previous SSE delta-accumulation fix worked: production traces now show non-empty tool arguments reaching `outlook_search_messages` and `outlook_count_messages`.

The remaining search failures are design-level issues in the search tool shape and execution path.

Trace evidence from `C:\Users\MDASR\AppData\Local\OutlookAI\trace.log` after the deployed SSE fix:

- User query at line 38: `wHat was my first email ever`.
- The model made many `outlook_count_messages` / `outlook_search_messages` probes instead of one direct oldest-first search.
- Dispatch args were populated, but included unintended defaults: `has_attachment:false`, `is_unread:false`, `is_flagged:false`, `importance:"normal"`, plus blank strings.
- `LiveOutlookSurface.BuildRestrictFilter` treated those defaults as real filters, producing DASL clauses like `hasattachment = 0`, `read = 1`, `flag_status <> 2`, and `importance = 1`.
- That is wrong for "first email ever" because the user did not ask to exclude unread, flagged, attached, low-importance, or high-importance mail.
- The tool has no way to request `oldest` ordering, so the model attempted date probing and narrowing, which is slow and brittle.

The fix needs to address the search contract, not just prompt wording.

## Goals

- Make common natural-language mailbox searches map to one or two precise tool calls.
- Support oldest-first and newest-first retrieval directly.
- Support current-folder-first behavior with all-mail broadening.
- Prevent false/default boolean pollution from becoming accidental Outlook filters.
- Improve search latency by avoiding unnecessary candidate counting in `SearchMessages`.
- Preserve current compact tool status UI.

## Non-Goals

- No local mailbox index or SQLite/Lucene search engine in this phase.
- No attachment content indexing in this phase.
- No delete/send behavior changes.
- No UI redesign.
- No broad refactor outside search tool parsing, schema, prompt guidance, and Outlook execution.

## Search Semantics

### New Model-Facing Fields

Add these fields to both `outlook_search_messages` and `outlook_count_messages` schemas:

| Field | Values | Behavior |
| --- | --- | --- |
| `sort_order` | `newest`, `oldest` | Controls `ReceivedTime` sort direction. Default `newest`. |
| `scope` | `current_folder`, `all_mail`, `auto` | Controls folder selection. Default `auto` in Inbox Copilot. |
| `attachment_filter` | `any`, `with`, `without` | Replaces model-facing `has_attachment`. `any` means no attachment clause. |
| `read_status` | `any`, `read`, `unread` | Replaces model-facing `is_unread`. `any` means no read clause. |
| `flag_status` | `any`, `flagged`, `unflagged` | Replaces model-facing `is_flagged`. `any` means no flag clause. |
| `importance_filter` | `any`, `low`, `normal`, `high` | Replaces model-facing `importance`. `any` means no importance clause. |

Keep the existing fields:

- `query`
- `from`
- `subject_contains`
- `body_contains`
- `folder_id`
- `date_from`
- `date_to`
- `max_results`

### Hidden Backward Compatibility

Do not advertise these old fields to the model anymore:

- `has_attachment`
- `is_unread`
- `is_flagged`
- `importance`

The parser may continue to accept them for old history and tests, but new model calls should be steered to the tri-state fields. Hidden backward-compatible handling must not let a default-looking old-shape payload pollute searches.

Compatibility rule: old boolean `true` values are still treated as explicit filters (`has_attachment:true`, `is_unread:true`, `is_flagged:true`). Old boolean `false` values are ignored unless the new tri-state field explicitly requests the negative state (`attachment_filter=without`, `read_status=read`, `flag_status=unflagged`). Old `importance:"normal"` is ignored unless `importance_filter=normal` is present.

Specifically, old-shape defaults like this should not produce active filters:

```json
{
  "query": "",
  "from": "",
  "subject_contains": "",
  "body_contains": "",
  "has_attachment": false,
  "is_unread": false,
  "is_flagged": false,
  "importance": "normal"
}
```

Explicit new-shape filters should still work:

```json
{
  "read_status": "read",
  "attachment_filter": "without",
  "flag_status": "unflagged",
  "importance_filter": "normal"
}
```

## Scope Behavior

### `current_folder`

Use the active Explorer's current folder. If no Explorer is available, fall back to Inbox.

If `folder_id` is provided, it takes precedence over `scope` and searches only that resolved folder. This keeps explicit folder choices deterministic.

### `all_mail`

Search all searchable mail folders across stores:

- Inbox
- Sent
- Archive
- user-created mail folders

Exclude noisy or non-final folders by default:

- Deleted Items
- Junk Email
- Drafts
- Outbox
- Sync Issues
- RSS Feeds

### `auto`

Search `current_folder` first. If zero hits are returned, broaden to `all_mail`.

Prompt/tool descriptions should instruct the model to use `scope=all_mail` directly when the user says broad terms such as:

- ever
- any email
- anywhere
- everything
- all mail

## Outlook Execution Path

### `SearchMessages`

For each selected folder:

1. Build the DASL filter from the normalized `SearchMessagesArgs`.
2. Use `folder.Items.Restrict(filter)` when there is a filter; otherwise use `folder.Items`.
3. Sort by `ReceivedTime`:
   - `newest` -> `items.Sort("[ReceivedTime]", true)`
   - `oldest` -> `items.Sort("[ReceivedTime]", false)`
4. Enumerate only until `max_results` hits from that folder.

For `all_mail`, collect up to `max_results` hits per folder, merge them in memory, sort globally by `ReceivedAt`, and return the top `max_results`.

For `auto`, first search the current folder. If it returns zero hits, retry with the all-mail folder set.

Do not call `items.Count` in `SearchMessages` just for diagnostics. Counting restricted Outlook item sets is expensive and not needed to return summaries.

### `CountMessages`

`CountMessages` should continue to use `Items.Count`, because counting is its purpose. It should support the same normalized filters and `scope` behavior. `scope=all_mail` sums counts across the selected folder set.

## Prompt And Tool Description Guidance

Update `ToolCatalogSchema` and `InboxCopilotPromptBuilder` with examples:

- "What was my first email ever?" -> `scope=all_mail`, `sort_order=oldest`, `max_results=1`, no filters.
- "Latest email from Jane with an attachment" -> `from=Jane`, `attachment_filter=with`, `sort_order=newest`, `max_results=1`.
- "Find invoices from before 2020" -> `query` or `body_contains` for `invoice`, `date_to=2020-01-01T00:00:00Z`, no default boolean filters.
- "Find an email with EIN" -> search `EIN` current folder first, broaden to all mail if zero; if still zero, try `Employer Identification Number`.

The model should not pass empty/default argument objects for targeted searches.

## Diagnostics

Improve trace visibility:

- Expand dispatch argument trace from 200 characters to either 500 characters or structured fields for search args.
- Add `SearchMessages` trace fields: `scope`, `sort_order`, `folders_searched`, `returned`, and whether `auto` broadened.
- Keep tracing defensive; trace failures must never break a tool call.

## Tests

Add tests before production code changes.

### Parser Tests

- Old default-looking boolean payloads do not create active filters.
- New tri-state fields map to `SearchMessagesArgs` correctly.
- `any` tri-state values produce unset/null filter semantics.
- Blank strings are normalized away.

### Schema Tests

- Search/count schemas advertise `scope`, `sort_order`, and tri-state filter fields.
- Search/count schemas do not advertise old optional booleans.
- Tool descriptions include oldest-first, all-mail, and EIN examples.

### Filter Tests

- `read_status=unread` produces `urn:schemas:httpmail:read = 0`.
- `read_status=read` produces `urn:schemas:httpmail:read = 1`.
- `read_status=any` produces no read clause.
- Equivalent tests for attachment, flag, and importance filters.
- "first email ever" normalized args produce no false/default DASL clauses.

### Execution Tests

- `sort_order=oldest` sorts ascending and returns oldest first.
- `sort_order=newest` sorts descending and returns newest first.
- `scope=auto` searches current folder and broadens only after zero hits.
- `scope=all_mail` merges per-folder results and globally sorts.
- `SearchMessages` does not call `Items.Count` for diagnostics.
- `CountMessages` still counts and supports `scope=all_mail`.

## Acceptance Cases

### Oldest/Newest

- "What was my first email ever?" returns the oldest message across all mail in one `outlook_search_messages` call, or at most two if `auto` broadens.
- "What is my newest email?" returns the newest message in the current folder unless the user says all mail.

### Sender

- "Find my earliest email from Bob" uses `from=Bob`, `sort_order=oldest`, `scope=auto`.
- "Latest email from Jane with an attachment" uses `from=Jane`, `attachment_filter=with`, `sort_order=newest`.

### Keyword/Body

- "Find an email with EIN" searches current folder first, broadens to all mail if zero, and can try `Employer Identification Number` if literal `EIN` misses.
- "Find invoices from before 2020" uses an invoice keyword plus `date_to=2020-01-01T00:00:00Z`, with no default boolean filters.

### State Filters

- "Show unread emails from last week" uses `read_status=unread` and a date range, with no attachment/flag/importance clauses.
- "Find read emails with no attachments from Alice" uses `read_status=read`, `attachment_filter=without`, and `from=Alice`.
- "Find high importance flagged emails" uses `importance_filter=high` and `flag_status=flagged`.

### Scope

- "Search all mail for Cisco UC560" uses `scope=all_mail`.
- "Search this folder for Cisco UC560" uses `scope=current_folder`.
- "Find any email from Kirkland" uses `scope=all_mail` or `scope=auto` that broadens if current folder misses.

### Safety

- Blank/default string fields are ignored.
- `any` tri-state values produce no DASL clause.
- Old hidden boolean false fields from prior model history do not pollute searches.

## Rollout

Implement behind the existing tool names so the UI does not change. After tests pass:

1. Build Debug and run the full test suite.
2. Publish Release.
3. Install elevated with the correct `-SourcePath`.
4. Smoke-test the acceptance cases in Outlook.
5. Confirm trace logs show populated args, safe filters, correct scope/sort, and no unintended default clauses.
