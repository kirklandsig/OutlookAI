using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Static JSON schema definitions for every Phase 2 tool. Same schema is
    /// (a) sent to the model in the Responses request as <c>tools[]</c>
    /// entries, and (b) used by <c>ToolDispatcher</c> to validate arguments
    /// before dispatch.
    /// </summary>
    public static class ToolCatalogSchema
    {
        public static JArray BuildResponsesToolsArray(bool includeWriteTools)
        {
            var arr = new JArray
            {
                BuildToolEntry("outlook_get_current_compose_state",
                    "Read the current compose-window state (subject, recipients, body, thread, attachments). Set include_full_body=true for the full body instead of a 1000-char summary.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject(
                            new JProperty("include_full_body", new JObject(
                                new JProperty("type", "boolean"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_list_folders",
                    "List the user's mail folders (max depth 6, max 200 nodes).",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject()),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_search_messages",
                    "Search messages. Combine any subset of filters via AND. Returns id+metadata+snippet for up to max_results (default 25, hard cap 100). "
                    + "ALWAYS translate the user's natural-language query into the structured fields below; do NOT dump the whole user sentence into 'query'. "
                    + "PRESERVE the user's exact spelling and whitespace for company / vendor / brand names - 'IT Creations' is one phrase with a space; do NOT compact it to 'ITCreations'. Whitespace is a literal character in the Outlook LIKE match. "
                    + "Examples: "
                    + "User says 'what was my first email ever' -> {scope:'all_mail', sort_order:'oldest', max_results:1}. "
                    + "User says 'emails from Alice last week' -> {from:'Alice', date_from:<7d ago ISO>, date_to:<today ISO>}. "
                    + "User says 'messages I sent to Susan about servers' -> first use outlook_list_folders to find Sent Items/Outbound, then {to:'Susan', query:'servers', folder_id:<sent folder id>, max_results:25}; do not put recipient in query or from. For multiple recipients, make separate precise searches. "
                    + "User says 'quotes from IT Creations' (vendor lookup) -> {from:'IT Creations', scope:'all_mail'}. Put the vendor name in 'from', not 'query'; 'from' matches both display name and email address case-insensitively. "
                    + "User says 'email from before 2020 with the EIN' -> {query:'EIN', date_to:'2020-01-01T00:00:00Z'}. "
                    + "User says 'find an email with EIN' -> {query:'EIN', scope:'auto'}; if zero, try {query:'Employer Identification Number', scope:'auto'}. "
                    + "User says 'unread invoices' -> {body_contains:'invoice', read_status:'unread'}. "
                    + "User says 'flagged messages with attachments' -> {flag_status:'flagged', attachment_filter:'with'}. "
                    + "Zero-result fallback ladder: if a search returns 0 hits, do NOT report 'no emails' immediately - try at least three variations: "
                    + "(1) drop any AND filters you added beyond what the user explicitly asked for (e.g. body_contains:'quote' if they said 'quotes from X', try without it first); "
                    + "(2) move the company name from 'query' to 'from' (or vice versa) - vendors usually match 'from'; "
                    + "(3) try a compact spelling without spaces only AFTER the user's exact spelling fails (e.g. 'IT Creations' first, then 'itcreations'); "
                    + "(4) broaden scope to all_mail if you started with auto. "
                    + "Never send default false filters such as has_attachment:false or is_unread:false; use attachment_filter/read_status/flag_status/importance_filter and use 'any' when not requested. "
                    + "For 'I sent' / sent / outgoing report requests, prefer outlook_list_folders to discover Sent Items or Outbound and pass folder_id rather than scope:'all_mail' when possible. "
                    + "Do not use max_results:100 on broad all_mail targeted lookups; use 25 unless the user explicitly asks for a large row count or a specific folder is selected. "
                    + "If you pass NO filters, the tool returns the newest 25 messages in the current folder - almost never what the user asked for. "
                    + "Prefer one precise call over many. After search, use outlook_read_message on the most-relevant id for full body. "
                    + "For tabular Excel exports over many messages, prefer metadata-only synthesis: the snippet field already contains the first ~200 chars of each match, so build Excel rows directly from search results and only call outlook_read_messages when a column genuinely requires full body content. Reading full bodies for 100+ messages will exceed the context window.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject(
                            new JProperty("query",            new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Free-form keyword(s) matched against subject + body (e.g. 'EIN', 'invoice', 'contract'). Do NOT put dates, sender names, recipient names, or other structured info here - use the dedicated fields."))),
                            new JProperty("from",             new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Sender substring; matches display name OR email (case-insensitive). Example: 'Alice' or 'alice@example.com'."))),
                            new JProperty("to",               new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Recipient substring; for multiple recipients, make separate precise searches rather than concatenating names."))),
                            new JProperty("subject_contains", new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Substring match on subject only."))),
                            new JProperty("body_contains",    new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Substring match on body only."))),
                            new JProperty("scope",            new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("current_folder","all_mail","auto")),
                                                              new JProperty("description","Default auto. Use all_mail directly for 'ever', 'any email', 'anywhere', 'everything', or 'all mail'. folder_id overrides scope."))),
                            new JProperty("sort_order",       new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("newest","oldest")),
                                                              new JProperty("description","Default newest. Use oldest for 'first', 'earliest', or 'oldest'."))),
                            new JProperty("attachment_filter", new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","with","without")),
                                                              new JProperty("description","Use with only when the user asks for attachments; without only when they ask for no attachments; otherwise any."))),
                            new JProperty("read_status",      new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","read","unread")),
                                                              new JProperty("description","Use unread/read only when explicitly requested; otherwise any."))),
                            new JProperty("flag_status",      new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","flagged","unflagged")),
                                                              new JProperty("description","Use flagged/unflagged only when explicitly requested; otherwise any."))),
                            new JProperty("importance_filter", new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","low","normal","high")),
                                                              new JProperty("description","Use low/normal/high only when explicitly requested; otherwise any."))),
                            new JProperty("folder_id",        new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Default: Inbox. Use outlook_list_folders to discover folder ids."))),
                            new JProperty("date_from",        new JObject(new JProperty("type","string"),
                                                                          new JProperty("format","date-time"),
                                                                          new JProperty("description","Inclusive lower bound, ISO-8601 UTC. For 'last week' compute 7 days ago. For 'this year' use Jan 1 of current year."))),
                            new JProperty("date_to",          new JObject(new JProperty("type","string"),
                                                                          new JProperty("format","date-time"),
                                                                          new JProperty("description","Exclusive upper bound, ISO-8601 UTC. For 'before 2020' use '2020-01-01T00:00:00Z'. For 'older than 2 years' use today minus 2y."))),
                            new JProperty("max_results",      new JObject(new JProperty("type","integer"),
                                                                          new JProperty("minimum",1),
                                                                          new JProperty("maximum",100),
                                                                          new JProperty("description","Default 25. Raise to 100 only when summarizing a folder; never for a targeted lookup."))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_read_message",
                    "Fetch one message by id. Body always plaintext; truncated at 32 KB.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("message_id")),
                        new JProperty("properties", new JObject(
                            new JProperty("message_id", new JObject(new JProperty("type","string"))),
                            new JProperty("include_full_body", new JObject(new JProperty("type","boolean"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_count_messages",
                    "Count messages matching the given filters without returning bodies. Same filter fields as outlook_search_messages; same rule: translate the user's intent into structured fields, do NOT put everything in 'query'. Examples: 'how many unread from Bob' -> {from:'Bob', read_status:'unread'}; 'how many emails this year' -> {date_from:<Jan 1 ISO>}. Empty filters => current-folder count unless scope:'all_mail'.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("properties", new JObject(
                            new JProperty("query",            new JObject(new JProperty("type","string"))),
                            new JProperty("from",             new JObject(new JProperty("type","string"))),
                            new JProperty("to",               new JObject(new JProperty("type","string"),
                                                              new JProperty("description","Recipient substring; for multiple recipients, make separate precise searches rather than concatenating names."))),
                            new JProperty("subject_contains", new JObject(new JProperty("type","string"))),
                            new JProperty("body_contains",    new JObject(new JProperty("type","string"))),
                            new JProperty("scope",            new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("current_folder","all_mail","auto")))),
                            new JProperty("sort_order",       new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("newest","oldest")))),
                            new JProperty("attachment_filter", new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","with","without")))),
                            new JProperty("read_status",      new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","read","unread")))),
                            new JProperty("flag_status",      new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","flagged","unflagged")))),
                            new JProperty("importance_filter", new JObject(new JProperty("type","string"),
                                                              new JProperty("enum", new JArray("any","low","normal","high")))),
                            new JProperty("folder_id",        new JObject(new JProperty("type","string"))),
                            new JProperty("date_from",        new JObject(new JProperty("type","string"),
                                                                          new JProperty("format","date-time"))),
                            new JProperty("date_to",          new JObject(new JProperty("type","string"),
                                                                          new JProperty("format","date-time"))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_read_messages",
                    "Bulk-read message details by short ID array. Returns subject/sender/date/body/attachments per id. Use this instead of multiple outlook_read_message calls; 5-10x faster when you have many IDs from a prior outlook_search_messages and you need bodies (action items, topic-status, conversation summary reports). Avoid calling on more than ~25 IDs when each body is long. For large result sets you intend to export, work from search metadata + snippet instead of bulk-reading bodies.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("ids")),
                        new JProperty("properties", new JObject(
                            new JProperty("ids", new JObject(
                                new JProperty("type", "array"),
                                new JProperty("items", new JObject(new JProperty("type", "string"))),
                                new JProperty("description", "Short message IDs from a prior outlook_search_messages."))),
                            new JProperty("include_body", new JObject(
                                new JProperty("type", "boolean"),
                                new JProperty("description", "Default true. Set false for a lightweight metadata-only read."))),
                            new JProperty("max_items", new JObject(
                                new JProperty("type", "integer"),
                                new JProperty("minimum", 1),
                                new JProperty("maximum", 100),
                                new JProperty("description", "Default 25, hard cap 100."))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_aggregate_messages",
                    "Group matching messages by sender, day, or folder and return the top-N buckets by count. Use this instead of calling outlook_count_messages many times when the user wants statistics ('top 10 senders this month', 'busiest days last week', 'breakdown by folder').",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("group_by")),
                        new JProperty("properties", new JObject(
                            new JProperty("scope", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("enum", new JArray("auto", "current_folder", "all_mail")),
                                new JProperty("description", "Default auto."))),
                            new JProperty("folder_id", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional explicit folder id. Overrides scope when provided."))),
                            new JProperty("date_from", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("format", "date-time"),
                                new JProperty("description", "Optional lower bound on ReceivedTime."))),
                            new JProperty("date_to", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("format", "date-time"),
                                new JProperty("description", "Optional upper bound on ReceivedTime."))),
                            new JProperty("from", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional sender name or email substring."))),
                            new JProperty("subject_contains", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional subject substring."))),
                            new JProperty("body_contains", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional body substring."))),
                            new JProperty("group_by", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("enum", new JArray("sender", "day", "folder")),
                                new JProperty("description", "How to bucket matching messages."))),
                            new JProperty("top_n", new JObject(
                                new JProperty("type", "integer"),
                                new JProperty("minimum", 1),
                                new JProperty("maximum", 100),
                                new JProperty("description", "Default 10, hard cap 100."))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_export_excel",
                    "Create a spreadsheet Excel .xlsx tabular export. Use when the user asks for spreadsheet, Excel, xlsx, or tabular export; pass typed columns and rows. Produces a styled workbook with bold/frozen header, autofilter, and per-column formatting. Best for vendor lists, message tables, aggregations by sender/day, and structured search-result exports. Do NOT use for prose, narrative reports, or arbitrary text; choose outlook_export_pdf for those. Maximum 10000 rows; aggregate or filter first if more. Example flow: search messages first, then export projected rows. File is saved to ~\\Documents\\OutlookAI\\Reports\\; the UI surfaces Open/Show-in-folder later. For exports over 50+ messages, prefer projecting columns (subject, from, to, received_at, snippet) from outlook_search_messages results without reading full bodies; this is metadata-only synthesis and stays within the context window. Reading full bodies for 100+ messages will exceed context.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("columns", "rows")),
                        new JProperty("properties", new JObject(
                            new JProperty("filename_hint", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional base filename; sanitized and timestamped, with a safe default when omitted."))),
                            new JProperty("sheet_name", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional worksheet name; defaults to filename_hint."))),
                            new JProperty("columns", new JObject(
                                new JProperty("type", "array"),
                                new JProperty("description", "Column definitions in display order. Type controls Excel formatting: date/datetime for ISO dates, number/currency for numeric cells, boolean for true/false, text otherwise."),
                                new JProperty("items", new JObject(
                                    new JProperty("type", "object"),
                                    new JProperty("required", new JArray("name", "type")),
                                    new JProperty("properties", new JObject(
                                        new JProperty("name", new JObject(
                                            new JProperty("type", "string"),
                                            new JProperty("description", "Column header text."))),
                                        new JProperty("type", new JObject(
                                            new JProperty("type", "string"),
                                            new JProperty("enum", new JArray("text", "date", "datetime", "number", "currency", "boolean")),
                                            new JProperty("description", "Cell type used for validation and per-column Excel formatting."))))),
                                    new JProperty("additionalProperties", false))))),
                            new JProperty("rows", new JObject(
                                new JProperty("type", "array"),
                                new JProperty("description", "Row arrays; each row length must equal columns.length. Cells may be strings, numbers, ISO dates, or booleans matching the column type."),
                                new JProperty("items", new JObject(
                                    new JProperty("type", "array"))))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_export_pdf",
                    "Save a polished, printable PDF report. Use when the user asks for PDF, printable, document, or a shareable report. Pass polished markdown with a clear title, headings, tables, and lists; the tool renders a clean A4 document with a header bar and no chat UI chrome. Prefer outlook_export_excel for tabular structured data, real spreadsheets, Excel workbooks, and rows/columns users may sort or analyze; use PDF for narrative reports, summaries, action items, digest-style outputs, and client-ready documents. You may compose fresh markdown specifically for the PDF rather than reusing chat text. Example flow: search/read messages for a weekly or customer report, synthesize sections and action items as markdown, then export the PDF. File is saved to ~\\Documents\\OutlookAI\\Reports\\; the UI surfaces an Open/Show-in-folder card. Max content_markdown length: 250000 chars.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("content_markdown")),
                        new JProperty("properties", new JObject(
                            new JProperty("filename_hint", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("description", "Optional base filename; sanitized and timestamped. Defaults to title or OutlookAI-Report when omitted."))),
                            new JProperty("title", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("maxLength", 200),
                                new JProperty("description", "Optional but recommended report title, max 200 chars."))),
                            new JProperty("subtitle", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("maxLength", 400),
                                new JProperty("description", "Optional report subtitle, max 400 chars."))),
                            new JProperty("content_markdown", new JObject(
                                new JProperty("type", "string"),
                                new JProperty("maxLength", 250000),
                                new JProperty("description", "Required markdown body, max 250000 chars. Supports headings, GFM tables, lists, blockquotes, code blocks, bold/italic, and links; inline images are stripped."))))),
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_get_current_selection",
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
                        new JProperty("additionalProperties", false))),

                BuildToolEntry("outlook_list_recent_threads_with",
                    "List the most recent conversation threads involving a specific recipient (Inbox + Sent), grouped by ConversationID.",
                    new JObject(
                        new JProperty("type", "object"),
                        new JProperty("required", new JArray("recipient_email")),
                        new JProperty("properties", new JObject(
                            new JProperty("recipient_email", new JObject(new JProperty("type","string"))),
                            new JProperty("max_threads", new JObject(new JProperty("type","integer"),
                                                                     new JProperty("minimum",1),
                                                                     new JProperty("maximum",20))))),
                        new JProperty("additionalProperties", false))),
            };

            if (includeWriteTools)
            {
                // Per-tool gate against Config.EnabledWriteTools so an admin
                // can disable just one or two of the safe-writes without
                // losing the others. Default = all four enabled.
                var enabled = Config.EnabledWriteTools ?? new System.Collections.Generic.HashSet<string>(Config.AllWriteTools);

                if (enabled.Contains("outlook_create_draft"))
                {
                    arr.Add(BuildToolEntry("outlook_create_draft",
                        "Create a draft in the Drafts folder. Never sends. If in_reply_to_message_id is given, seeds the draft via MailItem.Reply().",
                        new JObject(
                            new JProperty("type", "object"),
                            new JProperty("required", new JArray("subject","body_plaintext")),
                            new JProperty("properties", new JObject(
                                new JProperty("subject", new JObject(new JProperty("type","string"))),
                                new JProperty("body_plaintext", new JObject(new JProperty("type","string"))),
                                new JProperty("to", new JObject(new JProperty("type","array"),
                                                                new JProperty("items", new JObject(new JProperty("type","string"))))),
                                new JProperty("cc", new JObject(new JProperty("type","array"),
                                                                new JProperty("items", new JObject(new JProperty("type","string"))))),
                                new JProperty("in_reply_to_message_id", new JObject(new JProperty("type","string"))))),
                            new JProperty("additionalProperties", false))));
                }

                if (enabled.Contains("outlook_mark_as_read"))
                {
                    arr.Add(BuildToolEntry("outlook_mark_as_read",
                        "Set or clear the UnRead flag on a message.",
                        new JObject(
                            new JProperty("type", "object"),
                            new JProperty("required", new JArray("message_id","read")),
                            new JProperty("properties", new JObject(
                                new JProperty("message_id", new JObject(new JProperty("type","string"))),
                                new JProperty("read", new JObject(new JProperty("type","boolean"))))),
                            new JProperty("additionalProperties", false))));
                }

                if (enabled.Contains("outlook_flag_message"))
                {
                    arr.Add(BuildToolEntry("outlook_flag_message",
                        "Set follow-up flag status: none | todo | complete.",
                        new JObject(
                            new JProperty("type", "object"),
                            new JProperty("required", new JArray("message_id","flag")),
                            new JProperty("properties", new JObject(
                                new JProperty("message_id", new JObject(new JProperty("type","string"))),
                                new JProperty("flag", new JObject(new JProperty("type","string"),
                                                                  new JProperty("enum", new JArray("none","todo","complete")))))),
                            new JProperty("additionalProperties", false))));
                }

                if (enabled.Contains("outlook_set_category"))
                {
                    arr.Add(BuildToolEntry("outlook_set_category",
                        "Replace a message's Categories with the single given value.",
                        new JObject(
                            new JProperty("type", "object"),
                            new JProperty("required", new JArray("message_id","category")),
                            new JProperty("properties", new JObject(
                                new JProperty("message_id", new JObject(new JProperty("type","string"))),
                                new JProperty("category", new JObject(new JProperty("type","string"))))),
                            new JProperty("additionalProperties", false))));
                }
            }

            return arr;
        }

        private static JObject BuildToolEntry(string name, string description, JObject parameters)
        {
            return new JObject(
                new JProperty("type", "function"),
                new JProperty("name", name),
                new JProperty("description", description),
                new JProperty("parameters", parameters));
        }
    }
}
