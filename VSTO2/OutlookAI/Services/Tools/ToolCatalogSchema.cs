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
