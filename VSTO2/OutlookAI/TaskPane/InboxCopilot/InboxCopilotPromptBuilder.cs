using System.Text;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Builds the per-turn system instructions for the Inbox Copilot.
    /// Pure function over current folder state + selection so the chat
    /// service always sends fresh context. Extracted into its own class
    /// so it can be unit-tested without spinning up the controller.
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
                    {
                        sb.Append("  From: ").AppendLine(m.From);
                    }
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
            sb.AppendLine("Tips for searching:");
            sb.AppendLine("- Always translate the user's words into the structured fields of");
            sb.AppendLine("  outlook_search_messages / outlook_count_messages. Do NOT pass an");
            sb.AppendLine("  empty argument object - that returns the newest 25 messages in the");
            sb.AppendLine("  folder, which is almost never what the user wanted.");
            sb.AppendLine("- Sender names go in 'from', not 'query'.");
            sb.AppendLine("- Dates go in 'date_from' / 'date_to' as ISO-8601 UTC. 'Before 2020'");
            sb.AppendLine("  is date_to=2020-01-01T00:00:00Z. 'Last week' is date_from = today");
            sb.AppendLine("  minus 7 days. 'This year' is date_from = Jan 1 of current year.");
            sb.AppendLine("- A single keyword the user wants to find ('EIN', 'invoice number',");
            sb.AppendLine("  'contract') goes in 'query' or 'body_contains'.");
            sb.AppendLine("- After search, call outlook_read_message on the best-matching id to");
            sb.AppendLine("  get the full body before answering.");
            sb.AppendLine();
            sb.AppendLine("Reply concisely; the user is busy.");
            return sb.ToString();
        }
    }
}
