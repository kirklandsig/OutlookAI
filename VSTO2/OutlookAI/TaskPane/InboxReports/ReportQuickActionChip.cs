using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// One quick-action chip in the Inbox Reports pane. Clicking the chip
    /// prefills the chat input with TemplateText. The user can edit
    /// any [placeholders] before sending; the system prompt instructs
    /// the model to ask for clarification if a placeholder is still
    /// present at tool-call time.
    /// </summary>
    public sealed class ReportQuickActionChip
    {
        public string Label { get; set; }
        public string TemplateText { get; set; }

        public static IReadOnlyList<ReportQuickActionChip> Defaults()
        {
            return new[]
            {
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC5 This week's digest",
                    TemplateText = "Summarize what came into my Inbox over the past 7 days. Group by sender or topic. Highlight urgent items and emails I'm directly addressed in.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCAC Conversation summary",
                    TemplateText = "Summarize my recent email conversations with [name or email]. Show the chronological flow and key decisions/topics.",
                },
                new ReportQuickActionChip {
                    Label = "\u2713 Action items",
                    TemplateText = "Find action items I need to do based on emails from the past 7 days. Read the relevant messages, extract TODOs/deadlines/asks. Group by who's waiting on what.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC1 Project status",
                    TemplateText = "Summarize the status of [topic/project name]. Find relevant emails, read the most recent ones, and give me: latest update, open questions, action items, key participants.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCCA Email stats",
                    TemplateText = "Give me email statistics for the past 30 days: top 10 senders, busiest days, breakdown by folder. Use outlook_aggregate_messages.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83C\uDFD6\uFE0F While I was out",
                    TemplateText = "I was out from [start date] to [end date]. Show me what's important from that timeframe: urgent items, direct asks, replies needed. De-prioritize newsletters and automated mail.",
                },
            };
        }
    }
}
