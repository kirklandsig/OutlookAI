using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// A pre-canned prompt rendered as a clickable chip above the Inbox
    /// Copilot composer. Clicking a chip pre-fills the textarea and
    /// auto-sends (per Phase 3a spec). The set is computed server-side
    /// based on how many messages the user has selected in the active
    /// Explorer.
    /// </summary>
    public sealed class QuickActionChip
    {
        public string Label { get; set; }    // shown on the button
        public string Prompt { get; set; }   // pre-filled into the textarea

        /// <summary>
        /// Build the default chip set for a given selection count.
        /// Three static chips plus 0 or 2 dynamic chips:
        ///   0 selected -> static only
        ///   1 selected -> static + "Summarize this thread" + "Draft a reply"
        ///   2+ selected -> static + "Summarize all selected" + "Triage selected"
        /// </summary>
        public static IReadOnlyList<QuickActionChip> ComputeChipsForSelectionCount(int selectionCount)
        {
            var list = new List<QuickActionChip>
            {
                new QuickActionChip
                {
                    Label = "What needs my attention?",
                    Prompt = "Look at my inbox and tell me what needs attention. Prioritize by recency, importance, and sender. Be concise.",
                },
                new QuickActionChip
                {
                    Label = "Summarize unread",
                    Prompt = "Summarize all my unread messages. Group by sender or topic. Be concise.",
                },
                new QuickActionChip
                {
                    Label = "Today's emails",
                    Prompt = "Show me everything I received today, grouped by sender. Highlight anything that looks urgent.",
                },
            };

            if (selectionCount == 1)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Summarize this thread",
                    Prompt = "Summarize the selected message and the rest of its conversation thread.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Draft a reply",
                    Prompt = "Draft a reply to the selected message. Match the tone of the sender.",
                });
            }
            else if (selectionCount >= 2)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Summarize all selected",
                    Prompt = "Summarize all the selected messages.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Triage selected",
                    Prompt = "Triage the selected messages -- which need action, which can be archived, which can be marked read?",
                });
            }

            return list;
        }
    }
}
