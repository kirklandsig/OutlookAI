using System;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Arguments accepted by outlook_aggregate_messages. Mirrors
    /// SearchMessagesArgs for the filter portion, plus group_by + top_n.
    /// </summary>
    public sealed class AggregateMessagesArgs
    {
        public string Scope { get; set; } = "auto";    // current_folder | all_mail | auto
        public string FolderId { get; set; }            // optional explicit folder
        public DateTimeOffset? DateFrom { get; set; }
        public DateTimeOffset? DateTo { get; set; }
        public string From { get; set; }
        public string SubjectContains { get; set; }
        public string BodyContains { get; set; }
        public string GroupBy { get; set; } = "sender"; // sender | day | folder
        public int TopN { get; set; } = 10;
    }
}
