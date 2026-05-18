using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// DTO carrying the fields SearchResultProjector needs to build a
    /// MessageSummary. SnippetFactory lets the projector defer expensive
    /// body access until after sort + classifier-filter + top-N have run,
    /// so we only pay snippet cost for items we actually return.
    /// </summary>
    public sealed class MessageProjectionInput
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public bool HasAttachments { get; set; }
        public string FolderName { get; set; }
        public bool FolderDefaultItemTypeIsMail { get; set; }
        public Func<string> SnippetFactory { get; set; }
    }
}
