using System.Collections.Generic;
using System.Threading;
using OutlookAI.Services.Tools;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Base for per-tool test stubs. Each tool test overrides only the
    /// surface methods it needs; everything else throws NotImplementedException
    /// so a tool wandering into unrelated surface calls fails loudly.
    /// </summary>
    public abstract class MinimalSurface : IOutlookSurface
    {
        public virtual ComposeStateResult GetCurrentComposeState(bool includeFullBody)
            => throw new System.NotImplementedException();
        public virtual IReadOnlyList<FolderResult> ListFolders()
            => throw new System.NotImplementedException();
        public virtual IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
            => throw new System.NotImplementedException();
        public virtual MessageDetail ReadMessage(string messageId, bool includeFullBody)
            => throw new System.NotImplementedException();
        public virtual IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
            => throw new System.NotImplementedException();
        public virtual int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
            => throw new System.NotImplementedException();
        public virtual IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
            => throw new System.NotImplementedException();
        public virtual IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads)
            => throw new System.NotImplementedException();
        public virtual CreatedDraft CreateDraft(CreateDraftArgs args)
            => throw new System.NotImplementedException();
        public virtual void MarkAsRead(string messageId, bool read)
            => throw new System.NotImplementedException();
        public virtual void FlagMessage(string messageId, string flag)
            => throw new System.NotImplementedException();
        public virtual void SetCategory(string messageId, string category)
            => throw new System.NotImplementedException();
        public virtual CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
            => throw new System.NotImplementedException();
    }
}
