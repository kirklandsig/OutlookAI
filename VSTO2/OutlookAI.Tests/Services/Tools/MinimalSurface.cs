using System.Collections.Generic;
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
        public virtual IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args)
            => throw new System.NotImplementedException();
        public virtual MessageDetail ReadMessage(string messageId, bool includeFullBody)
            => throw new System.NotImplementedException();
        public virtual int CountMessages(SearchMessagesArgs args)
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
    }
}
