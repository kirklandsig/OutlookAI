using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// All Outlook OOM access used by Phase 2 tools flows through this interface.
    /// Production implementation (LiveOutlookSurface) marshals every call onto
    /// the Outlook STA UI thread via OutlookThreadMarshaller. Tests stub the
    /// methods they need without touching COM.
    /// </summary>
    public interface IOutlookSurface
    {
        ComposeStateResult GetCurrentComposeState(bool includeFullBody);
        IReadOnlyList<FolderResult> ListFolders();
        IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args);
        MessageDetail ReadMessage(string messageId, bool includeFullBody);
        int CountMessages(SearchMessagesArgs args);
        IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads);
        CreatedDraft CreateDraft(CreateDraftArgs args);
        void MarkAsRead(string messageId, bool read);
        void FlagMessage(string messageId, string flag);
        void SetCategory(string messageId, string category);
    }

    public sealed class ComposeStateResult
    {
        public string Subject { get; set; }
        public IReadOnlyList<string> ToRecipients { get; set; }
        public IReadOnlyList<string> CcRecipients { get; set; }
        public IReadOnlyList<string> BccRecipients { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public InReplyTo InReplyTo { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
    }

    public sealed class InReplyTo
    {
        public string ThreadTopic { get; set; }
        public IReadOnlyList<ThreadMessage> LastNMessages { get; set; }
    }

    public sealed class ThreadMessage
    {
        public string From { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
    }

    public sealed class AttachmentSummary
    {
        public string Filename { get; set; }
        public long SizeBytes { get; set; }
    }

    public sealed class FolderResult
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public int ItemCount { get; set; }
    }

    public sealed class SearchMessagesArgs
    {
        public string Query { get; set; }
        public string From { get; set; }
        public string SubjectContains { get; set; }
        public string BodyContains { get; set; }
        public bool? HasAttachment { get; set; }
        public bool? IsUnread { get; set; }
        public bool? IsFlagged { get; set; }
        /// <summary>One of "low" | "normal" | "high"; null = unset.</summary>
        public string Importance { get; set; }
        public string FolderId { get; set; }
        public DateTimeOffset? DateFrom { get; set; }
        public DateTimeOffset? DateTo { get; set; }
        public int MaxResults { get; set; } = 25;
    }

    public sealed class MessageSummary
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
        public bool HasAttachments { get; set; }
    }

    public sealed class MessageDetail
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
        public string InReplyToMessageId { get; set; }
        public string ConversationTopic { get; set; }
    }

    public sealed class ThreadSummary
    {
        public string ThreadTopic { get; set; }
        public DateTimeOffset LastMessageAt { get; set; }
        public int MessageCount { get; set; }
        public string Snippet { get; set; }
        public string ThreadId { get; set; }
    }

    public sealed class CreateDraftArgs
    {
        public string Subject { get; set; }
        public string BodyPlaintext { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public string InReplyToMessageId { get; set; }
    }

    public sealed class CreatedDraft
    {
        public string DraftId { get; set; }
        public string Location { get; set; }
    }
}
