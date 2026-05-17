using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using OutlookAI.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Production <see cref="IOutlookSurface"/> implementation. Wraps Outlook
    /// OOM with explicit STA marshalling via <see cref="OutlookThreadMarshaller"/>.
    /// Every public method blocks the calling thread (which is typically the
    /// chat-service task thread) while the OOM call runs on the Outlook UI
    /// thread. Each method catches <see cref="COMException"/> and returns
    /// graceful defaults (null/empty) so the tool layer can surface a
    /// structured <c>{"error":...}</c> back to the model.
    /// </summary>
    public sealed class LiveOutlookSurface : IOutlookSurface
    {
        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;
        private readonly IdResolver _ids;
        private readonly Outlook.Inspector _composeInspector;
        // 32 KB hard cap on body bytes (in characters; close enough for ASCII).
        private const int MaxBodyChars = 32 * 1024;
        private const int MaxFolders = 200;
        private const int MaxFolderDepth = 6;
        private const int SnippetChars = 160;

        public LiveOutlookSurface(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            IdResolver ids,
            Outlook.Inspector composeInspector)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _ids = ids ?? throw new ArgumentNullException(nameof(ids));
            _composeInspector = composeInspector; // may be null for non-compose contexts
        }

        public ComposeStateResult GetCurrentComposeState(bool includeFullBody) =>
            Run(() =>
            {
                if (_composeInspector == null) return EmptyCompose();
                var item = _composeInspector.CurrentItem as Outlook.MailItem;
                if (item == null) return EmptyCompose();

                var body = item.Body ?? "";
                bool truncated = false;
                if (!includeFullBody && body.Length > 1000)
                {
                    body = body.Substring(0, 1000);
                    truncated = true;
                }
                else if (body.Length > MaxBodyChars)
                {
                    body = body.Substring(0, MaxBodyChars);
                    truncated = true;
                }

                var attachments = new List<AttachmentSummary>();
                try
                {
                    foreach (Outlook.Attachment att in item.Attachments)
                    {
                        attachments.Add(new AttachmentSummary
                        {
                            Filename = att.FileName,
                            SizeBytes = att.Size,
                        });
                    }
                }
                catch (COMException) { /* ignore */ }

                InReplyTo reply = null;
                try
                {
                    var conv = item.GetConversation();
                    if (conv != null)
                    {
                        var msgs = new List<ThreadMessage>();
                        var table = conv.GetTable();
                        while (!table.EndOfTable)
                        {
                            var row = table.GetNextRow();
                            var entry = row["EntryID"]?.ToString();
                            if (string.IsNullOrEmpty(entry)) continue;
                            var threadItem = _application.Session.GetItemFromID(entry) as Outlook.MailItem;
                            if (threadItem == null) continue;
                            msgs.Add(new ThreadMessage
                            {
                                From = threadItem.SenderName ?? threadItem.SenderEmailAddress ?? "",
                                ReceivedAt = ToOffset(threadItem.ReceivedTime),
                                Snippet = SnippetOf(threadItem.Body),
                            });
                            if (msgs.Count >= 5) break;
                        }
                        reply = new InReplyTo
                        {
                            ThreadTopic = item.ConversationTopic ?? "",
                            LastNMessages = msgs,
                        };
                    }
                }
                catch (COMException) { /* ignore */ }
                catch (Exception) { /* defensive */ }

                Outlook.AddressEntry sender = null;
                try { sender = _application.Session.CurrentUser?.AddressEntry; } catch (COMException) { }

                return new ComposeStateResult
                {
                    Subject = item.Subject ?? "",
                    ToRecipients = SplitAddresses(item.To),
                    CcRecipients = SplitAddresses(item.CC),
                    BccRecipients = SplitAddresses(item.BCC),
                    SenderName = sender?.Name ?? _application.Session.CurrentUser?.Name ?? "",
                    SenderEmail = TryGetSmtp(sender) ?? "",
                    BodyPlaintext = body,
                    BodyTruncated = truncated,
                    Attachments = attachments,
                    InReplyTo = reply,
                };
            }) ?? EmptyCompose();

        public IReadOnlyList<FolderResult> ListFolders() =>
            Run(() =>
            {
                var results = new List<FolderResult>();
                try
                {
                    foreach (Outlook.Store store in _application.Session.Stores)
                    {
                        WalkFolders(store.GetRootFolder(), null, results, depth: 0);
                        if (results.Count >= MaxFolders) break;
                    }
                }
                catch (COMException) { }
                return (IReadOnlyList<FolderResult>)results;
            }) ?? (IReadOnlyList<FolderResult>)new List<FolderResult>();

        public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args) =>
            Run(() =>
            {
                var summaries = new List<MessageSummary>();
                try
                {
                    var folder = ResolveFolder(args.FolderId);
                    if (folder == null) return (IReadOnlyList<MessageSummary>)summaries;

                    var filter = BuildRestrictFilter(args);
                    Outlook.Items items = string.IsNullOrEmpty(filter)
                        ? folder.Items
                        : folder.Items.Restrict(filter);
                    items.Sort("[ReceivedTime]", true);

                    int taken = 0;
                    foreach (var obj in items)
                    {
                        if (taken >= args.MaxResults) break;
                        var mi = obj as Outlook.MailItem;
                        if (mi == null) continue;
                        summaries.Add(new MessageSummary
                        {
                            Id = _ids.Shorten(mi.EntryID),
                            Subject = mi.Subject ?? "",
                            From = mi.SenderName ?? mi.SenderEmailAddress ?? "",
                            To = SplitAddresses(mi.To),
                            ReceivedAt = ToOffset(mi.ReceivedTime),
                            Snippet = SnippetOf(mi.Body),
                            HasAttachments = mi.Attachments?.Count > 0,
                        });
                        taken++;
                    }
                }
                catch (COMException) { }
                return (IReadOnlyList<MessageSummary>)summaries;
            }) ?? (IReadOnlyList<MessageSummary>)new List<MessageSummary>();

        public MessageDetail ReadMessage(string messageId, bool includeFullBody) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    var item = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                    if (item == null) return null;

                    var body = item.Body ?? "";
                    bool truncated = false;
                    if (body.Length > MaxBodyChars)
                    {
                        body = body.Substring(0, MaxBodyChars);
                        truncated = true;
                    }
                    if (!includeFullBody && body.Length > 1000)
                    {
                        body = body.Substring(0, 1000);
                        truncated = true;
                    }

                    var attachments = new List<AttachmentSummary>();
                    try
                    {
                        foreach (Outlook.Attachment att in item.Attachments)
                        {
                            attachments.Add(new AttachmentSummary
                            {
                                Filename = att.FileName,
                                SizeBytes = att.Size,
                            });
                        }
                    }
                    catch (COMException) { }

                    return new MessageDetail
                    {
                        Id = messageId,
                        Subject = item.Subject ?? "",
                        From = item.SenderName ?? item.SenderEmailAddress ?? "",
                        To = SplitAddresses(item.To),
                        Cc = SplitAddresses(item.CC),
                        ReceivedAt = ToOffset(item.ReceivedTime),
                        BodyPlaintext = body,
                        BodyTruncated = truncated,
                        Attachments = attachments,
                        InReplyToMessageId = null,
                        ConversationTopic = item.ConversationTopic ?? "",
                    };
                }
                catch (COMException) { return null; }
                catch (KeyNotFoundException) { return null; }
            });

        public int CountMessages(SearchMessagesArgs args) =>
            Run(() =>
            {
                try
                {
                    var folder = ResolveFolder(args.FolderId);
                    if (folder == null) return 0;
                    var filter = BuildRestrictFilter(args);
                    var items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
                    return items.Count;
                }
                catch (COMException) { return 0; }
            });

        public IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) =>
            Run(() =>
            {
                var results = new List<ThreadSummary>();
                if (string.IsNullOrEmpty(recipientEmail)) return (IReadOnlyList<ThreadSummary>)results;
                try
                {
                    var byConvId = new Dictionary<string, ThreadSummary>(StringComparer.Ordinal);
                    foreach (var defaultFolder in new[] { Outlook.OlDefaultFolders.olFolderInbox, Outlook.OlDefaultFolders.olFolderSentMail })
                    {
                        Outlook.MAPIFolder folder = null;
                        try { folder = _application.Session.GetDefaultFolder(defaultFolder); }
                        catch (COMException) { continue; }
                        if (folder == null) continue;

                        var filter = "@SQL=" +
                            "(urn:schemas:httpmail:fromemail LIKE '%" + Escape(recipientEmail) + "%' " +
                            "OR urn:schemas:httpmail:displayto LIKE '%" + Escape(recipientEmail) + "%' " +
                            "OR urn:schemas:httpmail:displaycc LIKE '%" + Escape(recipientEmail) + "%')";
                        Outlook.Items items;
                        try { items = folder.Items.Restrict(filter); }
                        catch (COMException) { continue; }
                        try { items.Sort("[ReceivedTime]", true); } catch (COMException) { }

                        int scanned = 0;
                        foreach (var obj in items)
                        {
                            if (scanned++ >= 200) break;
                            var mi = obj as Outlook.MailItem;
                            if (mi == null) continue;
                            var convId = mi.ConversationID ?? mi.ConversationTopic ?? mi.EntryID;
                            if (string.IsNullOrEmpty(convId)) continue;
                            if (!byConvId.TryGetValue(convId, out var s))
                            {
                                s = new ThreadSummary
                                {
                                    ThreadTopic = mi.ConversationTopic ?? "",
                                    LastMessageAt = ToOffset(mi.ReceivedTime),
                                    MessageCount = 0,
                                    Snippet = SnippetOf(mi.Body),
                                    ThreadId = _ids.Shorten(convId),
                                };
                                byConvId[convId] = s;
                            }
                            s.MessageCount++;
                            var ra = ToOffset(mi.ReceivedTime);
                            if (ra > s.LastMessageAt) s.LastMessageAt = ra;
                        }
                    }
                    results = byConvId.Values
                        .OrderByDescending(t => t.LastMessageAt)
                        .Take(maxThreads)
                        .ToList();
                }
                catch (COMException) { }
                return (IReadOnlyList<ThreadSummary>)results;
            }) ?? (IReadOnlyList<ThreadSummary>)new List<ThreadSummary>();

        public CreatedDraft CreateDraft(CreateDraftArgs args) =>
            Run(() =>
            {
                try
                {
                    Outlook.MailItem draft;
                    if (!string.IsNullOrEmpty(args.InReplyToMessageId))
                    {
                        var entryId = _ids.Resolve(args.InReplyToMessageId);
                        var original = _application.Session.GetItemFromID(entryId) as Outlook.MailItem;
                        draft = original?.Reply()
                                ?? (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);
                    }
                    else
                    {
                        draft = (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);
                    }

                    draft.Subject = args.Subject ?? "";
                    draft.Body = args.BodyPlaintext ?? "";
                    if (args.To != null && args.To.Count > 0)
                        draft.To = string.Join("; ", args.To);
                    if (args.Cc != null && args.Cc.Count > 0)
                        draft.CC = string.Join("; ", args.Cc);
                    draft.Save();
                    return new CreatedDraft
                    {
                        DraftId = _ids.Shorten(draft.EntryID),
                        Location = "Drafts",
                    };
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException("CreateDraft failed: " + ex.Message, ex);
                }
            });

        public void MarkAsRead(string messageId, bool read) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        item.UnRead = !read;
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        public void FlagMessage(string messageId, string flag) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        // FlagStatus property numeric values:
                        //   olNoFlag = 0, olFlagComplete = 1, olFlagMarked = 2.
                        // "todo" => olFlagMarked, "complete" => olFlagComplete, "none" => olNoFlag.
                        switch (flag)
                        {
                            case "todo": item.FlagStatus = Outlook.OlFlagStatus.olFlagMarked; break;
                            case "complete": item.FlagStatus = Outlook.OlFlagStatus.olFlagComplete; break;
                            default: item.FlagStatus = Outlook.OlFlagStatus.olNoFlag; break;
                        }
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        public void SetCategory(string messageId, string category) =>
            Run(() =>
            {
                try
                {
                    var entryId = _ids.Resolve(messageId);
                    if (_application.Session.GetItemFromID(entryId) is Outlook.MailItem item)
                    {
                        item.Categories = category ?? "";
                        item.Save();
                    }
                }
                catch (COMException) { }
                catch (KeyNotFoundException) { }
            });

        // ---------- helpers ----------

        private T Run<T>(Func<T> fn) => _marshaller.RunAsync(fn, CancellationToken.None).GetAwaiter().GetResult();

        private void Run(Action fn) => _marshaller.RunAsync(fn, CancellationToken.None).GetAwaiter().GetResult();

        private static ComposeStateResult EmptyCompose() => new ComposeStateResult
        {
            Subject = "",
            ToRecipients = new string[0],
            CcRecipients = new string[0],
            BccRecipients = new string[0],
            SenderName = "",
            SenderEmail = "",
            BodyPlaintext = "",
            BodyTruncated = false,
            Attachments = new AttachmentSummary[0],
        };

        private static IReadOnlyList<string> SplitAddresses(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return new string[0];
            return raw.Split(';')
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        private static DateTimeOffset ToOffset(DateTime dt)
        {
            try
            {
                if (dt.Kind == DateTimeKind.Utc) return new DateTimeOffset(dt, TimeSpan.Zero);
                if (dt.Kind == DateTimeKind.Local) return new DateTimeOffset(dt);
                return new DateTimeOffset(DateTime.SpecifyKind(dt, DateTimeKind.Local));
            }
            catch { return DateTimeOffset.MinValue; }
        }

        private static string SnippetOf(string body)
        {
            if (string.IsNullOrEmpty(body)) return "";
            var collapsed = body.Replace("\r\n", " ").Replace('\n', ' ').Trim();
            return collapsed.Length <= SnippetChars
                ? collapsed
                : collapsed.Substring(0, SnippetChars) + "...";
        }

        private static string TryGetSmtp(Outlook.AddressEntry entry)
        {
            if (entry == null) return null;
            try { return entry.GetExchangeUser()?.PrimarySmtpAddress ?? entry.Address; }
            catch (COMException) { return null; }
        }

        private Outlook.MAPIFolder ResolveFolder(string folderId)
        {
            if (string.IsNullOrEmpty(folderId))
            {
                try { return _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); }
                catch (COMException) { return null; }
            }
            try
            {
                var entryId = _ids.Resolve(folderId);
                return _application.Session.GetFolderFromID(entryId) as Outlook.MAPIFolder;
            }
            catch (COMException) { return null; }
            catch (KeyNotFoundException) { return null; }
        }

        private void WalkFolders(Outlook.MAPIFolder folder, string parentId, List<FolderResult> results, int depth)
        {
            if (folder == null) return;
            if (depth > MaxFolderDepth) return;
            if (results.Count >= MaxFolders) return;

            string id;
            try { id = _ids.Shorten(folder.EntryID); }
            catch (COMException) { return; }

            int count = 0;
            try { count = folder.Items.Count; } catch (COMException) { }

            results.Add(new FolderResult
            {
                Id = id,
                Name = folder.Name ?? "",
                ParentId = parentId,
                ItemCount = count,
            });

            try
            {
                foreach (Outlook.MAPIFolder child in folder.Folders)
                {
                    WalkFolders(child, id, results, depth + 1);
                    if (results.Count >= MaxFolders) break;
                }
            }
            catch (COMException) { }
        }

        internal static string BuildRestrictFilter(SearchMessagesArgs args)
        {
            if (args == null) return null;
            var clauses = new List<string>();

            if (!string.IsNullOrEmpty(args.Query))
            {
                var q = Escape(args.Query);
                clauses.Add("(urn:schemas:httpmail:subject LIKE '%" + q + "%' OR " +
                            "urn:schemas:httpmail:textdescription LIKE '%" + q + "%')");
            }
            if (!string.IsNullOrEmpty(args.From))
            {
                var v = Escape(args.From);
                clauses.Add("(urn:schemas:httpmail:fromname LIKE '%" + v + "%' OR " +
                            "urn:schemas:httpmail:fromemail LIKE '%" + v + "%')");
            }
            if (!string.IsNullOrEmpty(args.SubjectContains))
            {
                clauses.Add("urn:schemas:httpmail:subject LIKE '%" + Escape(args.SubjectContains) + "%'");
            }
            if (!string.IsNullOrEmpty(args.BodyContains))
            {
                clauses.Add("urn:schemas:httpmail:textdescription LIKE '%" + Escape(args.BodyContains) + "%'");
            }
            if (args.HasAttachment.HasValue)
            {
                clauses.Add("urn:schemas:httpmail:hasattachment = " + (args.HasAttachment.Value ? "1" : "0"));
            }
            if (args.IsUnread.HasValue)
            {
                // urn:schemas:httpmail:read is 0 for unread, 1 for read.
                clauses.Add("urn:schemas:httpmail:read = " + (args.IsUnread.Value ? "0" : "1"));
            }
            if (args.IsFlagged.HasValue)
            {
                // PR_FLAG_STATUS (0x1090) PT_LONG (0x0003) => 0x10900003.
                // Value 2 = followup flagged.
                clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" " +
                            (args.IsFlagged.Value ? "= 2" : "<> 2"));
            }
            if (!string.IsNullOrEmpty(args.Importance))
            {
                // PR_IMPORTANCE (0x0017) PT_LONG => 0x00170003. 0=low, 1=normal, 2=high.
                var imp = args.Importance.Trim().ToLowerInvariant();
                if (imp == "low" || imp == "normal" || imp == "high")
                {
                    var val = imp == "low" ? "0" : imp == "normal" ? "1" : "2";
                    clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x00170003\" = " + val);
                }
                // Unknown importance values are silently ignored.
            }
            if (args.DateFrom.HasValue)
            {
                clauses.Add("urn:schemas:httpmail:datereceived >= '" +
                    args.DateFrom.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
            }
            if (args.DateTo.HasValue)
            {
                clauses.Add("urn:schemas:httpmail:datereceived <= '" +
                    args.DateTo.Value.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + "'");
            }

            if (clauses.Count == 0) return null;
            return "@SQL=" + string.Join(" AND ", clauses);
        }

        private static string Escape(string s) => (s ?? "").Replace("'", "''");
    }
}
