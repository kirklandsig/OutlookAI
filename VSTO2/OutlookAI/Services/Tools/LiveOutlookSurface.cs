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
        private readonly Outlook.Explorer _explorer;     // Phase 3a
        private readonly IOutlookAdvancedSearchRunner _runner;
        private readonly IFolderClassifier _classifier;
        private static readonly TimeSpan _searchTimeout = TimeSpan.FromSeconds(30);
        // 32 KB hard cap on body bytes (in characters; close enough for ASCII).
        private const int MaxBodyChars = 32 * 1024;
        private const int MaxFolders = 200;
        private const int MaxFolderDepth = 6;
        private const int SnippetChars = 160;

        /// <summary>
        /// Phase 3a / 3b primary constructor. Runner is required and shared
        /// process-wide so all AdvancedSearch invocations go through one
        /// serialiser and one event subscription.
        /// </summary>
        public LiveOutlookSurface(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            IdResolver ids,
            Outlook.Inspector composeInspector,
            Outlook.Explorer explorer,
            IOutlookAdvancedSearchRunner runner,
            IFolderClassifier classifier = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _ids = ids ?? throw new ArgumentNullException(nameof(ids));
            _composeInspector = composeInspector;
            _explorer = explorer;
            _runner = runner ?? throw new ArgumentNullException(nameof(runner));
            _classifier = classifier ?? new FolderClassifier();
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

        public IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
        {
            args = args ?? new SearchMessagesArgs();
            var filter = BuildRestrictFilter(args);
            var scopeMode = (args.Scope ?? "auto").Trim().ToLowerInvariant();

            // Step 1: resolve scope on UI thread.
            SearchScope scope;
            try
            {
                scope = _marshaller.RunAsync(() => BuildSearchScope(args, scopeMode), ct).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            OutlookAI.Diagnostics.TraceLog.Write(
                "SearchMessages primary=AdvancedSearch start scope_paths=" + (scope.ResolvedFolderPaths?.Count ?? 0)
                + " sub=" + scope.SearchSubFolders + " filter=" + (string.IsNullOrEmpty(filter) ? "<none>" : filter),
                "LiveOutlookSurface");

            // Step 2: try AdvancedSearch.
            AdvancedSearchResult primary;
            try
            {
                primary = _runner.RunAsync(scope.ScopeString, filter, scope.SearchSubFolders, _searchTimeout, ct)
                    .GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            if (primary.Cancelled)
            {
                OutlookAI.Diagnostics.TraceLog.Write("SearchMessages primary cancelled by user", "LiveOutlookSurface");
                throw new OperationCanceledException(ct);
            }

            if (primary.Completed && primary.Items != null)
            {
                OutlookAI.Diagnostics.TraceLog.Write(
                    "SearchMessages primary=AdvancedSearch complete raw_count=" + primary.Items.Count,
                    "LiveOutlookSurface");
                return _marshaller.RunAsync(
                    () => SearchResultProjector.Project(primary.Items, args, _classifier),
                    ct).GetAwaiter().GetResult();
            }

            var reason = primary.TimedOut ? "timeout" :
                         primary.Error != null ? "com:" + primary.Error.GetType().Name :
                         "null_results";
            OutlookAI.Diagnostics.TraceLog.Write("SearchMessages fallback reason=" + reason, "LiveOutlookSurface");
            return FallbackIterativeSearch(args, filter, scopeMode, ct);
        }

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

        public int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
        {
            args = args ?? new SearchMessagesArgs();
            // Reuse the non-blocking SearchMessages path so counts honour the
            // same engine, skip list, and cancellation. MaxResults=int.MaxValue
            // makes the projector return every survivor; we just take the count.
            var widened = new SearchMessagesArgs
            {
                Query = args.Query,
                From = args.From,
                SubjectContains = args.SubjectContains,
                BodyContains = args.BodyContains,
                HasAttachment = args.HasAttachment,
                IsUnread = args.IsUnread,
                IsFlagged = args.IsFlagged,
                Importance = args.Importance,
                FolderId = args.FolderId,
                DateFrom = args.DateFrom,
                DateTo = args.DateTo,
                Scope = args.Scope,
                SortOrder = args.SortOrder,
                AttachmentFilter = args.AttachmentFilter,
                ReadStatus = args.ReadStatus,
                FlagStatus = args.FlagStatus,
                ImportanceFilter = args.ImportanceFilter,
                MaxResults = int.MaxValue,
            };
            var hits = SearchMessages(widened, ct);
            return hits?.Count ?? 0;
        }

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

        public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems) =>
            Run(() =>
            {
                if (_explorer == null)
                {
                    return new CurrentSelectionResult
                    {
                        Folder = "",
                        FolderId = "",
                        Count = 0,
                        Messages = new MessageDetail[0],
                    };
                }

                string folderName = "";
                string folderId = "";
                try
                {
                    var folder = _explorer.CurrentFolder;
                    if (folder != null)
                    {
                        folderName = folder.Name ?? "";
                        folderId = _ids.Shorten(folder.EntryID ?? "");
                    }
                }
                catch (COMException) { }

                var selection = _explorer.Selection;
                int totalCount = 0;
                try { totalCount = selection.Count; } catch (COMException) { }

                var picked = new List<MessageDetail>();
                try
                {
                    int taken = 0;
                    // Outlook Selection is 1-based.
                    for (int i = 1; i <= totalCount && taken < maxItems; i++)
                    {
                        object item = null;
                        try { item = selection[i]; } catch (COMException) { continue; }
                        var mi = item as Outlook.MailItem;
                        if (mi == null) continue;

                        var body = mi.Body ?? "";
                        bool truncated = false;
                        if (!includeFullBodies && body.Length > 1000)
                        {
                            body = body.Substring(0, 1000);
                            truncated = true;
                        }
                        else if (body.Length > MaxBodyChars)
                        {
                            body = body.Substring(0, MaxBodyChars);
                            truncated = true;
                        }

                        var atts = new List<AttachmentSummary>();
                        try
                        {
                            foreach (Outlook.Attachment att in mi.Attachments)
                            {
                                atts.Add(new AttachmentSummary
                                {
                                    Filename = att.FileName,
                                    SizeBytes = att.Size,
                                });
                            }
                        }
                        catch (COMException) { }

                        picked.Add(new MessageDetail
                        {
                            Id = _ids.Shorten(mi.EntryID ?? ""),
                            Subject = mi.Subject ?? "",
                            From = (mi.SenderName ?? "") +
                                   (string.IsNullOrEmpty(mi.SenderEmailAddress) ? "" :
                                    " <" + mi.SenderEmailAddress + ">"),
                            To = SplitAddresses(mi.To),
                            Cc = SplitAddresses(mi.CC),
                            ReceivedAt = ToOffset(mi.ReceivedTime),
                            BodyPlaintext = body,
                            BodyTruncated = truncated,
                            Attachments = atts,
                            InReplyToMessageId = null,
                            ConversationTopic = mi.ConversationTopic ?? "",
                        });
                        taken++;
                    }
                }
                catch (COMException) { }

                return new CurrentSelectionResult
                {
                    Folder = folderName,
                    FolderId = folderId,
                    Count = totalCount,
                    Messages = picked,
                };
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

        // Phase 3b primary scope resolver. Runs on the UI thread via the
        // marshaller in SearchMessages. Builds a SearchScope (DASL scope
        // string + SearchSubFolders flag) for Application.AdvancedSearch.
        private SearchScope BuildSearchScope(SearchMessagesArgs args, string scopeMode)
        {
            var paths = new List<string>();
            bool searchSubFolders = true;

            try
            {
                if (!string.IsNullOrEmpty(args.FolderId))
                {
                    var folder = ResolveFolder(args.FolderId);
                    if (folder != null)
                    {
                        try { paths.Add(folder.FolderPath); } catch (COMException) { }
                    }
                }
                else if (scopeMode == "current_folder")
                {
                    var folder = ResolveCurrentFolder();
                    if (folder != null)
                    {
                        try { paths.Add(folder.FolderPath); } catch (COMException) { }
                    }
                    searchSubFolders = false;
                }
                else
                {
                    foreach (Outlook.Store store in _application.Session.Stores)
                    {
                        try
                        {
                            var root = store.GetRootFolder();
                            if (root != null)
                            {
                                try { paths.Add(root.FolderPath); } catch (COMException) { }
                            }
                        }
                        catch (COMException) { }
                    }
                }
            }
            catch (COMException ex)
            {
                try { OutlookAI.Diagnostics.TraceLog.Write("BuildSearchScope COMException: " + ex.Message, "LiveOutlookSurface"); } catch { }
            }

            return new SearchScope
            {
                ScopeString = SearchScopeFormatter.Format(paths),
                SearchSubFolders = searchSubFolders,
                ResolvedFolderPaths = paths,
            };
        }

        // Phase 3b yielding fallback. Walks each folder via one
        // marshaller.RunAsync call so the UI thread is released between
        // folders (Outlook can pump messages between marshalled calls).
        private IReadOnlyList<MessageSummary> FallbackIterativeSearch(
            SearchMessagesArgs args, string filter, string scopeMode, CancellationToken ct)
        {
            var allInputs = new List<MessageProjectionInput>();
            var searchAllMail = scopeMode == "all_mail" || scopeMode == "auto";

            IReadOnlyList<Outlook.MAPIFolder> folders;
            try
            {
                folders = _marshaller.RunAsync(
                    () => ResolveSearchFolders(args, allMail: searchAllMail),
                    ct).GetAwaiter().GetResult();
            }
            catch (OperationCanceledException) { throw; }

            var searched = 0;
            foreach (var folder in folders)
            {
                ct.ThrowIfCancellationRequested();
                List<MessageProjectionInput> folderInputs;
                try
                {
                    folderInputs = _marshaller.RunAsync(
                        () => CollectFolderInputs(folder, args, filter),
                        ct).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) { throw; }
                allInputs.AddRange(folderInputs);
                searched++;
                try
                {
                    OutlookAI.Diagnostics.TraceLog.Write(
                        "SearchMessages fallback folder_done=" + SafeFolderName(folder)
                        + " taken=" + folderInputs.Count + " searched=" + searched,
                        "LiveOutlookSurface");
                }
                catch { }
            }

            return _marshaller.RunAsync(
                () => SearchResultProjector.Project(allInputs, args, _classifier),
                ct).GetAwaiter().GetResult();
        }

        // Collect one folder's worth of MessageProjectionInput WITHOUT body
        // access. Snippet is deferred to the projector via SnippetFactory.
        // CRITICAL: sorts by [ReceivedTime] using Outlook's index, then
        // enumerates AT MOST SearchFallbackBudget.PerFolderItems items.
        // Without this cap a single 588-item folder enumerated ~63 seconds
        // on the UI thread (~107 ms per COM property access). With the cap
        // each folder costs at most a few hundred ms, so the UI thread is
        // released between folders fast enough that Outlook stays usable.
        private List<MessageProjectionInput> CollectFolderInputs(
            Outlook.MAPIFolder folder, SearchMessagesArgs args, string filter)
        {
            var inputs = new List<MessageProjectionInput>();
            if (folder == null) return inputs;

            string folderName = "";
            bool folderIsMail = true;
            try { folderName = folder.Name ?? ""; } catch (COMException) { }
            try { folderIsMail = folder.DefaultItemType == Outlook.OlItemType.olMailItem; } catch (COMException) { }
            if (_classifier.IsSystemFolder(folderName, folderIsMail)) return inputs;

            Outlook.Items items;
            try
            {
                items = string.IsNullOrEmpty(filter) ? folder.Items : folder.Items.Restrict(filter);
            }
            catch (COMException ex)
            {
                try { OutlookAI.Diagnostics.TraceLog.Write("CollectFolderInputs Restrict COMException folder=" + folderName + ": " + ex.Message, "LiveOutlookSurface"); } catch { }
                return inputs;
            }

            // Pre-sort using Outlook's index so the items we visit first
            // are the items most likely to survive the global sort + clamp.
            try { items.Sort("[ReceivedTime]", SearchFallbackBudget.DescendingForSortOrder(args.SortOrder)); }
            catch (COMException) { /* sort is best-effort; index may be unavailable on some stores */ }

            var limit = SearchFallbackBudget.PerFolderItems(args);
            int taken = 0;
            foreach (var obj in items)
            {
                if (taken >= limit) break;
                try
                {
                    var mi = obj as Outlook.MailItem;
                    if (mi == null) continue;
                    inputs.Add(BuildFallbackInput(mi, folderName, folderIsMail));
                    taken++;
                }
                catch (COMException) { /* skip the item */ }
            }
            return inputs;
        }

        private MessageProjectionInput BuildFallbackInput(
            Outlook.MailItem mi, string folderName, bool folderIsMail)
        {
            string id = ""; try { id = _ids.Shorten(mi.EntryID); } catch (COMException) { }
            string subject = ""; try { subject = mi.Subject ?? ""; } catch (COMException) { }
            string from = ""; try { from = mi.SenderName ?? mi.SenderEmailAddress ?? ""; } catch (COMException) { }
            string to = ""; try { to = mi.To ?? ""; } catch (COMException) { }
            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try { receivedAt = ToOffset(mi.ReceivedTime); } catch (COMException) { }
            bool hasAttachments = false;
            try { hasAttachments = mi.Attachments != null && mi.Attachments.Count > 0; } catch (COMException) { }

            var capturedMi = mi;
            Func<string> snippetFactory = () =>
            {
                try { return SnippetOf(capturedMi.Body); } catch (COMException) { return ""; }
            };

            return new MessageProjectionInput
            {
                Id = id,
                Subject = subject,
                From = from,
                To = SplitAddresses(to),
                ReceivedAt = receivedAt,
                HasAttachments = hasAttachments,
                FolderName = folderName,
                FolderDefaultItemTypeIsMail = folderIsMail,
                SnippetFactory = snippetFactory,
            };
        }

        private Outlook.MAPIFolder ResolveCurrentFolder()
        {
            try
            {
                var folder = _explorer?.CurrentFolder as Outlook.MAPIFolder;
                if (folder != null) return folder;
            }
            catch (COMException) { }

            try { return _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox); }
            catch (COMException) { return null; }
        }

        private IReadOnlyList<Outlook.MAPIFolder> ResolveSearchFolders(SearchMessagesArgs args, bool allMail)
        {
            var folders = new List<Outlook.MAPIFolder>();
            if (!string.IsNullOrEmpty(args?.FolderId))
            {
                var folder = ResolveFolder(args.FolderId);
                if (folder != null) folders.Add(folder);
                return folders;
            }

            if (!allMail)
            {
                var folder = ResolveCurrentFolder();
                if (folder != null) folders.Add(folder);
                return folders;
            }

            try
            {
                foreach (Outlook.Store store in _application.Session.Stores)
                {
                    WalkMailFolders(store.GetRootFolder(), folders, depth: 0);
                }
            }
            catch (COMException) { }
            return folders;
        }

        private void WalkMailFolders(Outlook.MAPIFolder folder, List<Outlook.MAPIFolder> results, int depth)
        {
            if (folder == null || depth > MaxFolderDepth) return;
            var name = SafeFolderName(folder);
            bool isMailFolder = false;
            try { isMailFolder = folder.DefaultItemType == Outlook.OlItemType.olMailItem; }
            catch (COMException) { }
            if (!_classifier.IsSystemFolder(name, isMailFolder))
            {
                if (isMailFolder) results.Add(folder);
            }

            try
            {
                foreach (Outlook.MAPIFolder child in folder.Folders)
                {
                    WalkMailFolders(child, results, depth + 1);
                }
            }
            catch (COMException) { }
        }

        private static string SafeFolderName(Outlook.MAPIFolder folder)
        {
            try { return folder?.Name ?? ""; }
            catch (COMException) { return ""; }
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
            var attachmentFilter = (args.AttachmentFilter ?? "any").Trim().ToLowerInvariant();
            if (attachmentFilter == "with" || args.HasAttachment == true)
            {
                clauses.Add("urn:schemas:httpmail:hasattachment = 1");
            }
            else if (attachmentFilter == "without")
            {
                clauses.Add("urn:schemas:httpmail:hasattachment = 0");
            }

            var readStatus = (args.ReadStatus ?? "any").Trim().ToLowerInvariant();
            if (readStatus == "unread" || args.IsUnread == true)
            {
                clauses.Add("urn:schemas:httpmail:read = 0");
            }
            else if (readStatus == "read")
            {
                clauses.Add("urn:schemas:httpmail:read = 1");
            }

            var flagStatus = (args.FlagStatus ?? "any").Trim().ToLowerInvariant();
            if (flagStatus == "flagged" || args.IsFlagged == true)
            {
                clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" = 2");
            }
            else if (flagStatus == "unflagged")
            {
                clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" <> 2");
            }

            var importanceFilter = (args.ImportanceFilter ?? "any").Trim().ToLowerInvariant();
            if (importanceFilter == "any" && !string.IsNullOrEmpty(args.Importance))
            {
                importanceFilter = args.Importance.Trim().ToLowerInvariant();
            }
            if (importanceFilter == "low" || importanceFilter == "normal" || importanceFilter == "high")
            {
                // PR_IMPORTANCE (0x0017) PT_LONG => 0x00170003. 0=low, 1=normal, 2=high.
                var val = importanceFilter == "low" ? "0" : importanceFilter == "normal" ? "1" : "2";
                clauses.Add("\"http://schemas.microsoft.com/mapi/proptag/0x00170003\" = " + val);
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
