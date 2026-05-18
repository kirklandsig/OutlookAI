using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Production <see cref="IAdvancedSearchHost"/>. Wraps
    /// <c>Application.AdvancedSearch</c> and the
    /// <c>AdvancedSearchComplete</c> event, projecting each completed
    /// Outlook <c>Search.Results</c> into <see cref="MessageProjectionInput"/>
    /// items on the Outlook UI thread (where COM access is legal). Every
    /// COM call goes through <see cref="OutlookThreadMarshaller"/> so the
    /// host can be called from any thread.
    /// </summary>
    public sealed class LiveAdvancedSearchHost : IAdvancedSearchHost, IDisposable
    {
        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;
        private readonly IdResolver _ids;
        private readonly ConcurrentDictionary<string, Outlook.Search> _tagToSearch
            = new ConcurrentDictionary<string, Outlook.Search>();
        private bool _subscribed;
        private bool _disposed;

        // Snippet character cap mirrors LiveOutlookSurface.SnippetChars.
        private const int SnippetChars = 160;
        // Body byte cap mirrors LiveOutlookSurface.MaxBodyChars.
        private const int MaxBodyChars = 32 * 1024;

        public event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;

        public LiveAdvancedSearchHost(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            IdResolver ids)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (marshaller == null) throw new ArgumentNullException("marshaller");
            if (ids == null) throw new ArgumentNullException("ids");
            _application = application;
            _marshaller = marshaller;
            _ids = ids;
            Subscribe();
        }

        private void Subscribe()
        {
            _marshaller.RunAsync(() =>
            {
                if (_subscribed) return;
                _application.AdvancedSearchComplete += OnAdvancedSearchComplete;
                _subscribed = true;
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        public void Start(string scope, string filter, bool searchSubFolders, string tag)
        {
            _marshaller.RunAsync(() =>
            {
                try
                {
                    TraceLog.Write("AdvancedSearch Start tag=" + tag
                        + " scope_len=" + (scope == null ? 0 : scope.Length)
                        + " filter=" + (string.IsNullOrEmpty(filter) ? "<none>" : filter)
                        + " sub=" + searchSubFolders, "LiveHost");
                }
                catch { }
                var search = _application.AdvancedSearch(
                    scope,
                    filter ?? "",
                    searchSubFolders,
                    tag);
                if (search != null) _tagToSearch[tag] = search;
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        public void Stop(string tag)
        {
            _marshaller.RunAsync(() =>
            {
                Outlook.Search search;
                if (_tagToSearch.TryRemove(tag, out search))
                {
                    try { search.Stop(); }
                    catch (COMException ex)
                    {
                        try { TraceLog.Write("AdvancedSearch Stop COMException tag=" + tag + " " + ex.Message, "LiveHost"); } catch { }
                    }
                }
            }, System.Threading.CancellationToken.None)
            .GetAwaiter().GetResult();
        }

        private void OnAdvancedSearchComplete(Outlook.Search search)
        {
            if (search == null) return;
            var tag = "";
            try { tag = search.Tag ?? ""; } catch (COMException) { }
            Outlook.Search _;
            _tagToSearch.TryRemove(tag, out _);

            var items = new List<MessageProjectionInput>();
            try
            {
                foreach (var obj in search.Results)
                {
                    try
                    {
                        var mi = obj as Outlook.MailItem;
                        if (mi == null) continue;
                        items.Add(BuildProjectionInput(mi));
                    }
                    catch (COMException ex)
                    {
                        try { TraceLog.Write("AdvancedSearch result item COMException: " + ex.Message, "LiveHost"); } catch { }
                    }
                }
            }
            catch (COMException ex)
            {
                try { TraceLog.Write("AdvancedSearch results enumeration COMException: " + ex.Message, "LiveHost"); } catch { }
            }

            try { TraceLog.Write("AdvancedSearch Complete tag=" + tag + " raw_count=" + items.Count, "LiveHost"); } catch { }

            var handler = Completed;
            if (handler != null)
            {
                try
                {
                    handler(this, new AdvancedSearchHostCompleteEventArgs { Tag = tag, Items = items });
                }
                catch (Exception ex)
                {
                    try { TraceLog.Write("Completed handler threw: " + ex.Message, "LiveHost"); } catch { }
                }
            }
        }

        private MessageProjectionInput BuildProjectionInput(Outlook.MailItem mi)
        {
            string folderName = "";
            bool folderIsMail = true;
            try
            {
                var parent = mi.Parent as Outlook.MAPIFolder;
                if (parent != null)
                {
                    try { folderName = parent.Name ?? ""; } catch (COMException) { }
                    try { folderIsMail = parent.DefaultItemType == Outlook.OlItemType.olMailItem; }
                    catch (COMException) { folderIsMail = true; }
                }
            }
            catch (COMException) { /* leave defaults */ }

            string id = ""; try { id = _ids.Shorten(mi.EntryID); } catch (COMException) { }
            string subject = ""; try { subject = mi.Subject ?? ""; } catch (COMException) { }
            string from = ""; try { from = mi.SenderName ?? mi.SenderEmailAddress ?? ""; } catch (COMException) { }
            string to = ""; try { to = mi.To ?? ""; } catch (COMException) { }
            DateTimeOffset receivedAt = DateTimeOffset.MinValue;
            try { receivedAt = ToOffset(mi.ReceivedTime); } catch (COMException) { }
            bool hasAttachments = false;
            try { hasAttachments = (mi.Attachments != null && mi.Attachments.Count > 0); } catch (COMException) { }

            // SnippetFactory closes over the MailItem. Caller (projector via
            // marshaller.RunAsync block in LiveOutlookSurface) invokes it on
            // the UI thread.
            var capturedMi = mi;
            Func<string> snippetFactory = () =>
            {
                try
                {
                    var body = capturedMi.Body ?? "";
                    return SnippetOf(body);
                }
                catch (COMException) { return ""; }
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
            if (body.Length > MaxBodyChars) body = body.Substring(0, MaxBodyChars);
            var collapsed = body.Replace("\r\n", " ").Replace('\n', ' ').Trim();
            return collapsed.Length <= SnippetChars
                ? collapsed
                : collapsed.Substring(0, SnippetChars) + "...";
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                _marshaller.RunAsync(() =>
                {
                    if (!_subscribed) return;
                    try { _application.AdvancedSearchComplete -= OnAdvancedSearchComplete; } catch { }
                    _subscribed = false;
                }, System.Threading.CancellationToken.None)
                .GetAwaiter().GetResult();
            }
            catch { /* shutdown best-effort */ }
        }
    }
}
