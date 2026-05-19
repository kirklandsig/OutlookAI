using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Result of one AdvancedSearch invocation. Exactly one of Completed /
    /// TimedOut / Cancelled / Error is meaningful; the others are false /
    /// null. Items is non-null only on Completed = true.
    /// </summary>
    public sealed class AdvancedSearchResult
    {
        public bool Completed { get; set; }
        public bool TimedOut { get; set; }
        public bool Cancelled { get; set; }
        public Exception Error { get; set; }
        public IReadOnlyList<MessageProjectionInput> Items { get; set; }
    }

    public interface IOutlookAdvancedSearchRunner : IDisposable
    {
        Task<AdvancedSearchResult> RunAsync(
            string scope,
            string filter,
            bool searchSubFolders,
            TimeSpan timeout,
            CancellationToken ct);
    }

    /// <summary>
    /// Drives an <see cref="IAdvancedSearchHost"/> with timeout, cancellation,
    /// tag-keyed dispatch, and process-wide serialisation (one in-flight
    /// AdvancedSearch at a time, which is the documented safe ceiling for
    /// the Outlook OOM).
    /// </summary>
    public sealed class OutlookAdvancedSearchRunner : IOutlookAdvancedSearchRunner
    {
        private readonly IAdvancedSearchHost _host;
        private readonly SemaphoreSlim _serialiser = new SemaphoreSlim(1, 1);
        private readonly ConcurrentDictionary<string, TaskCompletionSource<AdvancedSearchResult>> _pending
            = new ConcurrentDictionary<string, TaskCompletionSource<AdvancedSearchResult>>();
        private bool _disposed;

        public OutlookAdvancedSearchRunner(IAdvancedSearchHost host)
        {
            if (host == null) throw new ArgumentNullException("host");
            _host = host;
            _host.Completed += OnHostCompleted;
        }

        public async Task<AdvancedSearchResult> RunAsync(
            string scope,
            string filter,
            bool searchSubFolders,
            TimeSpan timeout,
            CancellationToken ct)
        {
            if (_disposed) throw new ObjectDisposedException("OutlookAdvancedSearchRunner");

            await _serialiser.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var tag = Guid.NewGuid().ToString("N");
                var tcs = new TaskCompletionSource<AdvancedSearchResult>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                _pending[tag] = tcs;

                try { _host.Start(scope, filter, searchSubFolders, tag); }
                catch (Exception ex)
                {
                    TaskCompletionSource<AdvancedSearchResult> _;
                    _pending.TryRemove(tag, out _);
                    try { TraceLog.Write("AdvancedSearch Start threw: " + ex.Message, "Runner"); } catch { }
                    return new AdvancedSearchResult { Error = ex };
                }

                using (var timeoutCts = new CancellationTokenSource(timeout))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token))
                using (linked.Token.Register(() =>
                {
                    TaskCompletionSource<AdvancedSearchResult> pendingTcs;
                    if (!_pending.TryRemove(tag, out pendingTcs)) return;
                    try { _host.Stop(tag); } catch { /* best-effort */ }
                    if (ct.IsCancellationRequested)
                        pendingTcs.TrySetResult(new AdvancedSearchResult { Cancelled = true });
                    else
                        pendingTcs.TrySetResult(new AdvancedSearchResult { TimedOut = true });
                }))
                {
                    return await tcs.Task.ConfigureAwait(false);
                }
            }
            finally
            {
                _serialiser.Release();
            }
        }

        private void OnHostCompleted(object sender, AdvancedSearchHostCompleteEventArgs e)
        {
            if (e == null || e.Tag == null) return;
            TaskCompletionSource<AdvancedSearchResult> tcs;
            if (_pending.TryRemove(e.Tag, out tcs))
            {
                tcs.TrySetResult(new AdvancedSearchResult
                {
                    Completed = true,
                    Items = e.Items ?? new MessageProjectionInput[0],
                });
            }
            // Unknown tag: silently ignore. Some other AdvancedSearch (or a
            // stale event) is not our problem.
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { _host.Completed -= OnHostCompleted; } catch { }
            try { _serialiser.Dispose(); } catch { }
        }
    }
}
