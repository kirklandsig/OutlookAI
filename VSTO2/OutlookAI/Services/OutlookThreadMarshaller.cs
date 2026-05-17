using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services
{
    /// <summary>
    /// Marshals callbacks onto a specific SynchronizationContext (the Outlook
    /// UI thread in production). Used by LiveOutlookSurface to make every COM
    /// access STA-correct regardless of which thread the chat service runs on.
    /// </summary>
    public sealed class OutlookThreadMarshaller
    {
        private readonly SynchronizationContext _context;

        public OutlookThreadMarshaller(SynchronizationContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public Task RunAsync(Action action, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Reentrancy guard: when the caller is already on the target sync
            // context (e.g. ChatController invoking surface methods from a
            // WebMessageReceived handler that fires on the UI thread), Post +
            // blocking-wait would deadlock. Invoke synchronously in that case.
            if (IsOnTargetContext())
            {
                TraceLog.Write("RunAsync reentrant branch (sync invoke)", "Marshaller");
                var tcsSync = new TaskCompletionSource<bool>();
                try { action(); tcsSync.TrySetResult(true); }
                catch (Exception ex) { tcsSync.TrySetException(ex); }
                return tcsSync.Task;
            }
            TraceLog.Write("RunAsync posting (off-thread caller)", "Marshaller");
            var tcs = new TaskCompletionSource<bool>();
            _context.Post(_ =>
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { action(); tcs.TrySetResult(true); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }

        public Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // See reentrancy comment in the non-generic overload above.
            if (IsOnTargetContext())
            {
                TraceLog.Write("RunAsync<T> reentrant branch (sync invoke)", "Marshaller");
                var tcsSync = new TaskCompletionSource<T>();
                try { tcsSync.TrySetResult(func()); }
                catch (Exception ex) { tcsSync.TrySetException(ex); }
                return tcsSync.Task;
            }
            TraceLog.Write("RunAsync<T> posting (off-thread caller)", "Marshaller");
            var tcs = new TaskCompletionSource<T>();
            _context.Post(_ =>
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }
                try { tcs.TrySetResult(func()); }
                catch (Exception ex) { tcs.TrySetException(ex); }
            }, null);
            return tcs.Task;
        }

        private bool IsOnTargetContext()
        {
            return SynchronizationContext.Current != null
                && ReferenceEquals(SynchronizationContext.Current, _context);
        }
    }
}
