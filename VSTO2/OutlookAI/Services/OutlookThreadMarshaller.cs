using System;
using System.Threading;
using System.Threading.Tasks;

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
    }
}
