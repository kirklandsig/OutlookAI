using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class OutlookThreadMarshallerTests
    {
        [Fact]
        public async Task RunAsync_InvokesOnConfiguredContext()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                int observed = -1;
                await marshaller.RunAsync(
                    () => { observed = Thread.CurrentThread.ManagedThreadId; },
                    CancellationToken.None);
                Assert.Equal(ctx.WorkerThreadId, observed);
                Assert.NotEqual(Thread.CurrentThread.ManagedThreadId, observed);
            }
        }

        [Fact]
        public async Task RunAsyncGeneric_ReturnsValueFromContextThread()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                var value = await marshaller.RunAsync(
                    () => 42 + Thread.CurrentThread.ManagedThreadId,
                    CancellationToken.None);
                Assert.Equal(42 + ctx.WorkerThreadId, value);
            }
        }

        [Fact]
        public async Task RunAsync_PropagatesException()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                var ex = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    marshaller.RunAsync(
                        () => throw new InvalidOperationException("boom"),
                        CancellationToken.None));
                Assert.Equal("boom", ex.Message);
            }
        }

        [Fact]
        public async Task RunAsync_RespectsCancellationBeforeDispatch()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                var cts = new CancellationTokenSource();
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    marshaller.RunAsync(() => { }, cts.Token));
            }
        }

        /// <summary>
        /// Regression test for the Phase 2 UI freeze: when the caller is
        /// already on the target SynchronizationContext (e.g.
        /// <c>ChatController.PushContextStripFromSurface</c> firing from
        /// <c>WebMessageReceived</c> on the UI thread), the marshaller must
        /// invoke the action synchronously on the current thread instead of
        /// posting + blocking. The old behaviour deadlocked because
        /// <c>.GetAwaiter().GetResult()</c> blocked the only thread that
        /// could run the posted continuation.
        /// </summary>
        [Fact]
        public void RunAsync_SameContextCaller_InvokesSynchronously_NoDeadlock()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                int? observedThreadId = null;

                // Hop onto the context's worker thread, then call RunAsync.
                // Inside the action we capture the thread id and verify that
                // we never had to post-and-wait. We use ctx.Send to get onto
                // the worker thread without async machinery.
                ctx.Send(_ =>
                {
                    var workerId = Thread.CurrentThread.ManagedThreadId;
                    // This call would hang the worker thread forever before
                    // the fix.
                    marshaller.RunAsync(
                        () => { observedThreadId = Thread.CurrentThread.ManagedThreadId; },
                        CancellationToken.None).GetAwaiter().GetResult();
                    Assert.Equal(workerId, observedThreadId);
                }, null);

                Assert.True(observedThreadId.HasValue);
                Assert.Equal(ctx.WorkerThreadId, observedThreadId.Value);
            }
        }

        /// <summary>
        /// Regression for today's freeze. VSTO/Outlook can swap the live
        /// <see cref="SynchronizationContext"/> on the UI thread between
        /// startup (when our marshaller captured one) and first use. The
        /// marshaller must still recognise it's on the UI thread - because
        /// the managed thread id is stable - and short-circuit instead of
        /// trying to post to the now-stale captured instance.
        /// </summary>
        [Fact]
        public void RunAsync_StaleCapturedContext_StillShortCircuits_WhenOnUiThread()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                int observed = -1;

                ctx.Send(_ =>
                {
                    // Simulate VSTO swapping the SyncContext under us: install
                    // a different SyncContext instance as Current on the same
                    // thread that the marshaller targets.
                    var swapped = new SimpleSyncContext();
                    SynchronizationContext.SetSynchronizationContext(swapped);

                    // Pre-fix: hangs because IsOnTargetContext compared by
                    // SyncContext reference equality and swapped != ctx.
                    // Post-fix: thread id comparison still matches, so we
                    // short-circuit and run synchronously.
                    marshaller.RunAsync(
                        () => { observed = Thread.CurrentThread.ManagedThreadId; },
                        CancellationToken.None).GetAwaiter().GetResult();
                }, null);

                Assert.Equal(ctx.WorkerThreadId, observed);
            }
        }

        /// <summary>
        /// After observing a stale SyncContext on the UI thread, the
        /// marshaller should adopt the live one so off-thread callers that
        /// arrive later still get their work executed.
        /// </summary>
        [Fact]
        public void RunAsync_AdoptsLiveSyncContext_OnFirstUiThreadCall()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                // Construct the marshaller with a placeholder context that
                // doesn't actually pump anything (mimics our startup-time
                // capture of a SyncContext that VSTO will later swap out).
                var orphan = new SimpleSyncContext();
                var marshaller = new OutlookThreadMarshaller(orphan, ctx.WorkerThreadId);
                Assert.Same(orphan, marshaller.CurrentContext);

                // Now make a call on the target thread. The marshaller should
                // notice SynchronizationContext.Current is different and adopt it.
                ctx.Send(_ =>
                {
                    marshaller.RunAsync(() => { }, CancellationToken.None).GetAwaiter().GetResult();
                }, null);

                Assert.NotSame(orphan, marshaller.CurrentContext);
            }
        }

        /// <summary>Minimal Post-only SyncContext used to simulate an orphan.</summary>
        private sealed class SimpleSyncContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback d, object state) { /* discard */ }
        }

        [Fact]
        public void RunAsyncGeneric_SameContextCaller_InvokesSynchronously_NoDeadlock()
        {
            using (var ctx = new SingleThreadSyncContext())
            {
                var marshaller = new OutlookThreadMarshaller(ctx, ctx.WorkerThreadId);
                int observed = -1;

                ctx.Send(_ =>
                {
                    // Pre-fix: hang. Post-fix: returns synchronously.
                    var v = marshaller.RunAsync(
                        () => Thread.CurrentThread.ManagedThreadId,
                        CancellationToken.None).GetAwaiter().GetResult();
                    observed = v;
                }, null);

                Assert.Equal(ctx.WorkerThreadId, observed);
            }
        }

        /// <summary>
        /// Minimal single-thread SynchronizationContext. The worker thread
        /// drains posted callbacks until the context is disposed. Exposes
        /// its worker's managed-thread-id for assertions.
        /// </summary>
        private sealed class SingleThreadSyncContext : SynchronizationContext, IDisposable
        {
            private readonly BlockingCollection<(SendOrPostCallback cb, object state)> _q
                = new BlockingCollection<(SendOrPostCallback, object)>();
            private readonly Thread _worker;
            private int _workerThreadId;
            private readonly ManualResetEventSlim _ready = new ManualResetEventSlim();

            public int WorkerThreadId
            {
                get { _ready.Wait(); return _workerThreadId; }
            }

            public SingleThreadSyncContext()
            {
                _worker = new Thread(Pump) { IsBackground = true, Name = "SingleThreadSyncContext" };
                _worker.Start();
            }

            private void Pump()
            {
                // Install ourselves as the current thread's SynchronizationContext
                // so reentrancy checks (e.g. SynchronizationContext.Current == _context)
                // work the same way production WinForms code sees them.
                SynchronizationContext.SetSynchronizationContext(this);
                _workerThreadId = Thread.CurrentThread.ManagedThreadId;
                _ready.Set();
                foreach (var (cb, st) in _q.GetConsumingEnumerable())
                {
                    cb(st);
                }
            }

            public override void Post(SendOrPostCallback d, object state) => _q.Add((d, state));

            public override void Send(SendOrPostCallback d, object state)
            {
                // If we're already on the worker, call directly.
                if (Thread.CurrentThread.ManagedThreadId == _workerThreadId)
                {
                    d(state);
                    return;
                }
                // Otherwise enqueue + block until the worker has finished.
                using (var done = new ManualResetEventSlim())
                {
                    Exception captured = null;
                    _q.Add((s =>
                    {
                        try { d(s); }
                        catch (Exception ex) { captured = ex; }
                        finally { done.Set(); }
                    }, state));
                    done.Wait();
                    if (captured != null) throw captured;
                }
            }

            public void Dispose()
            {
                _q.CompleteAdding();
                _worker.Join();
                _q.Dispose();
                _ready.Dispose();
            }
        }
    }
}
