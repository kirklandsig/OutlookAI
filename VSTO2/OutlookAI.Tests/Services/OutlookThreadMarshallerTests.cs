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
                var marshaller = new OutlookThreadMarshaller(ctx);
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
                var marshaller = new OutlookThreadMarshaller(ctx);
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
                var marshaller = new OutlookThreadMarshaller(ctx);
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
                var marshaller = new OutlookThreadMarshaller(ctx);
                var cts = new CancellationTokenSource();
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    marshaller.RunAsync(() => { }, cts.Token));
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
                _workerThreadId = Thread.CurrentThread.ManagedThreadId;
                _ready.Set();
                foreach (var (cb, st) in _q.GetConsumingEnumerable())
                {
                    cb(st);
                }
            }

            public override void Post(SendOrPostCallback d, object state) => _q.Add((d, state));
            public override void Send(SendOrPostCallback d, object state) => _q.Add((d, state));

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
