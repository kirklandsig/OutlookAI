using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookAdvancedSearchRunnerTests
    {
        private static MessageProjectionInput Item(string id) => new MessageProjectionInput
        {
            Id = id,
            Subject = id,
            From = "f",
            To = new string[0],
            ReceivedAt = DateTimeOffset.UtcNow,
            FolderName = "Inbox",
            FolderDefaultItemTypeIsMail = true,
            SnippetFactory = () => id,
        };

        [Fact]
        public async Task RunAsync_HappyPath_ReturnsCompletedWithItems()
        {
            var host = new FakeAdvancedSearchHost();
            using (var runner = new OutlookAdvancedSearchRunner(host))
            {
                var task = runner.RunAsync(
                    "'\\\\store\\Inbox'", "filter", true,
                    TimeSpan.FromSeconds(5), CancellationToken.None);

                var tag = await WaitForTag(host);
                host.RaiseCompleted(tag, new[] { Item("a"), Item("b") });

                var result = await task;
                Assert.True(result.Completed);
                Assert.False(result.TimedOut);
                Assert.False(result.Cancelled);
                Assert.Null(result.Error);
                Assert.Equal(2, result.Items.Count);
                Assert.Equal(new[] { "a", "b" }, result.Items.Select(i => i.Id).ToArray());
            }
        }

        [Fact]
        public async Task RunAsync_NoCompletion_TimesOutAndStops()
        {
            var host = new FakeAdvancedSearchHost();
            using (var runner = new OutlookAdvancedSearchRunner(host))
            {
                var task = runner.RunAsync(
                    "scope", "filter", true,
                    TimeSpan.FromMilliseconds(50), CancellationToken.None);
                var result = await task;

                Assert.False(result.Completed);
                Assert.True(result.TimedOut);
                Assert.Single(host.StartCalls);
                Assert.Single(host.StopCalls);
                Assert.Equal(host.StartCalls[0].Tag, host.StopCalls[0]);
            }
        }

        [Fact]
        public async Task RunAsync_CancellationToken_Cancelled_StopsAndReturnsCancelled()
        {
            var host = new FakeAdvancedSearchHost();
            using (var runner = new OutlookAdvancedSearchRunner(host))
            using (var cts = new CancellationTokenSource())
            {
                var task = runner.RunAsync(
                    "scope", "filter", true,
                    TimeSpan.FromSeconds(30), cts.Token);
                await WaitForTag(host);
                cts.Cancel();

                var result = await task;
                Assert.False(result.Completed);
                Assert.False(result.TimedOut);
                Assert.True(result.Cancelled);
                Assert.Single(host.StopCalls);
            }
        }

        [Fact]
        public async Task RunAsync_StartThrows_ReturnsError_NoSubscriptionLeak()
        {
            var host = new FakeAdvancedSearchHost
            {
                ThrowOnStart = _ => new COMException("simulated")
            };
            using (var runner = new OutlookAdvancedSearchRunner(host))
            {
                var result = await runner.RunAsync(
                    "scope", "filter", true,
                    TimeSpan.FromSeconds(5), CancellationToken.None);

                Assert.False(result.Completed);
                Assert.False(result.TimedOut);
                Assert.False(result.Cancelled);
                Assert.NotNull(result.Error);
                Assert.IsType<COMException>(result.Error);
            }
        }

        [Fact]
        public async Task RunAsync_UnknownTagCompletedEvent_IsIgnored()
        {
            var host = new FakeAdvancedSearchHost();
            using (var runner = new OutlookAdvancedSearchRunner(host))
            {
                var task = runner.RunAsync(
                    "scope", "filter", true,
                    TimeSpan.FromSeconds(5), CancellationToken.None);
                await WaitForTag(host);

                // Fire an event with a tag the runner never issued.
                host.RaiseCompleted("totally-unrelated-tag", new[] { Item("x") });

                // Pending task must still be pending: complete it with the right tag now.
                host.RaiseCompleted(host.StartCalls[0].Tag, new[] { Item("real") });
                var result = await task;

                Assert.True(result.Completed);
                Assert.Equal("real", result.Items[0].Id);
            }
        }

        [Fact]
        public async Task RunAsync_TwoCalls_AreSerialised()
        {
            var host = new FakeAdvancedSearchHost();
            using (var runner = new OutlookAdvancedSearchRunner(host))
            {
                var first = runner.RunAsync(
                    "a", null, true, TimeSpan.FromSeconds(5), CancellationToken.None);
                var second = runner.RunAsync(
                    "b", null, true, TimeSpan.FromSeconds(5), CancellationToken.None);

                // Allow scheduler to run.
                await Task.Delay(50);

                // Only the first Start should have been observed; second waits on the semaphore.
                Assert.Single(host.StartCalls);

                host.RaiseCompleted(host.StartCalls[0].Tag, new MessageProjectionInput[0]);
                await first;

                await WaitForCondition(() => host.StartCalls.Count == 2, TimeSpan.FromSeconds(2));
                host.RaiseCompleted(host.StartCalls[1].Tag, new MessageProjectionInput[0]);
                var secondResult = await second;
                Assert.True(secondResult.Completed);
            }
        }

        private static async Task<string> WaitForTag(FakeAdvancedSearchHost host)
        {
            await WaitForCondition(() => host.StartCalls.Count > 0, TimeSpan.FromSeconds(2));
            return host.StartCalls[host.StartCalls.Count - 1].Tag;
        }

        private static async Task WaitForCondition(Func<bool> cond, TimeSpan timeout)
        {
            var start = DateTime.UtcNow;
            while (!cond() && DateTime.UtcNow - start < timeout)
            {
                await Task.Delay(10);
            }
            if (!cond()) throw new TimeoutException("condition not met");
        }
    }
}
