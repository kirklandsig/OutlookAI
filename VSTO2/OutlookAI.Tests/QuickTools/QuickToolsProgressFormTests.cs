using System.Threading;
using System.Windows.Forms;
using OutlookAI.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.QuickTools
{
    public class QuickToolsProgressFormTests
    {
        // Closing the opened email destroys the window that owns the progress window,
        // and the progress window with it, without a FormClosing.
        [Fact]
        public void OwnerDestroyed_CancelsTheRun()
        {
            Sta.Run(() =>
            {
                var owner = new NativeWindow();
                owner.CreateHandle(new CreateParams());
                using (var cts = new CancellationTokenSource())
                using (var form = new QuickToolsProgressForm(cts))
                {
                    form.Show(owner);
                    owner.DestroyHandle();
                    Application.DoEvents();

                    Assert.True(cts.IsCancellationRequested);
                }
            });
        }

        [Fact]
        public void Finish_ClosesWithoutCancelling()
        {
            Sta.Run(() =>
            {
                var owner = new NativeWindow();
                owner.CreateHandle(new CreateParams());
                try
                {
                    using (var cts = new CancellationTokenSource())
                    using (var form = new QuickToolsProgressForm(cts))
                    {
                        form.Show(owner);
                        form.Finish();
                        Application.DoEvents();

                        Assert.False(cts.IsCancellationRequested);
                    }
                }
                finally
                {
                    owner.DestroyHandle();
                }
            });
        }

        private sealed class RawHandle : IWin32Window
        {
            public RawHandle(System.IntPtr handle)
            {
                Handle = handle;
            }

            public System.IntPtr Handle { get; }
        }

        // OwnerWindow keeps the raw handle, which outlives the window it named.
        [Fact]
        public void OwnerWindow_Alive_IsNullOnceTheWindowIsGone()
        {
            Sta.Run(() =>
            {
                var window = new NativeWindow();
                window.CreateHandle(new CreateParams());
                var owner = new RawHandle(window.Handle);

                Assert.Same(owner, OwnerWindow.Alive(owner));
                window.DestroyHandle();
                Assert.Null(OwnerWindow.Alive(owner));
                Assert.Null(OwnerWindow.Alive(null));
            });
        }
    }
}
