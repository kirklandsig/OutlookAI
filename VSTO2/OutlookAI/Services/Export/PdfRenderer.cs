using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using OutlookAI.Diagnostics;
using OutlookAI.TaskPane.Chat;

namespace OutlookAI.Services.Export
{
    public sealed class PdfRenderer : IDisposable
    {
        private const string LogSource = "PdfRenderer";
        private readonly SemaphoreSlim _renderLock = new SemaphoreSlim(1, 1);
        private readonly object _threadLock = new object();
        private readonly object _lifetimeLock = new object();
        private Thread _uiThread;
        private TaskCompletionSource<bool> _threadReady;
        private ApplicationContext _applicationContext;
        private Form _hostForm;
        private WebView2 _webView;
        private CoreWebView2Environment _environment;
        private bool _disposed;

        public async Task<long> RenderAsync(string html, string outputPath, CancellationToken ct)
        {
            lock (_lifetimeLock)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(PdfRenderer));
            }

            if (string.IsNullOrWhiteSpace(html)) throw new ArgumentException("HTML is required.", nameof(html));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path is required.", nameof(outputPath));

            await _renderLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                if (_disposed) throw new ObjectDisposedException(nameof(PdfRenderer));
                TraceLog.Write("RenderAsync start path=" + outputPath, LogSource);
                var length = await RunOnRendererThreadAsync(() => RenderOnRendererThreadAsync(html, outputPath, ct), ct).ConfigureAwait(false);
                TraceLog.Write("RenderAsync complete bytes=" + length, LogSource);
                return length;
            }
            finally
            {
                _renderLock.Release();
            }
        }

        private async Task<long> RenderOnRendererThreadAsync(string html, string outputPath, CancellationToken ct)
        {
            await EnsureInitializedAsync(ct).ConfigureAwait(true);
            ct.ThrowIfCancellationRequested();

            await NavigateAsync(html, ct).ConfigureAwait(true);
            await WaitForRenderReadyAsync(ct).ConfigureAwait(true);
            await PrintAsync(outputPath, ct).ConfigureAwait(true);

            return new FileInfo(outputPath).Length;
        }

        private async Task EnsureInitializedAsync(CancellationToken ct)
        {
            if (_webView.CoreWebView2 != null) return;

            try
            {
                if (!WebView2Bootstrap.IsRuntimeInstalled())
                {
                    throw new ExportException("webview2_missing", "Microsoft Edge WebView2 Runtime is required to export PDF files.");
                }

                TraceLog.Write("Initializing WebView2", LogSource);
                EnsureWebUiResources();
                ct.ThrowIfCancellationRequested();

                _environment = await CoreWebView2Environment.CreateAsync(null, WebView2Bootstrap.WebView2DataFolder, null).ConfigureAwait(true);
                ct.ThrowIfCancellationRequested();
                await _webView.EnsureCoreWebView2Async(_environment).ConfigureAwait(true);
                ct.ThrowIfCancellationRequested();

                _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                _webView.CoreWebView2.Settings.IsZoomControlEnabled = false;
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    WebView2Bootstrap.VirtualHost,
                    WebView2Bootstrap.WebUiFolder,
                    CoreWebView2HostResourceAccessKind.Allow);

                TraceLog.Write("WebView2 initialized", LogSource);
            }
            catch (ExportException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new ExportException("webview2_init_failed", ex.Message, ex);
            }
        }

        private async Task NavigateAsync(string html, CancellationToken ct)
        {
            var completion = new TaskCompletionSource<bool>();
            EventHandler<CoreWebView2NavigationCompletedEventArgs> handler = null;
            handler = (sender, args) =>
            {
                if (args.IsSuccess)
                {
                    completion.TrySetResult(true);
                    return;
                }

                if (ct.IsCancellationRequested)
                {
                    completion.TrySetResult(true);
                    return;
                }

                completion.TrySetException(new ExportException(
                    "pdf_render_failed",
                    "WebView2 navigation failed: " + args.WebErrorStatus));
            };

            _webView.CoreWebView2.NavigationCompleted += handler;
            try
            {
                ct.ThrowIfCancellationRequested();
                using (ct.Register(() => StopNavigation()))
                {
                    _webView.CoreWebView2.NavigateToString(WithVirtualHostBase(html));
                    await completion.Task.ConfigureAwait(true);
                }

                ct.ThrowIfCancellationRequested();
            }
            catch (ExportException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new ExportException("pdf_render_failed", ex.Message, ex);
            }
            finally
            {
                _webView.CoreWebView2.NavigationCompleted -= handler;
            }
        }

        private static string WithVirtualHostBase(string html)
        {
            if (html.IndexOf("<base", StringComparison.OrdinalIgnoreCase) >= 0) return html;

            var headStart = html.IndexOf("<head", StringComparison.OrdinalIgnoreCase);
            if (headStart < 0) return "<base href=\"https://" + WebView2Bootstrap.VirtualHost + "/\">" + html;

            var headEnd = html.IndexOf('>', headStart);
            if (headEnd < 0) return "<base href=\"https://" + WebView2Bootstrap.VirtualHost + "/\">" + html;

            return html.Insert(headEnd + 1, "<base href=\"https://" + WebView2Bootstrap.VirtualHost + "/\">");
        }

        private async Task WaitForRenderReadyAsync(CancellationToken ct)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();
                var json = await _webView.CoreWebView2.ExecuteScriptAsync("document.body ? document.body.getAttribute('data-render-state') : null")
                    .ConfigureAwait(true);
                ct.ThrowIfCancellationRequested();
                var state = JsonConvert.DeserializeObject<string>(json);
                if (string.Equals(state, "ready", StringComparison.OrdinalIgnoreCase)) return;
                if (string.Equals(state, "error", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ExportException("pdf_render_failed", "Print template reported a render error.");
                }

                await Task.Delay(100, ct).ConfigureAwait(true);
            }

            throw new ExportException("pdf_render_timeout", "Print template did not finish rendering within 5 seconds.");
        }

        private async Task PrintAsync(string outputPath, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var settings = _environment.CreatePrintSettings();
                settings.ShouldPrintBackgrounds = true;
                settings.ShouldPrintHeaderAndFooter = false;
                settings.Orientation = CoreWebView2PrintOrientation.Portrait;
                settings.ScaleFactor = 1.0;

                var printed = await _webView.CoreWebView2.PrintToPdfAsync(outputPath, settings).ConfigureAwait(true);
                if (ct.IsCancellationRequested)
                {
                    DeleteOutputFile(outputPath);
                    ct.ThrowIfCancellationRequested();
                }

                if (!printed)
                {
                    throw new ExportException("pdf_print_failed", "WebView2 reported that PDF printing failed.");
                }
            }
            catch (ExportException)
            {
                throw;
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                DeleteOutputFile(outputPath);
                throw;
            }
            catch (Exception ex)
            {
                throw new ExportException("pdf_print_failed", ex.Message, ex);
            }
        }

        private Task<T> RunOnRendererThreadAsync<T>(Func<Task<T>> action, CancellationToken ct)
        {
            return RunOnRendererThreadCoreAsync(action, ct);
        }

        private async Task<T> RunOnRendererThreadCoreAsync<T>(Func<Task<T>> action, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            EnsureRendererThread();
            await _threadReady.Task.ConfigureAwait(false);
            ct.ThrowIfCancellationRequested();

            var completion = new TaskCompletionSource<T>();
            try
            {
                _hostForm.BeginInvoke(new MethodInvoker(async () =>
                {
                    try
                    {
                        var result = await action().ConfigureAwait(true);
                        completion.TrySetResult(result);
                    }
                    catch (OperationCanceledException)
                    {
                        completion.TrySetCanceled();
                    }
                    catch (Exception ex)
                    {
                        completion.TrySetException(ex);
                    }
                }));
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }

            return await completion.Task.ConfigureAwait(false);
        }

        private void EnsureRendererThread()
        {
            lock (_threadLock)
            {
                if (_uiThread != null) return;

                _threadReady = new TaskCompletionSource<bool>();
                _uiThread = new Thread(RendererThreadMain);
                _uiThread.Name = "OutlookAI PDF Renderer";
                _uiThread.IsBackground = true;
                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.Start();
            }
        }

        private void RendererThreadMain()
        {
            try
            {
                WindowsFormsSynchronizationContext.AutoInstall = true;
                _hostForm = CreateHostForm();
                _webView = new WebView2 { Dock = DockStyle.Fill };
                _hostForm.Controls.Add(_webView);
                _applicationContext = new ApplicationContext(_hostForm);
                _hostForm.Show();
                _threadReady.TrySetResult(true);
                Application.Run(_applicationContext);
            }
            catch (Exception ex)
            {
                _threadReady.TrySetException(ex);
            }
        }

        private void StopNavigation()
        {
            try
            {
                var form = _hostForm;
                if (form == null || form.IsDisposed) return;

                var stop = new MethodInvoker(() =>
                {
                    try
                    {
                        if (_webView != null && _webView.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.Stop();
                        }
                    }
                    catch (Exception ex)
                    {
                        TraceLog.Write("StopNavigation invoke failed: " + ex.Message, LogSource);
                    }
                });

                if (form.InvokeRequired)
                {
                    form.BeginInvoke(stop);
                }
                else
                {
                    stop();
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("StopNavigation failed: " + ex.Message, LogSource);
            }
        }

        private static Form CreateHostForm()
        {
            return new HiddenHostForm
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Location = new System.Drawing.Point(-32000, -32000),
                Size = new System.Drawing.Size(1024, 768),
                Opacity = 0,
                Text = "OutlookAI PDF Renderer"
            };
        }

        private static void EnsureWebUiResources()
        {
            Directory.CreateDirectory(WebView2Bootstrap.WebUiFolder);
            Directory.CreateDirectory(WebView2Bootstrap.WebView2DataFolder);

            var assembly = Assembly.GetExecutingAssembly();
            const string prefix = "OutlookAI.WebUI.";
            foreach (var name in assembly.GetManifestResourceNames())
            {
                if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                var relative = WebView2Bootstrap.ResourceNameToRelativePath(name.Substring(prefix.Length));
                var targetPath = Path.Combine(WebView2Bootstrap.WebUiFolder, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(targetPath));

                using (var src = assembly.GetManifestResourceStream(name))
                using (var dst = File.Create(targetPath))
                {
                    src.CopyTo(dst);
                }
            }
        }

        private static void DeleteOutputFile(string outputPath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("Delete canceled PDF failed: " + ex.Message, LogSource);
            }
        }

        public void Dispose()
        {
            lock (_lifetimeLock)
            {
                if (_disposed) return;
                _disposed = true;
            }

            var lockTaken = false;
            try
            {
                lockTaken = _renderLock.Wait(TimeSpan.FromSeconds(10));
                if (!lockTaken)
                {
                    TraceLog.Write("Dispose timed out waiting for active render", LogSource);
                    return;
                }

                DisposeRendererThreadResources();
            }
            catch (Exception ex)
            {
                TraceLog.Write("Dispose failed: " + ex.Message, LogSource);
            }
            finally
            {
                if (lockTaken)
                {
                    _renderLock.Release();
                }
            }
        }

        private void DisposeRendererThreadResources()
        {
            if (_hostForm != null && !_hostForm.IsDisposed)
            {
                var cleanup = new MethodInvoker(() =>
                {
                    try { _webView?.Dispose(); } catch { }
                    try { _hostForm.Dispose(); } catch { }
                    try { _applicationContext?.ExitThread(); } catch { }
                });

                if (_hostForm.InvokeRequired)
                {
                    _hostForm.Invoke(cleanup);
                }
                else
                {
                    cleanup();
                }
            }

            if (_uiThread != null && Thread.CurrentThread.ManagedThreadId != _uiThread.ManagedThreadId)
            {
                _uiThread.Join(TimeSpan.FromSeconds(5));
            }
        }

        private sealed class HiddenHostForm : Form
        {
            protected override bool ShowWithoutActivation => true;

            protected override CreateParams CreateParams
            {
                get
                {
                    const int wsExNoActivate = 0x08000000;
                    var cp = base.CreateParams;
                    cp.ExStyle |= wsExNoActivate;
                    return cp;
                }
            }
        }
    }
}
