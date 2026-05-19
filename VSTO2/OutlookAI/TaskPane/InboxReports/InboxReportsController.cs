using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Export;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// Drives the per-Explorer Inbox Reports chat surface. Mirrors
    /// <see cref="InboxCopilot.InboxCopilotController"/> with two
    /// differences: (1) builds the reports-focused system prompt via
    /// <see cref="InboxReportsPromptBuilder"/>, (2) pushes the six fixed
    /// <see cref="ReportQuickActionChip"/> templates once on ready
    /// rather than recomputing chips on selection / folder change.
    /// </summary>
    public sealed class InboxReportsController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CodexChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly InboxReportsPromptBuilder _promptBuilder = new InboxReportsPromptBuilder();
        private readonly ExportBridge _exportBridge;

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public InboxReportsController(
            Control hostContainer,
            CodexChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
        }

        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "InboxReports");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                ShowFallback("WebView2 runtime not installed.\r\nRun the installer or download:\r\n" +
                             "https://developer.microsoft.com/microsoft-edge/webview2/");
                return;
            }
            _webView = new WebView2 { Dock = DockStyle.Fill };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "InboxReports");
                ShowFallback("WebView2 failed to initialize: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null) { _fallbackLabel.Text = message; return; }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message,
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "InboxReports");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "InboxReports");
            }
        }

        private async Task HandleHostMessageAsync(string type, JObject payload)
        {
            try
            {
                if (_exportBridge != null && await _exportBridge.HandleAsync(
                    type, payload, _activeCts?.Token ?? CancellationToken.None).ConfigureAwait(false))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("ExportBridge EXCEPTION: " + ex, "InboxReports");
                return;
            }

            switch (type)
            {
                case "ready":
                    OnWebViewReady();
                    break;
                case "send":
                    _ = StartTurnAsync(
                        (string)payload?["text"] ?? "",
                        (string)payload?["reasoning"]);
                    break;
                case "stop":
                    try { _activeCts?.Cancel(); } catch { }
                    break;
                case "clear":
                    _store.Clear();
                    _ = RunScript("outlookai.clear();");
                    break;
                case "copy":
                    var clip = _store.ExportForClipboard();
                    try { Clipboard.SetText(clip ?? ""); } catch { }
                    break;
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "InboxReports");
            _isReady = true;
            _ = RunScript("outlookai.applyTheme('light');");
            PushReasoningOptions();
            PushReportChips();
            TraceLog.Write("OnWebViewReady completed", "InboxReports");
        }

        private void PushReasoningOptions()
        {
            try
            {
                var efforts = Config.ReasoningEffortsForModel(Config.Model);
                var arr = new JArray();
                foreach (var e in efforts) arr.Add(e);
                _ = RunScript("outlookai.setReasoningOptions(" +
                    arr.ToString(Newtonsoft.Json.Formatting.None) + ", '');");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushReasoningOptions error: " + ex.Message, "InboxReports");
            }
        }

        private void PushReportChips()
        {
            try
            {
                var chips = ReportQuickActionChip.Defaults();
                var chipsArr = new JArray();
                foreach (var c in chips)
                {
                    chipsArr.Add(new JObject(
                        new JProperty("label", c.Label),
                        new JProperty("prompt", c.TemplateText)));
                }
                // Reports chips are prefill-only (not auto-submit) so the
                // user can edit [placeholders] in the template before
                // sending. Mirrors the spec.
                _ = RunScript("outlookai.setQuickActions(" +
                    chipsArr.ToString(Newtonsoft.Json.Formatting.None) +
                    ", {autoSubmit: false});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushReportChips error: " + ex.Message, "InboxReports");
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxReports");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxReports");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("outlookai.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "rep_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = _promptBuilder.Build(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    ReasoningEffortOverride = string.IsNullOrEmpty(reasoningOverride) ? null : reasoningOverride,
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                {
                    _store.Append(ctx.History[i]);
                }
                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                await RunScript("outlookai.showError(" + JsString(ex.Message ?? "") + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "InboxReports");
            }
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
                {
                    await marshaller.RunAsync(() =>
                    {
                        if (_isDisposed) return;
                        var core = _webView?.CoreWebView2;
                        if (core == null) return;
                        _ = core.ExecuteScriptAsync(script);
                    }, CancellationToken.None).ConfigureAwait(false);
                    return;
                }
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "InboxReports");
            }
        }

        private static string JsString(string s)
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try { _webView?.Dispose(); } catch { }
        }

        private sealed class WebViewSink : ChatEventSink
        {
            private readonly InboxReportsController _owner;
            private readonly string _assistantId;
            public WebViewSink(InboxReportsController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                _ = _owner.RunScript("outlookai.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                TraceLog.Write("Sink.OnToolCallStart " + name + " args=" + (argsJson?.Length > 200 ? argsJson.Substring(0, 200) + "..." : argsJson), "WebViewSink");
                _ = _owner.RunScript("outlookai.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " summary=" + summary + " resultLen=" + (resultJson?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript("outlookai.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
            }
            public override void OnError(string message)
            {
                TraceLog.Write("Sink.OnError: " + message, "WebViewSink");
                _ = _owner.RunScript("outlookai.showError(" + JsString(message ?? "") + ");");
            }
            public override void OnAssistantMessageComplete(string text)
            {
                TraceLog.Write("Sink.OnAssistantMessageComplete len=" + (text?.Length ?? 0), "WebViewSink");
            }
        }
    }
}
