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
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Drives the per-Explorer Inbox Copilot chat surface. Mirrors
    /// Phase 2's ChatController but is anchored to an Explorer instead
    /// of an Inspector. Builds a fresh system prompt + quick-action
    /// chip set on every selection change and on every turn.
    /// </summary>
    public sealed class InboxCopilotController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CodexChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Outlook.Explorer _explorer;
        private readonly ExportBridge _exportBridge;

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public InboxCopilotController(
            Control hostContainer,
            CodexChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store,
            Outlook.Explorer explorer)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            _explorer = explorer;
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
        }

        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "InboxCopilot");
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
                if (_explorer != null)
                {
                    _explorer.SelectionChange += OnExplorerSelectionChange;
                    _explorer.FolderSwitch += OnExplorerFolderSwitch;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "InboxCopilot");
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

        private void OnExplorerSelectionChange()
        {
            if (_isDisposed || !_isReady) return;
            PushContextStripAndChips();
        }

        private void OnExplorerFolderSwitch()
        {
            if (_isDisposed || !_isReady) return;
            PushContextStripAndChips();
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "InboxCopilot");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "InboxCopilot");
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
            catch (Exception ex)
            {
                TraceLog.Write("HandleHostMessageAsync EXCEPTION: " + ex, "InboxCopilot");
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "InboxCopilot");
            _isReady = true;
            _ = RunScript("outlookai.applyTheme('light');");
            PushReasoningOptions();
            PushContextStripAndChips();
            TraceLog.Write("OnWebViewReady completed", "InboxCopilot");
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
                TraceLog.Write("PushReasoningOptions error: " + ex.Message, "InboxCopilot");
            }
        }

        private void PushContextStripAndChips()
        {
            try
            {
                CurrentSelectionResult sel = null;
                string folderName = "";
                int unreadCount = 0, totalCount = 0;
                try
                {
                    sel = _surface.GetCurrentSelection(includeFullBodies: false, maxItems: 5);
                    folderName = sel?.Folder ?? "";
                    var folder = _explorer?.CurrentFolder;
                    try
                    {
                        if (folder != null)
                        {
                            unreadCount = folder.UnReadItemCount;
                            totalCount = folder.Items.Count;
                        }
                    }
                    catch { }
                }
                catch (Exception ex) { TraceLog.Write("PushContextStrip surface error: " + ex.Message, "InboxCopilot"); }

                var ctx = new JObject(
                    new JProperty("folder", folderName),
                    new JProperty("unread_count", unreadCount),
                    new JProperty("total_count", totalCount));
                if (sel != null && sel.Count > 0 && sel.Messages != null && sel.Messages.Count > 0)
                {
                    var first = sel.Messages[0];
                    ctx.Add("selection", new JObject(
                        new JProperty("count", sel.Count),
                        new JProperty("subject", first.Subject ?? ""),
                        new JProperty("from", first.From ?? "")));
                }
                _ = RunScript("outlookai.setContextStrip(" +
                    ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");

                var selectionCount = sel?.Count ?? 0;
                var chips = QuickActionChip.ComputeChipsForSelectionCount(selectionCount);
                var chipsArr = new JArray();
                foreach (var c in chips)
                {
                    chipsArr.Add(new JObject(
                        new JProperty("label", c.Label),
                        new JProperty("prompt", c.Prompt)));
                }
                _ = RunScript("outlookai.setQuickActions(" +
                    chipsArr.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStripAndChips error: " + ex.Message, "InboxCopilot");
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxCopilot");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxCopilot");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("outlookai.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildSystemInstructionsForCurrentState(),
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
                TraceLog.Write("<< StartTurnAsync", "InboxCopilot");
            }
        }

        private string BuildSystemInstructionsForCurrentState()
        {
            try
            {
                var sel = _surface?.GetCurrentSelection(includeFullBodies: false, maxItems: 1);
                int unreadCount = 0, totalCount = 0;
                string folderName = sel?.Folder ?? "Inbox";
                try
                {
                    var folder = _explorer?.CurrentFolder;
                    if (folder != null)
                    {
                        unreadCount = folder.UnReadItemCount;
                        totalCount = folder.Items.Count;
                    }
                }
                catch { }
                return InboxCopilotPromptBuilder.Build(folderName, unreadCount, totalCount, sel);
            }
            catch (Exception ex)
            {
                TraceLog.Write("BuildSystemInstructions error: " + ex, "InboxCopilot");
                return "You are the Outlook Inbox Copilot. Help the user with their mailbox.";
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
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "InboxCopilot");
            }
        }

        private static string JsString(string s)
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try
            {
                if (_explorer != null)
                {
                    _explorer.SelectionChange -= OnExplorerSelectionChange;
                    _explorer.FolderSwitch -= OnExplorerFolderSwitch;
                }
            }
            catch { }
            try { _webView?.Dispose(); } catch { }
        }

        private sealed class WebViewSink : ChatEventSink
        {
            private readonly InboxCopilotController _owner;
            private readonly string _assistantId;
            public WebViewSink(InboxCopilotController owner, string assistantId)
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
