using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Per-Explorer task pane that hosts the Inbox Copilot chat surface.
    /// Construction is cheap; <see cref="Bind"/> wires up the per-Explorer
    /// LiveOutlookSurface, OutlookToolHost, ConversationStore, and
    /// InboxCopilotController. The Explorer reference itself flows through
    /// to LiveOutlookSurface so outlook_get_current_selection can read
    /// Explorer.Selection.
    /// </summary>
    public partial class InboxCopilotPane : UserControl
    {
        private Outlook.Explorer _explorer;
        private LiveOutlookSurface _surface;
        private OutlookToolHost _toolHost;
        private ConversationStore _conversationStore;
        private InboxCopilotController _controller;

        public InboxCopilotPane()
        {
            using (TraceLog.Scope("ctor", "InboxCopilotPane"))
            {
                InitializeComponent();
                this.HandleCreated += (s, e) => TraceLog.Write("HandleCreated", "InboxCopilotPane");
                this.VisibleChanged += (s, e) => TraceLog.Write("VisibleChanged Visible=" + this.Visible, "InboxCopilotPane");
            }
        }

        private CodexChatService ChatService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.ChatService : null;

        /// <summary>
        /// Bind this pane to its owning Outlook Explorer. Called by
        /// <see cref="ThisAddIn.ShowExplorerTaskPane"/> immediately after
        /// construction.
        /// </summary>
        public void Bind(Outlook.Explorer explorer)
        {
            using (TraceLog.Scope("Bind", "InboxCopilotPane"))
            {
                _explorer = explorer;
                try
                {
                    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                    var ids = Globals.ThisAddIn?.IdResolver;
                    var app = Globals.ThisAddIn?.Application;
                    TraceLog.Write("Services: marshaller=" + (marshaller != null) +
                        " ids=" + (ids != null) + " app=" + (app != null), "InboxCopilotPane");
                    if (marshaller != null && ids != null && app != null)
                    {
                        _surface = new LiveOutlookSurface(app, marshaller, ids,
                            composeInspector: null, explorer: explorer);
                        _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
                        TraceLog.Write("surface + toolHost constructed", "InboxCopilotPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("surface/toolHost error: " + ex, "InboxCopilotPane");
                }

                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        _conversationStore = new ConversationStore();
                        _controller = new InboxCopilotController(
                            chatHost, ChatService, _toolHost, _surface, _conversationStore, explorer);
                        TraceLog.Write("Controller constructed; firing InitializeAsync", "InboxCopilotPane");
                        var initTask = _controller.InitializeAsync();
                        initTask.ContinueWith(t =>
                        {
                            if (t.IsFaulted)
                                TraceLog.Write("InitializeAsync FAULTED: " + t.Exception, "InboxCopilotPane");
                            else if (t.IsCanceled)
                                TraceLog.Write("InitializeAsync CANCELLED", "InboxCopilotPane");
                            else
                                TraceLog.Write("InitializeAsync completed", "InboxCopilotPane");
                        }, TaskScheduler.Default);
                    }
                    else
                    {
                        TraceLog.Write("Controller NOT created (ChatService/toolHost/surface null)", "InboxCopilotPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("Controller construction error: " + ex, "InboxCopilotPane");
                }
            }
        }

        partial void DisposeCustomResources()
        {
            try { _controller?.Dispose(); } catch { }
        }
    }
}
