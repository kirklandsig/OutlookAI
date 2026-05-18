using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// Per-Explorer task pane that hosts the Inbox Reports chat surface.
    /// Parallel to <see cref="InboxCopilot.InboxCopilotPane"/>. Builds a
    /// fresh LiveOutlookSurface + OutlookToolHost + ConversationStore +
    /// InboxReportsController on Bind. Independent state from the Copilot
    /// pane so the two coexist cleanly.
    /// </summary>
    public partial class InboxReportsPane : UserControl
    {
        private Outlook.Explorer _explorer;
        private LiveOutlookSurface _surface;
        private OutlookToolHost _toolHost;
        private ConversationStore _conversationStore;
        private InboxReportsController _controller;

        public InboxReportsPane()
        {
            using (TraceLog.Scope("ctor", "InboxReportsPane"))
            {
                InitializeComponent();
                this.HandleCreated += (s, e) => TraceLog.Write("HandleCreated", "InboxReportsPane");
                this.VisibleChanged += (s, e) => TraceLog.Write("VisibleChanged Visible=" + this.Visible, "InboxReportsPane");
            }
        }

        private CodexChatService ChatService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.ChatService : null;

        public void Bind(Outlook.Explorer explorer)
        {
            using (TraceLog.Scope("Bind", "InboxReportsPane"))
            {
                _explorer = explorer;
                try
                {
                    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                    var ids = Globals.ThisAddIn?.IdResolver;
                    var app = Globals.ThisAddIn?.Application;
                    var runner = Globals.ThisAddIn?.AdvancedSearchRunner;
                    var classifier = Globals.ThisAddIn?.FolderClassifier;
                    TraceLog.Write("Services: marshaller=" + (marshaller != null) +
                        " ids=" + (ids != null) + " app=" + (app != null) +
                        " runner=" + (runner != null), "InboxReportsPane");
                    if (marshaller != null && ids != null && app != null && runner != null)
                    {
                        _surface = new LiveOutlookSurface(app, marshaller, ids,
                            composeInspector: null, explorer: explorer,
                            runner: runner, classifier: classifier);
                        _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
                        TraceLog.Write("surface + toolHost constructed", "InboxReportsPane");
                    }
                    else
                    {
                        TraceLog.Write("surface NOT constructed (missing service); runner=" + (runner != null),
                            "InboxReportsPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("surface/toolHost error: " + ex, "InboxReportsPane");
                }

                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        _conversationStore = new ConversationStore();
                        _controller = new InboxReportsController(
                            chatHost, ChatService, _toolHost, _conversationStore);
                        TraceLog.Write("Controller constructed; firing InitializeAsync", "InboxReportsPane");
                        var initTask = _controller.InitializeAsync();
                        initTask.ContinueWith(t =>
                        {
                            if (t.IsFaulted)
                                TraceLog.Write("InitializeAsync FAULTED: " + t.Exception, "InboxReportsPane");
                            else if (t.IsCanceled)
                                TraceLog.Write("InitializeAsync CANCELLED", "InboxReportsPane");
                            else
                                TraceLog.Write("InitializeAsync completed", "InboxReportsPane");
                        }, TaskScheduler.Default);
                    }
                    else
                    {
                        TraceLog.Write("Controller NOT created (ChatService/toolHost/surface null)", "InboxReportsPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("Controller construction error: " + ex, "InboxReportsPane");
                }
            }
        }

        partial void DisposeCustomResources()
        {
            try { _controller?.Dispose(); } catch { }
        }
    }
}
