using System;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Export;
using OutlookAI.TaskPane;
using OutlookAI.TaskPane.InboxCopilot;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        public CodexAuthService AuthService { get; private set; }
        public CodexChatService ChatService { get; private set; }
        public RealtimeVoiceService VoiceService { get; private set; }
        public OutlookThreadMarshaller OutlookMarshaller { get; private set; }
        public IdResolver IdResolver { get; private set; }
        public OutlookAI.Services.Tools.IOutlookAdvancedSearchRunner AdvancedSearchRunner { get; private set; }
        public OutlookAI.Services.Tools.IFolderClassifier FolderClassifier { get; private set; }
        private PdfRenderer _pdfRenderer;
        public PdfRenderer PdfRenderer => _pdfRenderer ?? (_pdfRenderer = new PdfRenderer());
        private ExportPathResolver _exportPathResolver;
        public ExportPathResolver ExportPathResolver => _exportPathResolver ?? (_exportPathResolver = new ExportPathResolver());
        private IExportPathPolicy _exportPathPolicy;
        public IExportPathPolicy ExportPathPolicy => _exportPathPolicy ?? (_exportPathPolicy = new ExportPathPolicy(ExportPathResolver));

        private OutlookAI.Services.Tools.LiveAdvancedSearchHost _advancedSearchHost;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            TraceLog.MarkUiThread();
            TraceLog.Write("ThisAddIn_Startup entered", "ThisAddIn");

            // Catch every unhandled exception that escapes a WinForms event
            // handler or background task. Helps us see exceptions that would
            // otherwise be swallowed during this freeze investigation.
            System.Windows.Forms.Application.ThreadException += (s, ev) =>
                TraceLog.Write("Application.ThreadException: " + ev.Exception, "GlobalEx");
            System.AppDomain.CurrentDomain.UnhandledException += (s, ev) =>
                TraceLog.Write("AppDomain.UnhandledException: " + ev.ExceptionObject, "GlobalEx");
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, ev) =>
            {
                TraceLog.Write("UnobservedTaskException: " + ev.Exception, "GlobalEx");
                ev.SetObserved();
            };

            try
            {
                AuthService = new CodexAuthService(Config.CodexAuthPath);
                ChatService = new CodexChatService(AuthService);
                VoiceService = new RealtimeVoiceService(AuthService);

                // SynchronizationContext captured on the Outlook UI thread.
                // Forms apps install a WindowsFormsSynchronizationContext per
                // UI thread; for a VSTO add-in, that's done by the runtime
                // before ThisAddIn_Startup. If we somehow don't have one
                // (e.g. running under a non-Forms host in tests), we install
                // one explicitly so OutlookThreadMarshaller has somewhere to
                // post.
                var syncCtx = System.Threading.SynchronizationContext.Current;
                TraceLog.Write("SynchronizationContext.Current is "
                    + (syncCtx == null ? "NULL" : syncCtx.GetType().FullName + " hash=" + syncCtx.GetHashCode()),
                    "ThisAddIn");
                if (syncCtx == null)
                {
                    syncCtx = new System.Windows.Forms.WindowsFormsSynchronizationContext();
                    System.Threading.SynchronizationContext.SetSynchronizationContext(syncCtx);
                    TraceLog.Write("Installed new WindowsFormsSynchronizationContext hash=" + syncCtx.GetHashCode(), "ThisAddIn");
                    // CRITICAL: force the marshaling control's HWND to be
                    // created on THIS (UI) thread. WindowsFormsSynchronizationContext
                    // lazily creates its marshaling control on the thread that
                    // first calls Post against it. If we let a threadpool
                    // thread Post first, the HWND ends up on the threadpool
                    // thread, which has no message pump - every subsequent
                    // post is silently lost. This warm-up Post from the UI
                    // thread guarantees the HWND lives on the UI thread.
                    syncCtx.Post(_ => TraceLog.Write("Warm-up Post fired", "ThisAddIn"), null);
                    TraceLog.Write("Posted warm-up to force marshaling-control HWND on UI thread", "ThisAddIn");
                }
                OutlookMarshaller = new OutlookThreadMarshaller(syncCtx);
                IdResolver = new IdResolver();
                TraceLog.Write("Services initialized OK; marshaller _uiThreadId=" + OutlookMarshaller.UiThreadId, "ThisAddIn");

                // Phase 3b: shared AdvancedSearch host + runner. One host
                // owns the AdvancedSearchComplete subscription for the whole
                // process; one runner serialises in-flight searches.
                FolderClassifier = new OutlookAI.Services.Tools.FolderClassifier();
                _advancedSearchHost = new OutlookAI.Services.Tools.LiveAdvancedSearchHost(
                    this.Application, OutlookMarshaller, IdResolver);
                AdvancedSearchRunner = new OutlookAI.Services.Tools.OutlookAdvancedSearchRunner(_advancedSearchHost);
                TraceLog.Write("AdvancedSearch services constructed", "ThisAddIn");
            }
            catch (Exception ex)
            {
                TraceLog.Write("Startup error: " + ex, "ThisAddIn");
                System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup error: " + ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Runner before host: runner unsubscribes its handler from
                // the host on Dispose; the host then drops its own COM
                // subscription. Reverse order would leave a dangling handler.
                try { AdvancedSearchRunner?.Dispose(); }
                catch (Exception ex) { TraceLog.Write("Runner dispose: " + ex.Message, "ThisAddIn"); }
                try { _advancedSearchHost?.Dispose(); }
                catch (Exception ex) { TraceLog.Write("Host dispose: " + ex.Message, "ThisAddIn"); }
                try { _pdfRenderer?.Dispose(); }
                catch (Exception ex) { TraceLog.Write("PDF renderer dispose: " + ex.Message, "ThisAddIn"); }

                VoiceService?.Dispose();
                ChatService?.Dispose();
                AuthService?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ThisAddIn_Shutdown error: " + ex);
            }
        }

        /// <summary>
        /// Entry point from the ribbon button (registered on both
        /// TabNewMailMessage and TabMail in Ribbon.xml). Phase 3a routes
        /// based on the currently-active Outlook window: Inspector (compose)
        /// -> existing per-Inspector AITaskPane; Explorer (mailbox view)
        /// -> new per-Explorer InboxCopilotPane.
        /// </summary>
        public void ShowTaskPane()
        {
            using (TraceLog.Scope("ShowTaskPane", "ThisAddIn"))
            try
            {
                object activeWindow = null;
                try { activeWindow = this.Application.ActiveWindow(); } catch { }
                TraceLog.Write("ActiveWindow=" + (activeWindow?.GetType().FullName ?? "<null>"), "ThisAddIn");

                if (activeWindow is Outlook.Inspector insp)
                {
                    ShowInspectorTaskPane(insp);
                    return;
                }
                if (activeWindow is Outlook.Explorer expl)
                {
                    ShowExplorerTaskPane(expl);
                    return;
                }

                // Fallback: no compose window or inbox view in focus.
                TraceLog.Write("No Inspector or Explorer active; showing info dialog", "ThisAddIn");
                System.Windows.Forms.MessageBox.Show(
                    "Open Outlook to your Inbox or compose an email, then click AI Assistant.",
                    "AI Assistant",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                TraceLog.Write("ShowTaskPane error: " + ex, "ThisAddIn");
                System.Windows.Forms.MessageBox.Show(
                    $"Error: {ex.Message}",
                    "AI Assistant",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ShowInspectorTaskPane(Outlook.Inspector inspector)
        {
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Window == inspector)
                {
                    TraceLog.Write("Reusing existing Inspector CustomTaskPane", "ThisAddIn");
                    if (!pane.Visible)
                    {
                        var existingControl = pane.Control as AITaskPane;
                        existingControl?.ResetForNewEmail();
                    }
                    pane.Visible = !pane.Visible;
                    return;
                }
            }
            TraceLog.Write("Creating new AITaskPane for Inspector", "ThisAddIn");
            var taskPaneControl = new AITaskPane();
            taskPaneControl.Bind(inspector);
            var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
            customTaskPane.Width = 340;
            customTaskPane.Visible = true;
            TraceLog.Write("Inspector CustomTaskPane.Visible = true", "ThisAddIn");
        }

        private void ShowExplorerTaskPane(Outlook.Explorer explorer)
        {
            // Match only the InboxCopilotPane on this Explorer so the two
            // ribbon buttons each toggle their own pane. The Reports pane
            // also lives on this Explorer; we don't want the Copilot
            // button to toggle the Reports pane and vice versa.
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Window == explorer && pane.Control is InboxCopilotPane)
                {
                    TraceLog.Write("Reusing existing Explorer CustomTaskPane (toggle visibility)", "ThisAddIn");
                    pane.Visible = !pane.Visible;
                    return;
                }
            }
            TraceLog.Write("Creating new InboxCopilotPane for Explorer", "ThisAddIn");
            var paneControl = new InboxCopilotPane();
            paneControl.Bind(explorer);
            var ctp = this.CustomTaskPanes.Add(paneControl, "AI Assistant", explorer);
            ctp.Width = 340;
            ctp.Visible = true;
            TraceLog.Write("Explorer CustomTaskPane.Visible = true", "ThisAddIn");
        }

        public void ShowReportsTaskPane()
        {
            using (TraceLog.Scope("ShowReportsTaskPane", "ThisAddIn"))
                try
                {
                    object activeWindow = null;
                    try { activeWindow = this.Application.ActiveWindow(); } catch { }
                    TraceLog.Write("ActiveWindow=" + (activeWindow?.GetType().FullName ?? "<null>"), "ThisAddIn");

                    if (activeWindow is Outlook.Explorer expl)
                    {
                        ShowReportsExplorerTaskPane(expl);
                        return;
                    }
                    // Reports only makes sense on an Explorer (Inbox view),
                    // not on a compose window.
                    System.Windows.Forms.MessageBox.Show(
                        "Open Outlook to your Inbox, then click Reports.",
                        "Inbox Reports",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    TraceLog.Write("ShowReportsTaskPane error: " + ex, "ThisAddIn");
                    System.Windows.Forms.MessageBox.Show(
                        $"Error: {ex.Message}",
                        "Inbox Reports",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
        }

        private void ShowReportsExplorerTaskPane(Outlook.Explorer explorer)
        {
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Window == explorer && pane.Control is OutlookAI.TaskPane.InboxReports.InboxReportsPane)
                {
                    TraceLog.Write("Reusing existing Reports CustomTaskPane (toggle visibility)", "ThisAddIn");
                    pane.Visible = !pane.Visible;
                    return;
                }
            }
            TraceLog.Write("Creating new InboxReportsPane for Explorer", "ThisAddIn");
            var paneControl = new OutlookAI.TaskPane.InboxReports.InboxReportsPane();
            paneControl.Bind(explorer);
            var ctp = this.CustomTaskPanes.Add(paneControl, "Inbox Reports", explorer);
            ctp.Width = 340;
            ctp.Visible = true;
            TraceLog.Write("Reports CustomTaskPane.Visible = true", "ThisAddIn");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
