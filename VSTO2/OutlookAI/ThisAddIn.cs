using System;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.TaskPane;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        public CodexAuthService AuthService { get; private set; }
        public CodexChatService ChatService { get; private set; }
        public RealtimeVoiceService VoiceService { get; private set; }
        public OutlookThreadMarshaller OutlookMarshaller { get; private set; }
        public IdResolver IdResolver { get; private set; }

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
                VoiceService?.Dispose();
                ChatService?.Dispose();
                AuthService?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ThisAddIn_Shutdown error: " + ex);
            }
        }

        public void ShowTaskPane()
        {
            using (TraceLog.Scope("ShowTaskPane", "ThisAddIn"))
            try
            {
                var inspector = this.Application.ActiveInspector();
                if (inspector == null)
                {
                    TraceLog.Write("No active inspector", "ThisAddIn");
                    System.Windows.Forms.MessageBox.Show(
                        "Please open an email first.",
                        "AI Assistant",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Check if task pane already exists for this inspector
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Window == inspector)
                    {
                        TraceLog.Write("Reusing existing CustomTaskPane (toggling visibility)", "ThisAddIn");
                        // Toggle visibility, reset if showing
                        if (!pane.Visible)
                        {
                            var existingControl = pane.Control as AITaskPane;
                            existingControl?.ResetForNewEmail();
                        }
                        pane.Visible = !pane.Visible;
                        return;
                    }
                }

                TraceLog.Write("Creating new AITaskPane", "ThisAddIn");
                var taskPaneControl = new AITaskPane();
                TraceLog.Write("AITaskPane constructed; calling Bind", "ThisAddIn");
                taskPaneControl.Bind(inspector);
                TraceLog.Write("Bind returned; calling CustomTaskPanes.Add", "ThisAddIn");
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                TraceLog.Write("CustomTaskPane.Add returned", "ThisAddIn");
                customTaskPane.Width = 340;
                customTaskPane.Visible = true;
                TraceLog.Write("CustomTaskPane.Visible = true", "ThisAddIn");
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