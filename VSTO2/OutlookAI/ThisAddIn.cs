using System;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAI.Services;
using OutlookAI.TaskPane;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        public CodexAuthService AuthService { get; private set; }
        public CodexChatService ChatService { get; private set; }
        public RealtimeVoiceService VoiceService { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                AuthService = new CodexAuthService(Config.CodexAuthPath);
                ChatService = new CodexChatService(AuthService);
                VoiceService = new RealtimeVoiceService(AuthService);
            }
            catch (Exception ex)
            {
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
            try
            {
                var inspector = this.Application.ActiveInspector();
                if (inspector == null)
                {
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

                // Create new task pane
                var taskPaneControl = new AITaskPane();
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                customTaskPane.Width = 280;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
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