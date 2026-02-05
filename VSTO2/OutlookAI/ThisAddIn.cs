using System;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAI.TaskPane;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Add-in started
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup
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