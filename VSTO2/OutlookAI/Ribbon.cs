using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookAI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            // Outlook calls GetCustomUI once per ribbon context. We serve
            // the same customUI XML to both compose (Phase 2 button on
            // TabNewMailMessage) and Explorer (Phase 3a button on TabMail).
            // Tab definitions that don't match the current context are
            // silently ignored by Office's ribbon-XML applier, so a single
            // XML payload safely covers both cases.
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("OutlookAI.Ribbon.xml");
            }
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAIAssistantClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string name in resourceNames)
            {
                if (name.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase))
                {
                    using (StreamReader reader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }
    }
}