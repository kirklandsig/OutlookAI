using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace OutlookAI.Tests
{
    public class RibbonTests
    {
        private static readonly XNamespace Ui = "http://schemas.microsoft.com/office/2009/07/customui";

        // The XML Outlook is served, for the Inbox and for compose windows.
        private static XDocument Served(string ribbonId) => XDocument.Parse(new Ribbon().GetCustomUI(ribbonId));

        private static XElement AiAssistantGroup(string tab) =>
            Served(tab == "TabMail" ? "Microsoft.Outlook.Explorer" : "Microsoft.Outlook.Mail.Compose")
                .Descendants(Ui + "tab").Single(t => (string)t.Attribute("idMso") == tab)
                .Elements(Ui + "group").Single(g => (string)g.Attribute("label") == "AI Assistant");

        // Settings (sign-in, model, updates) is on the ribbon wherever AI
        // Assistant is, not only behind the gear in the compose pane.
        [Theory]
        [InlineData("TabMail")]
        [InlineData("TabNewMailMessage")]
        public void SettingsButton_IsInTheAiAssistantGroup(string tab)
        {
            var settings = AiAssistantGroup(tab).Elements(Ui + "button")
                .SingleOrDefault(b => (string)b.Attribute("onAction") == "OnSettingsClick");

            Assert.NotNull(settings);
            Assert.Equal("Settings", (string)settings.Attribute("label"));
            Assert.False(string.IsNullOrEmpty((string)settings.Attribute("imageMso")));
        }

        // Office drops the add-in's whole ribbon when a callback is missing or
        // two controls share an id.
        [Fact]
        public void EveryCallback_HasAPublicHandler()
        {
            var callbacks = Served("Microsoft.Outlook.Explorer").Descendants()
                .SelectMany(e => e.Attributes())
                .Where(a => a.Name.LocalName.StartsWith("on") || a.Name.LocalName.StartsWith("get"))
                .Select(a => a.Value)
                .Distinct()
                .ToList();

            Assert.Contains("OnSettingsClick", callbacks);
            foreach (var name in callbacks)
            {
                var handler = typeof(Ribbon).GetMethod(name, BindingFlags.Public | BindingFlags.Instance);
                Assert.True(handler != null && handler.GetParameters().Length == 1, "No handler for " + name);
            }
        }

        [Fact]
        public void ControlIds_AreUnique()
        {
            var ids = Served("Microsoft.Outlook.Explorer").Descendants()
                .Select(e => (string)e.Attribute("id"))
                .Where(id => id != null)
                .ToList();

            Assert.Equal(ids.Count, ids.Distinct().Count());
        }
    }
}
