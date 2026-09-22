using System;
using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Models
{
    public class SettingsForm_ModelCatalog_SourceTests
    {
        private static string FindSourceFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate)) return candidate;
                current = current.Parent;
            }
            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }

        // Normalized so slicing doesn't depend on how the file was checked out.
        private static string SettingsFormSource =>
            File.ReadAllText(FindSourceFile("OutlookAI", "SettingsForm.cs")).Replace("\r\n", "\n");

        // From `start` up to the end of the statement/member that follows it.
        private static string Slice(string src, string start, string end)
        {
            var from = src.IndexOf(start, StringComparison.Ordinal);
            Assert.True(from >= 0, "Missing: " + start);
            var to = src.IndexOf(end, from, StringComparison.Ordinal);
            Assert.True(to > from, "Missing end marker after: " + start);
            return src.Substring(from, to - from);
        }

        [Fact]
        public void UpdateModelsButton_UsesExplicitColors_ForServer2025()
        {
            var init = Slice(SettingsFormSource, "_btnUpdateModels = new Button", "};");

            Assert.Contains("\"Update Models\"", init);
            Assert.Contains("ForeColor = Color.Black", init);
            Assert.Contains("BackColor = SystemColors.ButtonFace", init);
            Assert.Contains("UseVisualStyleBackColor = false", init);
        }

        [Fact]
        public void UpdateModels_IsAdminGated_AndSwapsInTheFetchedCatalog()
        {
            var handler = Slice(SettingsFormSource, "private async void BtnUpdateModels_Click", "\n        }\n");

            Assert.Contains("if (!_authenticated) return;", handler);
            Assert.Contains("new ModelCatalogUpdater(", handler);
            Assert.Contains("Config.ModelCatalogClientVersion", handler);
            Assert.Contains("Config.ModelsNamedInConfig", handler);   // configured models survive a filtered list
            Assert.Contains("Config.ModelCatalog = result.Catalog;", handler);
            Assert.Contains("finally", handler);
            // Values the old catalog rejected are re-applied, and open panes refresh.
            var swap = handler.IndexOf("Config.ModelCatalog = result.Catalog;", StringComparison.Ordinal);
            Assert.True(handler.IndexOf("Config.ReloadConfigFiles();", swap, StringComparison.Ordinal) > swap);
            Assert.True(handler.IndexOf("Config.NotifyAiSettingsChanged();", swap, StringComparison.Ordinal) > swap);
        }

        [Fact]
        public void SaveAiSettings_TellsOpenPanesToRefresh()
        {
            var handler = Slice(SettingsFormSource, "private void BtnSaveAiSettings_Click", "\n        }\n");

            Assert.True(handler.IndexOf("Config.NotifyAiSettingsChanged();", StringComparison.Ordinal)
                > handler.IndexOf("Config.SaveConfig();", StringComparison.Ordinal));
        }

        [Theory]
        [InlineData("Chat", "ChatController.cs")]
        [InlineData("InboxCopilot", "InboxCopilotController.cs")]
        [InlineData("InboxReports", "InboxReportsController.cs")]
        [InlineData("Variants", "VariantsController.cs")]
        public void EveryPaneWithAReasoningDropdown_RefreshesOnSettingsChanges(string folder, string file)
        {
            var src = File.ReadAllText(FindSourceFile("OutlookAI", "TaskPane", folder, file));

            Assert.Contains("Config.AiSettingsChanged += ", src);
            Assert.Contains("Config.AiSettingsChanged -= ", src);   // no leak from the static event
        }

        [Fact]
        public void ModelDropdown_ComesFromTheCatalog_NotAHardcodedList()
        {
            var src = SettingsFormSource;
            var populate = Slice(src, "private void PopulateModelChoices", "\n        }\n");

            Assert.Contains("Config.ModelCatalog", populate);
            Assert.Contains(".ListedSlugs", populate);
            Assert.DoesNotContain("Config.AvailableModels", src);
        }
    }
}
