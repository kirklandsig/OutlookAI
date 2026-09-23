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
            // Controls the admin hasn't touched follow the reloaded settings, so a
            // later Save doesn't write the values they showed before back.
            Assert.True(handler.IndexOf("ShowReloadedSettings();", swap, StringComparison.Ordinal) > swap);
        }

        [Fact]
        public void SaveAiSettings_TellsOpenPanesToRefresh()
        {
            var handler = Slice(SettingsFormSource, "private void BtnSaveAiSettings_Click", "\n        }\n");

            Assert.True(handler.IndexOf("Config.NotifyAiSettingsChanged();", StringComparison.Ordinal)
                > handler.IndexOf("SaveSettings(", StringComparison.Ordinal));
        }

        // Settings saves for every user: only what was changed in the dialog, so
        // another admin's newer settings aren't overwritten.
        [Theory]
        [InlineData("private void BtnSaveAiSettings_Click", "SaveSettings(changed.ToArray(), _toolsBaseline)")]
        [InlineData("private void BtnSavePassword_Click", "SaveSettings(new[] { \"AdminPassword\" }, null)")]
        public void SaveHandlers_SaveOnlyWhatChanged(string handlerStart, string save)
        {
            Assert.Contains(save, Slice(SettingsFormSource, handlerStart, "\n        }\n"));
        }

        [Fact]
        public void Settings_OpensOnWhatIsSavedNow()
        {
            var pane = File.ReadAllText(FindSourceFile("OutlookAI", "TaskPane", "AITaskPane.cs")).Replace("\r\n", "\n");
            var open = Slice(pane, "private static SettingsForm OpenSettings", "\n        }\n");

            var reload = open.IndexOf("Config.ReloadConfigFiles();", StringComparison.Ordinal);
            Assert.True(reload >= 0 && reload < open.IndexOf("new SettingsForm", StringComparison.Ordinal));
            // Every way the pane opens Settings goes through it.
            Assert.Equal(2, pane.Split(new[] { "new SettingsForm" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(2, open.Split(new[] { "new SettingsForm" }, StringSplitOptions.None).Length - 1);
        }

        // Settings saves for every user; a failed save must not look saved.
        [Theory]
        [InlineData("private void BtnSaveAiSettings_Click")]
        [InlineData("private void BtnSavePassword_Click")]
        public void SaveHandlers_StopAndSayWhy_WhenTheSaveFails(string handlerStart)
        {
            var handler = Slice(SettingsFormSource, handlerStart, "\n        }\n");

            // This session's settings are captured before the handler changes them
            // and put back when the save fails.
            var snapshot = handler.IndexOf("var before = new SettingsBeforeSave();", StringComparison.Ordinal);
            var firstChange = handler.IndexOf("Config.", snapshot < 0 ? 0 : snapshot, StringComparison.Ordinal);
            var save = handler.IndexOf("var saveError = SaveSettings(", StringComparison.Ordinal);
            var bail = System.Text.RegularExpressions.Regex.Match(handler,
                @"if \(saveError != null\)\s*\{\s*before\.Restore\(\);\s*ReportSaveFailed\(this, saveError\);\s*return;");
            Assert.True(snapshot >= 0 && firstChange > snapshot && save > snapshot && bail.Success && bail.Index > save);
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
            var populate = Slice(src, "private static List<string> ModelChoices", "\n        }\n");

            Assert.Contains("Config.ModelCatalog", populate);
            Assert.Contains(".ListedSlugs", populate);
            Assert.DoesNotContain("Config.AvailableModels", src);
        }
    }
}
