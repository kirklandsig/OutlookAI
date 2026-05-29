using System;
using System.IO;
using OutlookAI;
using Xunit;

namespace OutlookAI.Tests
{
    public class ConfigTests
    {
        private static (string global, string user) MakeTempPaths()
        {
            var dir = Path.Combine(Path.GetTempPath(),
                "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            return (Path.Combine(dir, "global.xml"), Path.Combine(dir, "user.xml"));
        }

        [Fact]
        public void LoadConfigFromPaths_UsesV2Defaults_WhenFilesAreMissing()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("admin", Config.AdminPassword);
            Assert.Equal(@"C:\ProgramData\OutlookAI\auth.json", Config.CodexAuthPath);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
        }

        [Fact]
        public void LoadConfigFromPaths_PerUserOverridesAdminPasswordOnly()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<AdminPassword>server</AdminPassword>"
                + "<CodexAuthPath>C:\\ProgramData\\OutlookAI\\auth.json</CodexAuthPath>"
                + "<Model>gpt-5.5</Model>"
                + "<VoiceModel>gpt-realtime-1.5</VoiceModel>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<AdminPassword>userpass</AdminPassword>"
                + "<Model>claude-opus-4-6</Model>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("userpass", Config.AdminPassword);
            // Server-authoritative Model is not overridden by per-user;
            // unknown Claude-era model names also do not override it.
            Assert.Equal("gpt-5.5", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresLegacyV1Fields()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ApiKey>anthropic-key</ApiKey>"
                + "<OpenAIApiKey>openai-key</OpenAIApiKey>"
                + "<WhisperModel>whisper-1</WhisperModel>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesReasoningEffortFromGlobal()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("High", Config.ReasoningEffort);
        }

        [Fact]
        public void LoadConfigFromPaths_UserOverridesReasoningEffortAndWriteTools()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ReasoningEffort>Medium</ReasoningEffort>"
                + "<WriteToolsEnabled>true</WriteToolsEnabled>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<ReasoningEffort>Low</ReasoningEffort>"
                + "<WriteToolsEnabled>false</WriteToolsEnabled>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Low", Config.ReasoningEffort);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresUnknownReasoningEffort()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>Extreme</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            // Unknown value -> fall back to default "None".
            Assert.Equal("None", Config.ReasoningEffort);
        }

        [Fact]
        public void ReasoningEffortsForModel_RestrictsForNonReasoningModels()
        {
            Assert.Equal(new[] { "None" }, Config.ReasoningEffortsForModel("gpt-4.1-nano"));
            Assert.Equal(new[] { "None" }, Config.ReasoningEffortsForModel("gpt-4.1-mini"));
            Assert.Contains("High", Config.ReasoningEffortsForModel("gpt-5.5"));
            Assert.Contains("Medium", Config.ReasoningEffortsForModel("gpt-5.5-pro"));
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt55_ExcludesMinimal_IncludesXHigh()
        {
            // Backend ground truth (captured from a real Codex error):
            // 'minimal' is not supported with gpt-5.5; the supported set
            // is none/low/medium/high/xhigh.
            var efforts = Config.ReasoningEffortsForModel("gpt-5.5");
            Assert.DoesNotContain("Minimal", efforts);
            Assert.Contains("XHigh", efforts);
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh" }, efforts);
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt54_IncludesMinimalAndXHigh()
        {
            // Per OpenAI docs: gpt-5.4 family supports the full enum.
            var efforts = Config.ReasoningEffortsForModel("gpt-5.4");
            Assert.Contains("Minimal", efforts);
            Assert.Contains("XHigh", efforts);
        }

        [Fact]
        public void AvailableReasoningEfforts_MatchesOpenAiPublicEnum()
        {
            // Per OpenAI 'Reasoning models' guide:
            // "Supported values are model-dependent and can include
            //  'none', 'minimal', 'low', 'medium', 'high', and 'xhigh'."
            var expected = new[] { "None", "Minimal", "Low", "Medium", "High", "XHigh" };
            Assert.Equal(expected, Config.AvailableReasoningEfforts);
        }

        [Fact]
        public void AvailableModels_ContainsExpectedCatalog()
        {
            Assert.Contains("gpt-5.5", Config.AvailableModels);
            Assert.Contains("gpt-5.5-pro", Config.AvailableModels);
            Assert.Contains("gpt-4.1-mini", Config.AvailableModels);
            Assert.Contains("gpt-5.3-codex", Config.AvailableModels);
        }

        [Fact]
        public void Defaults_AreV2()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("None", Config.ReasoningEffort);
            Assert.True(Config.WriteToolsEnabled);
            // Default: all four write tools enabled.
            Assert.Equal(4, Config.EnabledWriteTools.Count);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
            Assert.Contains("outlook_mark_as_read", Config.EnabledWriteTools);
            Assert.Contains("outlook_flag_message", Config.EnabledWriteTools);
            Assert.Contains("outlook_set_category", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesEnabledWriteToolsFromCSV()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<EnabledWriteTools>outlook_create_draft, outlook_set_category</EnabledWriteTools>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(2, Config.EnabledWriteTools.Count);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
            Assert.Contains("outlook_set_category", Config.EnabledWriteTools);
            Assert.DoesNotContain("outlook_mark_as_read", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_FiltersUnknownWriteToolNames()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<EnabledWriteTools>outlook_create_draft,outlook_send_now,bogus_tool</EnabledWriteTools>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            // Unknown names are silently dropped (forward-compat with admin
            // configs that reference future or removed tools).
            Assert.Single(Config.EnabledWriteTools);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesModelFromUserOverride_WhenInCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>gpt-5.5</Model></Config>");
            File.WriteAllText(u, "<Config><Model>gpt-5.5-pro</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-5.5-pro", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresModelNotInCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>nonexistent-gpt-9</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            // Should fall back to the v2 default since the requested model
            // isn't in AvailableModels.
            Assert.Equal("gpt-5.5", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_SharedDefaultsAppliedWhenUserAbsent()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            // No user file
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Medium", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(s)) File.Delete(s);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_UserOverridesSharedDefaults()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            File.WriteAllText(u, "<Config><ReasoningEffort>Low</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Low", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(s)) File.Delete(s);
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_SharedConfigPathNull_TreatedAsAbsent()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(u, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal("High", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_GlobalAndSharedAndUser_UserBeatsSharedBeatsGlobal()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            File.WriteAllText(u, "<Config><ReasoningEffort>Low</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Low", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(g)) File.Delete(g);
                if (File.Exists(s)) File.Delete(s);
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void MaxBulkExportRows_DefaultsTo2000()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
            Assert.Equal(2000, Config.MaxBulkExportRows);
        }

        [Fact]
        public void MaxBulkExportRows_LoadsFromGlobalConfig()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><MaxBulkExportRows>500</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(500, Config.MaxBulkExportRows);
            }
            finally { if (File.Exists(g)) File.Delete(g); }
        }

        [Fact]
        public void MaxBulkExportRows_ClampsToFloorAndCeiling()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(g, "<Config><MaxBulkExportRows>0</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(1, Config.MaxBulkExportRows);   // floor

                File.WriteAllText(g, "<Config><MaxBulkExportRows>999999</MaxBulkExportRows></Config>");
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(50000, Config.MaxBulkExportRows);  // ceiling
            }
            finally { if (File.Exists(g)) File.Delete(g); }
        }

        [Fact]
        public void MaxBulkExportRows_NotUserOverridable()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
            File.WriteAllText(u, "<Config><MaxBulkExportRows>3000</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(750, Config.MaxBulkExportRows);  // user value ignored
            }
            finally { if (File.Exists(g)) File.Delete(g); if (File.Exists(u)) File.Delete(u); }
        }
    }
}
