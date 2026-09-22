using System;
using System.IO;
using OutlookAI;
using OutlookAI.Services.Models;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests
{
    [Collection("Config")]
    public class ConfigTests : IDisposable
    {
        // Built-in catalog + a clock before the gpt-5.5 retirement, restored after each test.
        private readonly ConfigStateScope _scope = new ConfigStateScope();

        public void Dispose()
        {
            _scope.Dispose();
        }

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
            // Default model = the catalog's top listed, non-retired model.
            Assert.Equal("gpt-6-astra", Config.DefaultModel);
            Assert.Equal("gpt-6-astra", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
            Assert.Equal("", Config.ModelCatalogClientVersion);
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

            Assert.Equal(Config.DefaultModel, Config.Model);
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
        public void ReasoningEffortsForModel_ModelOutsideTheCatalog_OffersOnlyNone()
        {
            Assert.Equal(new[] { "None" }, Config.ReasoningEffortsForModel("gpt-4.1-nano"));
            Assert.Equal(new[] { "None" }, Config.ReasoningEffortsForModel("gpt-5.5-pro"));
            Assert.Contains("High", Config.ReasoningEffortsForModel("gpt-5.5"));
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt55_ExcludesMinimal_IncludesXHigh()
        {
            // Backend ground truth: 'minimal' and 'max' are rejected for gpt-5.5.
            var efforts = Config.ReasoningEffortsForModel("gpt-5.5");
            Assert.DoesNotContain("Minimal", efforts);
            Assert.Contains("XHigh", efforts);
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh" }, efforts);
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt6Astra_OffersMax_NotTheClientOnlyUltra()
        {
            var efforts = Config.ReasoningEffortsForModel("gpt-6-astra");
            Assert.Contains("Max", efforts);
            Assert.DoesNotContain("Ultra", efforts);
        }

        [Fact]
        public void ReasoningEffortsForModel_FollowsARefreshedCatalog()
        {
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-7", 1, new[] { "low", "extreme" }));

            Assert.Equal(new[] { "None", "Low" }, Config.ReasoningEffortsForModel("gpt-7"));
            Config.ModelCatalog = new ModelCatalog(Config.ModelCatalog.Models, new[] { "low", "extreme" },
                ModelCatalog.SourceChatGpt, TestCatalogs.FixedNow, "0.160.0");
            Assert.Equal(new[] { "None", "Low", "Extreme" }, Config.ReasoningEffortsForModel("gpt-7"));
        }

        [Fact]
        public void LoadConfigFromPaths_RecordsEveryModelNamedInConfig_EvenRejectedOnes()
        {
            // Update Models keeps these alive when a filtered list omits them.
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><Model>gpt-5.5</Model></Config>");
            File.WriteAllText(s, "<Config><Model> gpt-7 </Model></Config>");
            File.WriteAllText(u, "<Config><Model>GPT-6-SOL</Model></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);

                Assert.Equal(new[] { "gpt-5.5", "gpt-7", "GPT-6-SOL" }, Config.ModelsNamedInConfig);
                Assert.Equal("gpt-6-sol", Config.Model);
            }
            finally
            {
                File.Delete(g);
                File.Delete(s);
                File.Delete(u);
            }
        }

        [Fact]
        public void NotifyAiSettingsChanged_RunsEveryHandler_EvenIfOneThrows()
        {
            var calls = 0;
            EventHandler failing = (s, e) => { calls++; throw new InvalidOperationException("pane disposed"); };
            EventHandler working = (s, e) => calls++;
            Config.AiSettingsChanged += failing;
            Config.AiSettingsChanged += working;
            try
            {
                Config.NotifyAiSettingsChanged();
                Assert.Equal(2, calls);
            }
            finally
            {
                Config.AiSettingsChanged -= failing;
                Config.AiSettingsChanged -= working;
            }
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
            File.WriteAllText(u, "<Config><Model>gpt-6-sol</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-6-sol", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresModelNotInCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>nonexistent-gpt-9</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            // Falls back to the catalog default since the requested model
            // isn't in the catalog.
            Assert.Equal(Config.DefaultModel, Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_AcceptsModelsFromARefreshedCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>gpt-7</Model></Config>");

            Config.LoadConfigFromPaths(g, u);
            Assert.NotEqual("gpt-7", Config.Model);   // built-in catalog doesn't know it

            // gpt-7 is deliberately not the catalog default, so only the config
            // value can put it there.
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-top", 1, new[] { "low" }),
                TestCatalogs.Entry("gpt-7", 2, new[] { "low" }),
                TestCatalogs.Entry("gpt-reserve", 3, new[] { "low" }, listed: false));
            Config.LoadConfigFromPaths(g, u);
            Assert.Equal("gpt-7", Config.Model);

            // Hidden catalog models are still valid when set explicitly.
            File.WriteAllText(g, "<Config><Model>gpt-reserve</Model></Config>");
            Config.LoadConfigFromPaths(g, u);
            Assert.Equal("gpt-reserve", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_MatchesModelCaseInsensitively_AndStoresTheCanonicalSlug()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>  GPT-5.6-SOL </Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-5.6-sol", Config.Model);
        }

        [Theory]
        [InlineData("max", "Max")]
        [InlineData("XHIGH", "XHigh")]
        [InlineData("medium", "Medium")]
        public void LoadConfigFromPaths_AcceptsCatalogEfforts_WithCanonicalCasing(string raw, string expected)
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>" + raw + "</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(expected, Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("Ultra")]     // Codex client-only mode; the server rejects it
        [InlineData("Minimal")]   // no current model accepts it
        public void LoadConfigFromPaths_IgnoresEffortsNoModelOffers(string raw)
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>" + raw + "</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("None", Config.ReasoningEffort);
        }

        [Fact]
        public void LoadConfigFromPaths_ModelCatalogClientVersion_IsServerAuthoritative()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ModelCatalogClientVersion> 0.160.0 </ModelCatalogClientVersion></Config>");
            File.WriteAllText(u, "<Config><ModelCatalogClientVersion>0.170.0</ModelCatalogClientVersion></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("0.160.0", Config.ModelCatalogClientVersion);
        }

        [Fact]
        public void EffectiveModel_SwitchesToTheUpgradeTarget_OnceTheModelRetires()
        {
            Config.Model = "gpt-5.5";
            var retirement = new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero);

            Config.Clock = () => retirement.AddMinutes(-1);
            Assert.Equal("gpt-5.5", Config.EffectiveModel);

            Config.Clock = () => retirement;
            Assert.Equal("gpt-5.6-sol", Config.EffectiveModel);
            Assert.Equal("gpt-5.5", Config.Model);   // the saved choice is untouched
        }

        [Fact]
        public void EffectiveModel_UsesTheDefault_WhenTheModelLeavesTheCatalog()
        {
            Config.Model = "gpt-5.5";
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-7", 1, new[] { "low" }),
                TestCatalogs.Entry("gpt-6-sol", 2, new[] { "low" }));

            Assert.Equal("gpt-7", Config.EffectiveModel);
            Assert.Equal("gpt-7", Config.DefaultModel);
        }

        [Fact]
        public void ResetDefaults_KeepsTheLoadedCatalog()
        {
            var refreshed = TestCatalogs.Catalog(TestCatalogs.Entry("gpt-7", 1, new[] { "low" }));
            Config.ModelCatalog = refreshed;

            Config.ResetDefaults();

            Assert.Same(refreshed, Config.ModelCatalog);
            Assert.Equal("gpt-7", Config.Model);
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
                // Ceiling shares the interactive export cap (#12.1) so neither
                // Excel path can exceed BulkExportRowCap.Max (10,000).
                Assert.Equal(10000, Config.MaxBulkExportRows);  // ceiling
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
