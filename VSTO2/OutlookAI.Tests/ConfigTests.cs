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
            Assert.Equal(65536, Config.MaxTokens);
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
                + "<MaxTokens>65536</MaxTokens>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<AdminPassword>userpass</AdminPassword>"
                + "<Model>claude-opus-4-6</Model>"
                + "<MaxTokens>2048</MaxTokens>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("userpass", Config.AdminPassword);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal(65536, Config.MaxTokens);
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
        }
    }
}
