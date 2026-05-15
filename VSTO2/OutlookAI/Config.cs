using System;
using System.IO;
using System.Xml.Linq;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION DEFAULTS (v2 - ChatGPT OAuth)
        // Server-authoritative fields (CodexAuthPath, Model, MaxTokens)
        // load from defaults -> global config (Program Files). Per-user
        // AppData config may override only AdminPassword. Legacy v1
        // elements (ApiKey, OpenAIApiKey, WhisperModel, TranscribeModel,
        // and Claude model names) are ignored if encountered.
        // ============================================================

        public const string DefaultModel = "gpt-5.5";
        public const string DefaultVoiceModel = "gpt-realtime-1.5";
        public const string DefaultCodexAuthPath = @"C:\ProgramData\OutlookAI\auth.json";
        public const int DefaultMaxTokens = 65536;

        public static string AdminPassword { get; set; } = "admin";
        public static string CodexAuthPath { get; set; } = DefaultCodexAuthPath;
        public static string Model { get; set; } = DefaultModel;
        public static string VoiceModel { get; set; } = DefaultVoiceModel;
        public static int MaxTokens { get; set; } = DefaultMaxTokens;

        // ============================================================
        // END CONFIGURATION
        // ============================================================

        // Global config: admin-controlled, applies to all users on this server
        private static readonly string GlobalConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OutlookAI",
            "config.xml"
        );

        // Per-user config: may override AdminPassword only
        private static readonly string UserConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookAI",
            "config.xml"
        );

        static Config()
        {
            LoadConfig();
        }

        public static void LoadConfig()
        {
            LoadConfigFromPaths(GlobalConfigFilePath, UserConfigFilePath);
        }

        // Test seam: explicit paths so we don't touch Program Files / AppData
        // during unit tests.
        public static void LoadConfigFromPaths(string globalConfigPath, string userConfigPath)
        {
            ResetDefaults();
            LoadFromFile(globalConfigPath, allowServerFields: true);
            LoadFromFile(userConfigPath, allowServerFields: false);
        }

        public static void ResetDefaults()
        {
            AdminPassword = "admin";
            CodexAuthPath = DefaultCodexAuthPath;
            Model = DefaultModel;
            VoiceModel = DefaultVoiceModel;
            MaxTokens = DefaultMaxTokens;
        }

        private static void LoadFromFile(string filePath, bool allowServerFields)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return;
                }

                var doc = XDocument.Load(filePath);
                var root = doc.Root;
                if (root == null)
                {
                    return;
                }

                // AdminPassword is the only field per-user config can change.
                var adminPassword = root.Element("AdminPassword");
                if (adminPassword != null && !string.IsNullOrEmpty(adminPassword.Value))
                {
                    AdminPassword = adminPassword.Value;
                }

                if (!allowServerFields)
                {
                    return;
                }

                var codexAuthPath = root.Element("CodexAuthPath");
                if (codexAuthPath != null && !string.IsNullOrWhiteSpace(codexAuthPath.Value))
                {
                    CodexAuthPath = codexAuthPath.Value;
                }

                var model = root.Element("Model");
                if (model != null && !string.IsNullOrWhiteSpace(model.Value))
                {
                    Model = model.Value;
                }

                var voiceModel = root.Element("VoiceModel");
                if (voiceModel != null && !string.IsNullOrWhiteSpace(voiceModel.Value))
                {
                    VoiceModel = voiceModel.Value;
                }

                var maxTokens = root.Element("MaxTokens");
                if (maxTokens != null && int.TryParse(maxTokens.Value, out var parsed) && parsed > 0)
                {
                    MaxTokens = parsed;
                }
            }
            catch
            {
                // Skip if file is missing or invalid; defaults stay in place.
            }
        }

        public static void SaveConfig()
        {
            try
            {
                var dir = Path.GetDirectoryName(UserConfigFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Per-user config persists only AdminPassword. Server-authoritative
                // fields are owned by the global config in Program Files.
                var doc = new XDocument(
                    new XElement("Config",
                        new XElement("AdminPassword", AdminPassword)
                    )
                );

                doc.Save(UserConfigFilePath);
            }
            catch
            {
                // Silently fail if can't save (e.g. read-only AppData).
            }
        }
    }
}
