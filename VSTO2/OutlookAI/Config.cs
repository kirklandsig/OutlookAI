using System;
using System.IO;
using System.Xml.Linq;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION DEFAULTS
        // These are overridden by global config (Program Files) and
        // then by per-user config (AppData). After deploying, edit
        // C:\Program Files\OutlookAI\config.xml to change for all users.
        // ============================================================

        public static string ApiKey { get; set; } = "";
        public static string OpenAIApiKey { get; set; } = "";
        public static string AdminPassword { get; set; } = "admin";
        public static string Model { get; set; } = "claude-opus-4-6";
        public static string WhisperModel { get; set; } = "gpt-4o-transcribe";
        public static int MaxTokens { get; set; } = 2048;

        // ============================================================
        // END CONFIGURATION
        // ============================================================

        // Global config: admin-controlled, applies to all users
        private static readonly string GlobalConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OutlookAI",
            "config.xml"
        );

        // Per-user config: overrides global if user has saved settings
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
            // Load order: hardcoded defaults -> global config -> per-user config
            LoadFromFile(GlobalConfigFilePath);
            LoadFromFile(UserConfigFilePath);
        }

        private static void LoadFromFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var doc = XDocument.Load(filePath);
                    var root = doc.Root;

                    if (root.Element("ApiKey") != null)
                        ApiKey = root.Element("ApiKey").Value;
                    if (root.Element("OpenAIApiKey") != null)
                        OpenAIApiKey = root.Element("OpenAIApiKey").Value;
                    if (root.Element("AdminPassword") != null)
                        AdminPassword = root.Element("AdminPassword").Value;
                    if (root.Element("Model") != null)
                        Model = root.Element("Model").Value;
                    if (root.Element("WhisperModel") != null)
                        WhisperModel = root.Element("WhisperModel").Value;
                    if (root.Element("MaxTokens") != null)
                        MaxTokens = int.Parse(root.Element("MaxTokens").Value);
                }
            }
            catch
            {
                // Skip if file is missing or invalid
            }
        }

        public static void SaveConfig()
        {
            try
            {
                var dir = Path.GetDirectoryName(UserConfigFilePath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var doc = new XDocument(
                    new XElement("Config",
                        new XElement("ApiKey", ApiKey),
                        new XElement("OpenAIApiKey", OpenAIApiKey),
                        new XElement("AdminPassword", AdminPassword),
                        new XElement("Model", Model),
                        new XElement("WhisperModel", WhisperModel),
                        new XElement("MaxTokens", MaxTokens)
                    )
                );

                doc.Save(UserConfigFilePath);
            }
            catch
            {
                // Silently fail if can't save
            }
        }

        public static readonly string[] AvailableModels = new[]
        {
              "claude-sonnet-4-6",
              "claude-opus-4-6",
              "claude-haiku-4-5-20251001"
          };
    }
}