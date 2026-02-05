using System;
using System.IO;
using System.Xml.Linq;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION
        // These defaults are empty - configure via Settings panel or
        // edit before building for pre-configured deployment
        // ============================================================

        // Your Anthropic API Key (get one at https://console.anthropic.com)
        public static string ApiKey { get; set; } = "";

        // OpenAI API Key for Whisper speech-to-text (get one at https://platform.openai.com)
        // Optional - voice input will be disabled if not set
        public static string OpenAIApiKey { get; set; } = "";

        // Admin password for settings panel (set your own password)
        public static string AdminPassword { get; set; } = "admin";

        // Default Claude model
        public static string Model { get; set; } = "claude-sonnet-4-20250514";

        // Max tokens for responses
        public static int MaxTokens { get; set; } = 2048;

        // ============================================================
        // END CONFIGURATION
        // ============================================================

        private static readonly string ConfigFilePath = Path.Combine(
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
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var doc = XDocument.Load(ConfigFilePath);
                    var root = doc.Root;

                    if (root.Element("ApiKey") != null)
                        ApiKey = root.Element("ApiKey").Value;
                    if (root.Element("OpenAIApiKey") != null)
                        OpenAIApiKey = root.Element("OpenAIApiKey").Value;
                    if (root.Element("AdminPassword") != null)
                        AdminPassword = root.Element("AdminPassword").Value;
                    if (root.Element("Model") != null)
                        Model = root.Element("Model").Value;
                    if (root.Element("MaxTokens") != null)
                        MaxTokens = int.Parse(root.Element("MaxTokens").Value);
                }
            }
            catch
            {
                // Use defaults if config file is invalid
            }
        }

        public static void SaveConfig()
        {
            try
            {
                var dir = Path.GetDirectoryName(ConfigFilePath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var doc = new XDocument(
                    new XElement("Config",
                        new XElement("ApiKey", ApiKey),
                        new XElement("OpenAIApiKey", OpenAIApiKey),
                        new XElement("AdminPassword", AdminPassword),
                        new XElement("Model", Model),
                        new XElement("MaxTokens", MaxTokens)
                    )
                );

                doc.Save(ConfigFilePath);
            }
            catch
            {
                // Silently fail if can't save
            }
        }

        public static readonly string[] AvailableModels = new[]
        {
              "claude-sonnet-4-20250514",
              "claude-opus-4-5-20251101",
              "claude-haiku-3-5-20241022"
          };
    }
}