using System;
using System.IO;
using System.IO.Compression;
using Newtonsoft.Json;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// version.json shipped inside every Release ZIP and copied into
    /// C:\Program Files\OutlookAI\ by the installer.
    /// </summary>
    public sealed class UpdateManifest
    {
        public const string DevSentinel = "(dev build)";

        [JsonProperty("tag")]
        public string Tag { get; set; }

        [JsonProperty("commit")]
        public string Commit { get; set; }

        [JsonProperty("build_date")]
        public DateTimeOffset BuildDate { get; set; }

        [JsonProperty("repo")]
        public string Repo { get; set; }

        [JsonIgnore]
        public bool IsDevBuild => string.IsNullOrEmpty(Tag) || Tag == DevSentinel;

        public static UpdateManifest LoadFromInstallDir()
            => LoadFromFile(UpdatePaths.InstalledVersionJson);

        public static UpdateManifest LoadFromFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return Dev();
                var json = File.ReadAllText(path);
                return Parse(json);
            }
            catch { return Dev(); }
        }

        public static UpdateManifest LoadFromZip(string zipPath)
        {
            try
            {
                using (var archive = ZipFile.OpenRead(zipPath))
                {
                    var entry = archive.GetEntry("version.json");
                    if (entry == null) return Dev();
                    using (var reader = new StreamReader(entry.Open()))
                    {
                        return Parse(reader.ReadToEnd());
                    }
                }
            }
            catch { return Dev(); }
        }

        private static UpdateManifest Parse(string json)
        {
            var parsed = JsonConvert.DeserializeObject<UpdateManifest>(json);
            if (parsed == null || string.IsNullOrWhiteSpace(parsed.Tag)) return Dev();
            return parsed;
        }

        private static UpdateManifest Dev() => new UpdateManifest { Tag = DevSentinel };
    }
}
