using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Append-only JSON log of update activity. Caps at 50 entries; oldest
    /// dropped on overflow. Tolerant of missing or malformed files (returns
    /// empty list).
    /// </summary>
    public sealed class UpdateHistoryLog
    {
        public sealed class Entry
        {
            [JsonProperty("ts")]      public DateTimeOffset Ts { get; set; }
            [JsonProperty("action")]  public string Action { get; set; }
            [JsonProperty("result")]  public string Result { get; set; }
            [JsonProperty("tag")]     public string Tag { get; set; }
            [JsonProperty("details")] public string Details { get; set; }
        }

        public const int MaxEntries = 50;

        private readonly string _path;
        private readonly object _gate = new object();

        public UpdateHistoryLog() : this(UpdatePaths.HistoryLog) { }

        public UpdateHistoryLog(string path)
        {
            _path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public void Append(string action, string result, string tag, string details)
        {
            lock (_gate)
            {
                var entries = ReadAll();
                entries.Add(new Entry
                {
                    Ts = DateTimeOffset.UtcNow,
                    Action = action ?? "",
                    Result = result ?? "",
                    Tag = tag ?? "",
                    Details = details ?? "",
                });
                while (entries.Count > MaxEntries) entries.RemoveAt(0);

                var dir = Path.GetDirectoryName(_path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var json = JsonConvert.SerializeObject(entries, Formatting.Indented);
                File.WriteAllText(_path, json);
            }
        }

        public List<Entry> ReadAll()
        {
            lock (_gate)
            {
                try
                {
                    if (!File.Exists(_path)) return new List<Entry>();
                    var json = File.ReadAllText(_path);
                    var parsed = JsonConvert.DeserializeObject<List<Entry>>(json);
                    return parsed ?? new List<Entry>();
                }
                catch { return new List<Entry>(); }
            }
        }
    }
}
