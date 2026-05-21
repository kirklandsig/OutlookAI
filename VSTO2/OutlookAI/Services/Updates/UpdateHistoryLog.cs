using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Append-only JSON log of update activity. Caps at 50 entries; oldest
    /// dropped on overflow. <see cref="ReadAll"/> is tolerant of missing or
    /// malformed files (returns empty list). <see cref="Append"/> may throw
    /// on disk I/O failures — callers should treat it as best-effort and
    /// swallow exceptions so logging never breaks an update flow.
    /// </summary>
    /// <remarks>
    /// Thread-safe within a single instance via an internal lock. NOT safe
    /// across multiple <see cref="UpdateHistoryLog"/> instances pointing at
    /// the same path, nor across processes. This is acceptable because the
    /// updater workflow is admin-gated and user-initiated from
    /// <c>SettingsForm</c>; there is no background thread or timer that
    /// would race with it.
    /// </remarks>
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
                var tmp = _path + ".tmp";
                try
                {
                    File.WriteAllText(tmp, json);
                    if (File.Exists(_path))
                        File.Replace(tmp, _path, null);
                    else
                        File.Move(tmp, _path);
                }
                catch
                {
                    try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                    throw;
                }
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
