using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// Parsing for the Codex model-catalog response and for OutlookAI's trimmed
    /// models.json cache. Slugs and efforts end up in config.xml, the Settings UI
    /// and request bodies, so both the network response and a cache file in the
    /// shared ProgramData folder are validated as untrusted input.
    /// </summary>
    public static class ModelCatalogJson
    {
        public const int CacheSchemaVersion = 1;
        public const int MaxModels = 200;
        public const int MaxEffortsPerModel = 16;
        public const int MaxDisplayNameLength = 100;
        public const int MaxDescriptionLength = 300;
        public const int DefaultPriority = 1000;

        private static readonly Regex SlugPattern =
            new Regex(@"^[A-Za-z0-9][A-Za-z0-9._:\-]{0,99}\z", RegexOptions.CultureInvariant);

        private static readonly Regex EffortPattern =
            new Regex(@"^[a-z][a-z0-9_\-]{0,31}\z", RegexOptions.CultureInvariant);

        private static readonly Regex QuotedEffort =
            new Regex(@"'([a-z][a-z0-9_\-]{0,31})'", RegexOptions.CultureInvariant);

        private static readonly Regex AccountIdPattern =
            new Regex(@"^[A-Za-z0-9_\-]{1,100}\z", RegexOptions.CultureInvariant);

        private static readonly string[] ClientVersionSources = { "config", "github", "cache", "built-in" };

        // Anything outside this range is corrupt, and far-future years can't even
        // be formatted under some calendars (e.g. UmAlQura stops at 2077).
        private static readonly DateTimeOffset EarliestDate = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
        private static readonly DateTimeOffset LatestDate = new DateTimeOffset(2200, 1, 1, 0, 0, 0, TimeSpan.Zero);

        public static bool IsValidSlug(string slug)
        {
            return slug != null && SlugPattern.IsMatch(slug);
        }

        /// <summary>
        /// Parses a <c>/backend-api/codex/models</c> body. Invalid entries are
        /// skipped; a body that is not <c>{"models":[...]}</c> throws
        /// <see cref="FormatException"/> (or <see cref="JsonException"/> for
        /// malformed JSON).
        /// </summary>
        public static List<ModelCatalogEntry> ParseModelsResponse(string json)
        {
            var root = Load(json) as JObject;
            if (root == null) throw new FormatException("Expected a JSON object.");
            var models = root["models"] as JArray;
            if (models == null) throw new FormatException("The response has no 'models' array.");
            return ParseEntries(models);
        }

        /// <summary>Reads a models.json cache; null when the file is not a usable schema-1 cache.</summary>
        public static ModelCatalog ParseCache(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try
            {
                var root = Load(json) as JObject;
                if (root == null || IntValue(root["schema_version"]) != CacheSchemaVersion) return null;

                var fetchedAt = ParseDate(StringValue(root["fetched_at"]));
                if (fetchedAt == null) return null;

                var models = ParseEntries(root["models"] as JArray);
                if (!models.Any(m => m.Listed)) return null;

                List<string> serverEfforts = null;
                var effortsArray = root["server_reasoning_efforts"] as JArray;
                if (effortsArray != null)
                {
                    serverEfforts = effortsArray
                        .Select(StringValue)
                        .Select(NormalizeEffortValue)
                        .Where(e => e != null)
                        .Distinct()
                        .ToList();
                }

                var clientVersionSource = StringValue(root["client_version_source"]);
                var accountId = StringValue(root["account_id"]);
                return new ModelCatalog(
                    models,
                    serverEfforts,
                    ModelCatalog.SourceChatGpt,
                    fetchedAt,
                    CodexClientVersion.Normalize(StringValue(root["client_version"])),
                    ClientVersionSources.Contains(clientVersionSource) ? clientVersionSource : null,
                    accountId != null && AccountIdPattern.IsMatch(accountId) ? accountId : null);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string SerializeCache(ModelCatalog catalog)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));

            var models = new JArray();
            foreach (var m in catalog.Models)
            {
                models.Add(new JObject(
                    new JProperty("slug", m.Slug),
                    new JProperty("display_name", m.DisplayName),
                    new JProperty("description", m.Description),
                    new JProperty("visibility", m.Listed ? "list" : "hide"),
                    new JProperty("priority", m.Priority),
                    new JProperty("supported_reasoning_levels",
                        new JArray(m.Efforts.Select(e => new JObject(new JProperty("effort", e))))),
                    new JProperty("upgrade", m.Upgrade == null
                        ? null
                        : new JObject(
                            new JProperty("model", m.Upgrade.Model),
                            new JProperty("retirement_at", FormatDate(m.Upgrade.RetirementAt))))));
            }

            var root = new JObject(
                new JProperty("schema_version", CacheSchemaVersion),
                new JProperty("source", catalog.Source),
                new JProperty("fetched_at", FormatDate(catalog.FetchedAt)),
                new JProperty("client_version", catalog.ClientVersion),
                new JProperty("client_version_source", catalog.ClientVersionSource),
                new JProperty("account_id", catalog.AccountId),
                new JProperty("server_reasoning_efforts",
                    catalog.ServerEfforts == null ? null : new JArray(catalog.ServerEfforts)),
                new JProperty("models", models));
            return root.ToString(Formatting.Indented);
        }

        /// <summary>
        /// Reads the server's global effort set out of the 400 that /responses
        /// returns for an invalid reasoning.effort, e.g.
        /// <c>Invalid value: 'x'. Supported values are: 'none', 'minimal', …, and 'max'.</c>
        /// Null for any other response.
        /// </summary>
        public static IReadOnlyList<string> ParseSupportedEffortsFromError(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;
            JObject error;
            try
            {
                error = (Load(body) as JObject)?["error"] as JObject;
            }
            catch (Exception)
            {
                return null;
            }
            if (error == null) return null;

            var param = StringValue(error["param"]);
            if (param != null && !string.Equals(param, "reasoning.effort", StringComparison.Ordinal)) return null;

            var message = StringValue(error["message"]);
            if (message == null) return null;
            var at = message.IndexOf("Supported values are:", StringComparison.OrdinalIgnoreCase);
            if (at < 0) return null;

            var values = QuotedEffort.Matches(message.Substring(at))
                .Cast<Match>()
                .Select(m => m.Groups[1].Value)
                .Distinct()
                .ToList();
            return values.Count >= 2 ? values.AsReadOnly() : null;
        }

        internal static string NormalizeEffortValue(string raw)
        {
            if (raw == null) return null;
            var effort = raw.Trim().ToLowerInvariant();
            return EffortPattern.IsMatch(effort) ? effort : null;
        }

        private static List<ModelCatalogEntry> ParseEntries(JArray array)
        {
            var result = new List<ModelCatalogEntry>();
            if (array == null) return result;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in array)
            {
                if (result.Count >= MaxModels) break;
                var entry = TryParseEntry(token as JObject);
                if (entry != null && seen.Add(entry.Slug)) result.Add(entry);
            }
            return result;
        }

        private static ModelCatalogEntry TryParseEntry(JObject o)
        {
            if (o == null) return null;
            var slug = StringValue(o["slug"])?.Trim();
            if (!IsValidSlug(slug)) return null;

            // Only an explicit non-"list" visibility hides a model; a catalog that
            // stops sending the field should not suddenly hide everything.
            var visibility = StringValue(o["visibility"]);
            var listed = visibility == null
                || string.Equals(visibility.Trim(), "list", StringComparison.OrdinalIgnoreCase);

            return new ModelCatalogEntry(
                slug,
                CleanText(StringValue(o["display_name"]), MaxDisplayNameLength),
                CleanText(StringValue(o["description"]), MaxDescriptionLength),
                listed,
                IntValue(o["priority"]) ?? DefaultPriority,
                ParseEfforts(o["supported_reasoning_levels"] as JArray),
                ParseUpgrade(o["upgrade"] as JObject));
        }

        private static List<string> ParseEfforts(JArray levels)
        {
            var result = new List<string>();
            if (levels == null) return result;
            foreach (var level in levels)
            {
                if (result.Count >= MaxEffortsPerModel) break;
                var raw = level is JObject obj ? StringValue(obj["effort"]) : StringValue(level);
                var effort = NormalizeEffortValue(raw);
                if (effort != null && !result.Contains(effort)) result.Add(effort);
            }
            return result;
        }

        private static ModelUpgradeInfo ParseUpgrade(JObject upgrade)
        {
            if (upgrade == null) return null;
            var model = StringValue(upgrade["model"])?.Trim();
            if (!IsValidSlug(model)) model = null;
            var retirementAt = ParseDate(StringValue(upgrade["retirement_at"]));
            return model == null && retirementAt == null ? null : new ModelUpgradeInfo(model, retirementAt);
        }

        // DateParseHandling.None keeps ISO timestamps as strings; Json.NET's default
        // turns them into local DateTime values before we can parse them as UTC.
        private static JToken Load(string json)
        {
            using (var reader = new JsonTextReader(new StringReader(json ?? "")) { DateParseHandling = DateParseHandling.None })
            {
                return JToken.ReadFrom(reader);
            }
        }

        private static string StringValue(JToken token)
        {
            return token != null && token.Type == JTokenType.String ? (string)token : null;
        }

        private static int? IntValue(JToken token)
        {
            if (token == null || token.Type != JTokenType.Integer) return null;
            try
            {
                var value = (long)token;
                if (value < int.MinValue || value > int.MaxValue) return null;
                return (int)value;
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private static DateTimeOffset? ParseDate(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            DateTimeOffset parsed;
            if (!DateTimeOffset.TryParse(
                raw.Trim(),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out parsed))
            {
                return null;
            }
            return parsed >= EarliestDate && parsed < LatestDate ? parsed : (DateTimeOffset?)null;
        }

        private static string FormatDate(DateTimeOffset? value)
        {
            return value.HasValue ? value.Value.ToString("o", CultureInfo.InvariantCulture) : null;
        }

        // Collapses whitespace/control characters (these strings land in WinForms
        // labels) and clips to a sane length without splitting a surrogate pair.
        private static string CleanText(string value, int maxLength)
        {
            if (value == null) return null;
            var sb = new StringBuilder(Math.Min(value.Length, maxLength + 1));
            var pendingSpace = false;
            foreach (var ch in value)
            {
                if (char.IsWhiteSpace(ch) || char.IsControl(ch))
                {
                    pendingSpace = true;
                    continue;
                }
                if (pendingSpace && sb.Length > 0) sb.Append(' ');
                pendingSpace = false;
                sb.Append(ch);
                if (sb.Length >= maxLength) break;
            }
            if (sb.Length > maxLength) sb.Length = maxLength;
            if (sb.Length > 0 && char.IsHighSurrogate(sb[sb.Length - 1])) sb.Length--;
            return sb.ToString().TrimEnd();
        }
    }
}
