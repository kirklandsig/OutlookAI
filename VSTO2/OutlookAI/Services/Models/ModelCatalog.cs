using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// Immutable set of models the Codex backend offers this ChatGPT account and
    /// the reasoning efforts each accepts. <see cref="BuiltInModelCatalog"/> until
    /// an admin runs Settings → Update Models (<see cref="ModelCatalogUpdater"/>),
    /// then the cached models.json.
    /// </summary>
    public sealed class ModelCatalog
    {
        public const string SourceBuiltIn = "built-in";
        public const string SourceChatGpt = "chatgpt";

        /// <summary>
        /// reasoning.effort values /responses accepted on 2026-09-22, read from its
        /// own invalid_value error. Filters catalog efforts when no probed set is
        /// available (the built-in catalog, or a refresh whose probe failed).
        /// </summary>
        public static readonly IReadOnlyList<string> KnownServerEfforts =
            new[] { "none", "minimal", "low", "medium", "high", "xhigh", "max" };

        private const int MaxUpgradeHops = 5;

        private readonly Dictionary<string, ModelCatalogEntry> _bySlug;

        // Per slug: the efforts offered in the UI, as (display, wire) pairs.
        private readonly Dictionary<string, List<KeyValuePair<string, string>>> _offered;

        public ModelCatalog(
            IEnumerable<ModelCatalogEntry> models,
            IEnumerable<string> serverEfforts,
            string source,
            DateTimeOffset? fetchedAt,
            string clientVersion,
            string clientVersionSource = null,
            string accountId = null)
        {
            if (models == null) throw new ArgumentNullException(nameof(models));

            _bySlug = new Dictionary<string, ModelCatalogEntry>(StringComparer.OrdinalIgnoreCase);
            var unique = new List<ModelCatalogEntry>();
            foreach (var model in models)
            {
                if (model == null || _bySlug.ContainsKey(model.Slug)) continue;
                _bySlug[model.Slug] = model;
                unique.Add(model);
            }
            if (unique.Count == 0)
            {
                throw new ArgumentException("A model catalog needs at least one model.", nameof(models));
            }

            Models = unique
                .OrderBy(m => m.Priority)
                .ThenBy(m => m.Slug, StringComparer.Ordinal)
                .ToList()
                .AsReadOnly();
            ListedSlugs = Models.Where(m => m.Listed).Select(m => m.Slug).ToList().AsReadOnly();

            var probed = serverEfforts == null
                ? null
                : serverEfforts
                    .Select(ModelCatalogJson.NormalizeEffortValue)
                    .Where(e => e != null)
                    .Distinct()
                    .ToList();
            ServerEfforts = probed != null && probed.Count > 0 ? probed.AsReadOnly() : null;
            var accepted = new HashSet<string>(ServerEfforts ?? KnownServerEfforts, StringComparer.Ordinal);

            _offered = new Dictionary<string, List<KeyValuePair<string, string>>>(StringComparer.OrdinalIgnoreCase);
            var all = new List<string> { ReasoningEffortNames.Auto };
            foreach (var model in Models)
            {
                var offered = new List<KeyValuePair<string, string>>();
                foreach (var raw in model.Efforts)
                {
                    var wire = ModelCatalogJson.NormalizeEffortValue(raw);
                    // Left out: "none" (no reasoning at all) is a different setting
                    // from the app's Auto (omit the field), no current model lists
                    // it, and "None" was Auto's old name in config files. An "auto"
                    // level would show as a second Auto that could never be sent.
                    if (wire == null || wire == "none" || wire == "auto") continue;
                    // Catalog-only modes the server rejects (Codex's client-side "ultra").
                    if (!accepted.Contains(wire)) continue;
                    var display = ReasoningEffortNames.ToDisplay(wire);
                    if (offered.Any(p => p.Key == display)) continue;
                    offered.Add(new KeyValuePair<string, string>(display, wire));
                    if (!all.Contains(display)) all.Add(display);
                }
                _offered[model.Slug] = offered;
            }
            AllOfferedEfforts = all.AsReadOnly();

            Source = string.IsNullOrWhiteSpace(source) ? SourceChatGpt : source;
            FetchedAt = fetchedAt;
            ClientVersion = clientVersion;
            ClientVersionSource = clientVersionSource;
            AccountId = accountId;
        }

        /// <summary>
        /// The candidate with the most recent data (<see cref="FetchedAt"/>;
        /// undated counts as oldest), the earliest one on a tie; null if all are null.
        /// </summary>
        public static ModelCatalog Newest(params ModelCatalog[] candidates)
        {
            ModelCatalog newest = null;
            foreach (var candidate in candidates ?? new ModelCatalog[0])
            {
                if (candidate == null) continue;
                if (newest == null
                    || (candidate.FetchedAt ?? DateTimeOffset.MinValue) > (newest.FetchedAt ?? DateTimeOffset.MinValue))
                {
                    newest = candidate;
                }
            }
            return newest;
        }

        /// <summary>Every model, hidden ones included, in catalog order (priority, then slug).</summary>
        public IReadOnlyList<ModelCatalogEntry> Models { get; }

        /// <summary>Slugs to show in the model picker (catalog visibility "list").</summary>
        public IReadOnlyList<string> ListedSlugs { get; }

        /// <summary>The server's global effort set as probed at refresh time; null when not probed.</summary>
        public IReadOnlyList<string> ServerEfforts { get; }

        /// <summary><c>Auto</c> plus every effort any model offers, display-cased, first-seen order.</summary>
        public IReadOnlyList<string> AllOfferedEfforts { get; }

        public string Source { get; }

        /// <summary>When the data was fetched (for the built-in list: when it was captured).</summary>
        public DateTimeOffset? FetchedAt { get; }

        /// <summary>The Codex client_version the catalog was fetched with.</summary>
        public string ClientVersion { get; }

        /// <summary>Where <see cref="ClientVersion"/> came from: config, github, cache or built-in.</summary>
        public string ClientVersionSource { get; }

        /// <summary>ChatGPT account the catalog was fetched for; null when unknown.</summary>
        public string AccountId { get; }

        public bool IsBuiltIn => Source == SourceBuiltIn;

        public ModelCatalogEntry Find(string slug)
        {
            if (string.IsNullOrWhiteSpace(slug)) return null;
            ModelCatalogEntry entry;
            return _bySlug.TryGetValue(slug.Trim(), out entry) ? entry : null;
        }

        /// <summary>
        /// Dropdown options for a model: <c>Auto</c> first, then the efforts the
        /// model accepts. An unknown model gets only <c>Auto</c>, which is always
        /// safe on the wire.
        /// </summary>
        public string[] EffortsFor(string slug)
        {
            var entry = Find(slug);
            var result = new List<string> { ReasoningEffortNames.Auto };
            if (entry != null) result.AddRange(_offered[entry.Slug].Select(p => p.Key));
            return result.ToArray();
        }

        /// <summary>
        /// Canonical display spelling of <paramref name="effort"/> when it is
        /// <c>Auto</c> (or its legacy name <c>None</c>) or offered by some model;
        /// otherwise null.
        /// </summary>
        public string NormalizeEffort(string effort)
        {
            if (string.IsNullOrWhiteSpace(effort)) return null;
            if (ReasoningEffortNames.IsAuto(effort)) return ReasoningEffortNames.Auto;
            var trimmed = effort.Trim();
            foreach (var display in AllOfferedEfforts)
            {
                if (string.Equals(display, trimmed, StringComparison.OrdinalIgnoreCase)) return display;
            }
            return null;
        }

        /// <summary>First listed model that has not retired by <paramref name="now"/>.</summary>
        public string DefaultModelAt(DateTimeOffset now)
        {
            foreach (var model in Models)
            {
                if (model.Listed && (model.Upgrade == null || !model.Upgrade.IsRetiredAt(now))) return model.Slug;
            }
            foreach (var model in Models)
            {
                if (model.Listed) return model.Slug;
            }
            return Models[0].Slug;
        }

        /// <summary>
        /// The model to send: <paramref name="configured"/> while the catalog
        /// offers it and it has not retired; after its retirement date, its
        /// upgrade target (following chains); otherwise the catalog default.
        /// </summary>
        public string ResolveEffectiveModel(string configured, DateTimeOffset now)
        {
            var entry = Find(configured);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int hop = 0; entry != null && hop <= MaxUpgradeHops; hop++)
            {
                if (entry.Upgrade == null || !entry.Upgrade.IsRetiredAt(now)) return entry.Slug;
                if (!visited.Add(entry.Slug)) break;
                entry = Find(entry.Upgrade.Model);
            }
            return DefaultModelAt(now);
        }

        /// <summary>
        /// Wire value for reasoning.effort, or null to omit the field. Efforts the
        /// catalog says this model doesn't take are omitted (server default)
        /// rather than sent, so a stale selection never fails the request. For a
        /// model outside the catalog the effort passes through lowercased.
        /// </summary>
        public string ResolveWireEffort(string slug, string effort)
        {
            if (ReasoningEffortNames.IsAuto(effort)) return null;
            var trimmed = effort.Trim();
            var entry = Find(slug);
            if (entry == null) return trimmed.ToLowerInvariant();
            foreach (var pair in _offered[entry.Slug])
            {
                if (string.Equals(pair.Key, trimmed, StringComparison.OrdinalIgnoreCase)) return pair.Value;
            }
            return null;
        }
    }
}
