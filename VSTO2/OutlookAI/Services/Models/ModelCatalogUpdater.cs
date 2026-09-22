using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;
using OutlookAI.Services.Updates;

namespace OutlookAI.Services.Models
{
    public sealed class ModelCatalogUpdateResult
    {
        private static readonly IReadOnlyList<string> None = new string[0];

        private ModelCatalogUpdateResult()
        {
        }

        public bool Succeeded { get; private set; }

        /// <summary>User-facing failure reason; null on success.</summary>
        public string Error { get; private set; }

        public ModelCatalog Catalog { get; private set; }
        public ModelCatalogSaveResult Save { get; private set; }

        /// <summary>Listed slugs that are new / gone compared with the previous catalog.</summary>
        public IReadOnlyList<string> Added { get; private set; } = None;
        public IReadOnlyList<string> Removed { get; private set; } = None;

        /// <summary>Where client_version came from: config, github, cache or built-in.</summary>
        public string ClientVersionSource { get; private set; }

        /// <summary>
        /// False when this refresh couldn't read the server's effort set and fell
        /// back to the last known one.
        /// </summary>
        public bool ServerEffortsProbed { get; private set; }

        /// <summary>
        /// Configured models the fresh list left out that still route to
        /// themselves (kept, since absence alone isn't evidence of retirement).
        /// </summary>
        public IReadOnlyList<string> KeptUnlisted { get; private set; } = None;

        /// <summary>
        /// Another session saved a newer list while this refresh ran; that list
        /// was kept and is <see cref="Catalog"/>.
        /// </summary>
        public bool SupersededByNewer { get; private set; }

        internal static ModelCatalogUpdateResult Failed(string error)
        {
            return new ModelCatalogUpdateResult { Error = error };
        }

        internal static ModelCatalogUpdateResult Success(
            ModelCatalog catalog, ModelCatalogSaveResult save, ModelCatalog previous, string clientVersionSource,
            bool serverEffortsProbed, IReadOnlyList<string> keptUnlisted, bool supersededByNewer)
        {
            var before = previous != null ? previous.ListedSlugs : None;
            var after = catalog.ListedSlugs;
            return new ModelCatalogUpdateResult
            {
                Succeeded = true,
                Catalog = catalog,
                Save = save,
                Added = after.Where(s => !before.Contains(s, StringComparer.OrdinalIgnoreCase)).ToList().AsReadOnly(),
                Removed = before.Where(s => !after.Contains(s, StringComparer.OrdinalIgnoreCase)).ToList().AsReadOnly(),
                ClientVersionSource = clientVersionSource,
                ServerEffortsProbed = serverEffortsProbed,
                KeptUnlisted = keptUnlisted ?? None,
                SupersededByNewer = supersededByNewer,
            };
        }
    }

    /// <summary>
    /// Settings → Update Models: fetch this account's model catalog from the
    /// Codex backend, learn the server's reasoning-effort set, and cache both so
    /// new models and efforts appear without an OutlookAI release.
    /// </summary>
    public sealed class ModelCatalogUpdater
    {
        public const string CodexRepo = "openai/codex";

        // A saved choice can keep following a vanished model's replacement for
        // this long; after that it falls back to the catalog default.
        private static readonly TimeSpan CarryForwardWindow = TimeSpan.FromDays(365);

        private readonly HttpClient _http;
        private readonly ModelCatalogStore _store;
        private readonly Func<DateTimeOffset> _clock;

        public ModelCatalogUpdater(HttpClient http, ModelCatalogStore store, Func<DateTimeOffset> clock = null)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _clock = clock ?? (() => DateTimeOffset.UtcNow);
        }

        // The optional steps get their own budgets so a stalled GitHub or probe
        // call can't use up the caller's whole timeout.
        internal TimeSpan GitHubLookupTimeout { get; set; } = TimeSpan.FromSeconds(10);
        internal TimeSpan FetchTimeout { get; set; } = TimeSpan.FromSeconds(60);
        internal TimeSpan ProbeTimeout { get; set; } = TimeSpan.FromSeconds(15);

        /// <param name="current">The catalog in use now; the diff baseline and one source of prior knowledge.</param>
        /// <param name="clientVersionOverride">Config.ModelCatalogClientVersion; blank or invalid → auto-detect.</param>
        /// <param name="keepModels">Models named in config; kept if the fresh list omits them.</param>
        public async Task<ModelCatalogUpdateResult> UpdateAsync(
            CodexAuthService auth, ModelCatalog current, string clientVersionOverride,
            IEnumerable<string> keepModels, CancellationToken ct)
        {
            if (auth == null) return ModelCatalogUpdateResult.Failed("ChatGPT sign-in is unavailable. Restart Outlook.");

            string token;
            try
            {
                token = await auth.GetAccessTokenAsync(ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                return ModelCatalogUpdateResult.Failed(
                    "Couldn't get a ChatGPT sign-in token (" + ex.Message + "). Check ChatGPT Account above.");
            }
            var accountId = auth.GetStatus().AccountId;

            // Everything this machine already knows, newest first: the running
            // session's list, a refresh another admin saved, this build's own data.
            var known = new[] { current, _store.Load(), BuiltInModelCatalog.Instance }
                .Where(c => c != null)
                .OrderByDescending(c => c.FetchedAt ?? DateTimeOffset.MinValue)
                .ToList();

            var version = await ResolveClientVersionAsync(known, clientVersionOverride, ct).ConfigureAwait(false);
            Trace("fetching catalog, client_version=" + version.Value + " (" + version.Source + ")");

            var client = new CodexModelsClient(_http);
            ModelsFetchResult fetch;
            using (var fetchCts = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                fetchCts.CancelAfter(FetchTimeout);
                try
                {
                    fetch = await client.FetchModelsAsync(token, accountId, version.Value, fetchCts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                {
                    fetch = null;
                }
            }
            if (fetch == null || !fetch.Succeeded)
            {
                var error = fetch == null ? "ChatGPT didn't send the model list in time." : fetch.Error;
                Trace("fetch failed: " + error);
                return ModelCatalogUpdateResult.Failed(error);
            }

            var now = _clock();
            var fetched = fetch.Models;
            var keep = new HashSet<string>(
                (keepModels ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()),
                StringComparer.OrdinalIgnoreCase);

            // This fetch as a catalog: the probe target and, below, a metadata
            // source (it may be the only one that knows a newly configured model).
            var fetchedCatalog = new ModelCatalog(fetched, null, ModelCatalog.SourceChatGpt, now, version.Value);

            // Probe with a model the backend accepts, so the reply is the effort error.
            var probed = await ProbeAsync(client, token, accountId, fetchedCatalog.DefaultModelAt(now), ct).ConfigureAwait(false);
            if (probed != null && !fetched.Any(m => m.Efforts.Any(e => probed.Contains(e))))
            {
                Trace("ignoring probed effort set that matches no catalog effort: " + string.Join(",", probed));
                probed = null;
            }
            // Without a fresh probe, keep the last set a refresh did read rather
            // than regress to the set this build shipped with.
            var serverEfforts = probed ?? known
                .Where(c => c.ServerEfforts != null)
                .Select(c => c.ServerEfforts)
                .FirstOrDefault();

            // Built under the store's lock against whatever is on disk at that
            // moment, so a refresh another session committed while this one was
            // on the network is merged, never overwritten.
            Func<ModelCatalog, ModelCatalog> build = onDisk =>
            {
                var sources = new[] { onDisk, fetchedCatalog }.Concat(known)
                    .Where(c => c != null)
                    .Distinct()
                    .OrderByDescending(c => c.FetchedAt ?? DateTimeOffset.MinValue)
                    .ToList();
                if (onDisk != null && onDisk.FetchedAt > now)
                {
                    // That session fetched after this one: its list is fresher. Keep
                    // it, plus what this session must not lose.
                    var merged = Retain(onDisk.Models, sources, keep, now);
                    return merged.Count == onDisk.Models.Count
                        ? onDisk
                        : new ModelCatalog(merged, onDisk.ServerEfforts, ModelCatalog.SourceChatGpt, onDisk.FetchedAt,
                            onDisk.ClientVersion, onDisk.ClientVersionSource, onDisk.AccountId);
                }
                return new ModelCatalog(Retain(fetched, sources, keep, now), serverEfforts, ModelCatalog.SourceChatGpt,
                    now, version.Value, version.Source, accountId);
            };
            var save = _store.Commit(build);
            var catalog = save.Committed;
            var superseded = catalog.FetchedAt > now;

            var keptUnlisted = catalog.Models
                .Where(entry => keep.Contains(entry.Slug)
                    && !entry.Listed
                    && !fetched.Any(m => string.Equals(m.Slug, entry.Slug, StringComparison.OrdinalIgnoreCase))
                    && string.Equals(catalog.ResolveEffectiveModel(entry.Slug, now), entry.Slug, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.Slug)
                .ToList();

            Trace("fetched " + fetched.Count + " models; " + catalog.Models.Count + " kept after merge; listed: "
                + string.Join(",", catalog.ListedSlugs)
                + "; server efforts=" + (probed != null ? string.Join(",", probed) : "<not probed>")
                + "; shared save=" + (save.SharedSaved ? "ok" : save.SharedError)
                + "; user save=" + (save.UserSaved ? "ok" : save.UserError)
                + (superseded ? "; built on a newer list saved meanwhile" : ""));
            return ModelCatalogUpdateResult.Success(
                catalog, save, current, version.Source, probed != null, keptUnlisted.AsReadOnly(), superseded);
        }

        /// <summary>
        /// <paramref name="baseModels"/> plus the models it leaves out that there is
        /// evidence to keep: an announced retirement (for up to a year past its
        /// date; date and replacement kept, so traffic moves only when it passes),
        /// or a model config names. Absence alone is not a retirement: the list is
        /// filtered by client_version and plan, while inference is not. Replacements
        /// an announced retirement routes to are kept too, or the route would dead-end.
        /// <paramref name="newestFirst"/> supplies metadata; the newest entry per slug
        /// decides, and a withdrawn retirement is kept as that newest entry so an
        /// older list can't revive it, now or at a later refresh.
        /// </summary>
        internal static List<ModelCatalogEntry> Retain(
            IReadOnlyList<ModelCatalogEntry> baseModels, IEnumerable<ModelCatalog> newestFirst,
            ICollection<string> keep, DateTimeOffset now)
        {
            var newest = new Dictionary<string, ModelCatalogEntry>(StringComparer.OrdinalIgnoreCase);
            var discovered = new List<ModelCatalogEntry>();
            var announcedSomewhere = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in newestFirst.SelectMany(c => c.Models))
            {
                if (IsAnnouncedRetirement(entry, now)) announcedSomewhere.Add(entry.Slug);
                if (newest.ContainsKey(entry.Slug)) continue;
                newest[entry.Slug] = entry;
                discovered.Add(entry);
            }

            var result = new List<ModelCatalogEntry>(baseModels);
            var present = new HashSet<string>(baseModels.Select(m => m.Slug), StringComparer.OrdinalIgnoreCase);
            Action<ModelCatalogEntry> carry = entry =>
            {
                result.Add(new ModelCatalogEntry(entry.Slug, entry.DisplayName, entry.Description, false,
                    entry.Priority, entry.Efforts, entry.Upgrade));
                present.Add(entry.Slug);
            };

            foreach (var entry in discovered)
            {
                if (result.Count >= ModelCatalogJson.MaxModels) break;
                if (present.Contains(entry.Slug)) continue;
                // The last clause keeps a withdrawal: the newest entry has no
                // retirement but an older list (e.g. the built-in one) still
                // announces one, which would win again at the next refresh.
                if (keep.Contains(entry.Slug) || IsAnnouncedRetirement(entry, now) || announcedSomewhere.Contains(entry.Slug))
                {
                    carry(entry);
                }
            }

            // Follow each announced retirement to its replacement, transitively
            // (entries appended here are visited by this same loop).
            for (int i = 0; i < result.Count && result.Count < ModelCatalogJson.MaxModels; i++)
            {
                if (!IsAnnouncedRetirement(result[i], now) || present.Contains(result[i].Upgrade.Model)) continue;
                ModelCatalogEntry target;
                if (newest.TryGetValue(result[i].Upgrade.Model, out target)) carry(target);
            }
            return result;
        }

        private static bool IsAnnouncedRetirement(ModelCatalogEntry entry, DateTimeOffset now)
        {
            return entry.Upgrade != null
                && entry.Upgrade.RetirementAt.HasValue
                && ModelCatalogJson.IsValidSlug(entry.Upgrade.Model)
                && now - entry.Upgrade.RetirementAt.Value <= CarryForwardWindow;
        }

        // Explicit config override, else the highest of: the built-in floor, the
        // versions earlier refreshes detected, and the latest openai/codex release,
        // so the list keeps tracking Codex without an app update. A version an admin
        // pinned is never reused as a floor; it would outlive the pin.
        private async Task<(string Value, string Source)> ResolveClientVersionAsync(
            IReadOnlyList<ModelCatalog> known, string overrideValue, CancellationToken ct)
        {
            var configured = CodexClientVersion.Normalize(overrideValue);
            if (configured != null) return (configured, "config");
            if (!string.IsNullOrWhiteSpace(overrideValue))
            {
                Trace("ignoring invalid ModelCatalogClientVersion '" + overrideValue + "'");
            }

            var latest = await LookUpLatestCodexVersionAsync(ct).ConfigureAwait(false);
            var cached = CodexClientVersion.Max(known
                .Where(c => !c.IsBuiltIn && c.ClientVersionSource != "config")
                .Select(c => c.ClientVersion)
                .ToArray());
            var best = CodexClientVersion.Max(CodexClientVersion.BuiltInFloor, cached, latest);
            var source = best == latest ? "github" : best == cached ? "cache" : "built-in";
            return (best, source);
        }

        private async Task<string> LookUpLatestCodexVersionAsync(CancellationToken ct)
        {
            using (var lookupCts = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                var lookup = new GitHubReleaseClient(_http, CodexRepo, "OutlookAI-ModelCatalog")
                    .GetLatestStableAsync(lookupCts.Token);
                // WhenAny, not only the token: the lookup must not be able to hold
                // the refresh hostage even if its HTTP stack ignores cancellation.
                var finished = await Task.WhenAny(lookup, Task.Delay(GitHubLookupTimeout, ct)).ConfigureAwait(false);
                if (finished != lookup)
                {
                    lookupCts.Cancel();
                    ObserveFailure(lookup);
                    ct.ThrowIfCancellationRequested();
                    Trace("latest Codex release lookup timed out");
                    return null;
                }
                try
                {
                    var found = await lookup.ConfigureAwait(false) as ReleaseFound;
                    if (found != null) return CodexClientVersion.Normalize(found.Info.Tag);
                    Trace("latest Codex release lookup: " + lookup.Result.GetType().Name);
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Trace("latest Codex release lookup failed: " + ex.Message);
                }
                return null;
            }
        }

        private async Task<IReadOnlyList<string>> ProbeAsync(
            CodexModelsClient client, string token, string accountId, string model, CancellationToken ct)
        {
            using (var probeCts = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                probeCts.CancelAfter(ProbeTimeout);
                try
                {
                    return await client.ProbeServerEffortsAsync(token, accountId, model, probeCts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                {
                    Trace("effort probe timed out");
                    return null;
                }
            }
        }

        private static void ObserveFailure(Task task)
        {
            task.ContinueWith(
                t => { var ignored = t.Exception; },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
        }

        private static void Trace(string message)
        {
            TraceLog.Write("ModelCatalogUpdater: " + message, "Models");
        }
    }
}
