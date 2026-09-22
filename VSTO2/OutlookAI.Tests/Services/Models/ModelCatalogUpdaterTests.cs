using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Models;
using OutlookAI.Tests.Helpers;
using Xunit;
using static OutlookAI.Tests.Helpers.TestCatalogs;

namespace OutlookAI.Tests.Services.Models
{
    public class ModelCatalogUpdaterTests : IDisposable
    {
        private const string GitHubLatest = "https://api.github.com/repos/openai/codex/releases/latest";
        private readonly string _dir;
        private readonly string _authPath;
        private readonly ModelCatalogStore _store;
        private readonly HttpClient _authHttp = new HttpClient(new FakeHttpMessageHandler());

        public ModelCatalogUpdaterTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "outlookai-updater-models", Path.GetRandomFileName());
            Directory.CreateDirectory(_dir);
            _authPath = Path.Combine(_dir, "auth.json");
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-test\",\"id_token\":\"\",\"refresh_token\":\"r1\","
                + "\"account_id\":\"acct-123\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");
            _store = new ModelCatalogStore(Path.Combine(_dir, "shared", "models.json"), Path.Combine(_dir, "user", "models.json"));
        }

        public void Dispose()
        {
            _authHttp.Dispose();
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private static void QueueGitHub(FakeHttpMessageHandler fake, string tag)
        {
            fake.QueueJson(HttpStatusCode.OK,
                "{\"tag_name\":\"" + tag + "\",\"html_url\":\"https://github.com/openai/codex/releases/tag/" + tag
                + "\",\"body\":\"\",\"published_at\":\"2026-09-20T00:00:00Z\",\"assets\":[]}");
        }

        private static ModelCatalog Current(string clientVersion)
        {
            return new ModelCatalog(
                new[] { Entry("gpt-5.5", 1, new[] { "low" }), Entry("gpt-4.1-mini", 2, new[] { "low" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddDays(-30), clientVersion);
        }

        private async Task<ModelCatalogUpdateResult> RunAsync(
            FakeHttpMessageHandler fake, ModelCatalog current, string overrideVersion = "", string authPath = null,
            TimeSpan? gitHubTimeout = null, string[] keep = null)
        {
            using (var http = new HttpClient(fake))
            using (var auth = new CodexAuthService(authPath ?? _authPath, _authHttp))
            {
                var updater = new ModelCatalogUpdater(http, _store, () => FixedNow);
                if (gitHubTimeout.HasValue) updater.GitHubLookupTimeout = gitHubTimeout.Value;
                return await updater.UpdateAsync(auth, current, overrideVersion, keep ?? new string[0], CancellationToken.None);
            }
        }

        private const string AstraAndSolOnly =
            "{\"models\":["
            + "{\"slug\":\"gpt-6-astra\",\"visibility\":\"list\",\"priority\":1,\"supported_reasoning_levels\":[{\"effort\":\"low\"}]},"
            + "{\"slug\":\"gpt-5.6-sol\",\"visibility\":\"list\",\"priority\":4,\"supported_reasoning_levels\":[{\"effort\":\"low\"}]}"
            + "]}";

        private static void QueueCatalogAndProbe(FakeHttpMessageHandler fake, string modelsJson = LiveModelsResponse)
        {
            fake.QueueJson(HttpStatusCode.OK, modelsJson);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);
        }

        [Fact]
        public async Task UpdateAsync_UsesLatestCodexRelease_ProbesEfforts_SavesAndDiffs()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, Current("0.155.1"));

            Assert.True(result.Succeeded, result.Error);
            Assert.Equal(3, fake.Requests.Count);
            Assert.Equal(GitHubLatest, fake.Requests[0].RequestUri.ToString());
            Assert.Null(fake.Requests[0].Headers.Authorization);   // no ChatGPT token to GitHub
            Assert.Equal("https://chatgpt.com/backend-api/codex/models?client_version=0.156.0", fake.Requests[1].RequestUri.ToString());
            Assert.Equal("acct-123", fake.Requests[1].Headers.GetValues("ChatGPT-Account-ID").Single());
            Assert.Contains("\"model\":\"gpt-6-astra\"", fake.RequestBodies[2]);   // probe uses the top listed model

            var catalog = result.Catalog;
            Assert.Equal(ModelCatalog.SourceChatGpt, catalog.Source);
            Assert.Equal(FixedNow, catalog.FetchedAt);
            Assert.Equal("0.156.0", catalog.ClientVersion);
            Assert.Equal("github", result.ClientVersionSource);
            Assert.Equal("github", catalog.ClientVersionSource);
            Assert.Equal("acct-123", catalog.AccountId);
            Assert.True(result.ServerEffortsProbed);
            Assert.DoesNotContain("Ultra", catalog.EffortsFor("gpt-6-astra"));
            Assert.Equal(new[] { "gpt-6-astra", "gpt-5.6-sol" }, result.Added);
            Assert.Equal(new[] { "gpt-4.1-mini" }, result.Removed);

            Assert.True(result.Save.SharedSaved);
            Assert.True(result.Save.UserSaved);
            Assert.Equal(catalog.ListedSlugs, _store.Load().ListedSlugs);
        }

        [Fact]
        public async Task UpdateAsync_ConfigOverride_SkipsGitHub()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, Current("0.155.1"), overrideVersion: " 0.170.0 ");

            Assert.True(result.Succeeded, result.Error);
            Assert.Equal("https://chatgpt.com/backend-api/codex/models?client_version=0.170.0", fake.Requests[0].RequestUri.ToString());
            Assert.Equal("config", result.ClientVersionSource);
        }

        [Fact]
        public async Task UpdateAsync_InvalidOverride_IsIgnored()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, Current("0.155.1"), overrideVersion: "latest");

            Assert.Equal(GitHubLatest, fake.Requests[0].RequestUri.ToString());
            Assert.Equal("0.156.0", result.Catalog.ClientVersion);
        }

        [Theory]
        [InlineData("0.158.0", "0.158.0", "cache")]
        [InlineData("0.100.0", CodexClientVersion.BuiltInFloor, "built-in")]
        [InlineData(null, CodexClientVersion.BuiltInFloor, "built-in")]
        public async Task UpdateAsync_GitHubUnavailable_UsesBestKnownVersion(string cached, string expected, string expectedSource)
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueText(HttpStatusCode.InternalServerError, "down");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, Current(cached));

            Assert.True(result.Succeeded, result.Error);
            Assert.Equal("https://chatgpt.com/backend-api/codex/models?client_version=" + expected, fake.Requests[1].RequestUri.ToString());
            Assert.Equal(expectedSource, result.ClientVersionSource);
        }

        [Fact]
        public async Task UpdateAsync_NeverGoesBelowTheCachedVersion()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.150.0");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, Current("0.158.0"));

            Assert.Equal("0.158.0", result.Catalog.ClientVersion);
        }

        [Fact]
        public async Task UpdateAsync_FetchFailure_KeepsTheExistingCacheUntouched()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueText(HttpStatusCode.BadGateway, "bad gateway");

            var result = await RunAsync(fake, Current("0.155.1"));

            Assert.False(result.Succeeded);
            Assert.Contains("502", result.Error);
            Assert.Null(result.Catalog);
            Assert.Equal(2, fake.Requests.Count);    // no probe after a failed fetch
            Assert.Null(_store.Load());
        }

        [Fact]
        public async Task UpdateAsync_NotSignedIn_FailsBeforeAnyRequest()
        {
            var fake = new FakeHttpMessageHandler();

            var result = await RunAsync(fake, Current("0.155.1"), authPath: Path.Combine(_dir, "missing-auth.json"));

            Assert.False(result.Succeeded);
            // Says what failed (any token failure, not only "signed out") and why.
            Assert.Contains("ChatGPT sign-in", result.Error);
            Assert.Contains("not signed in", result.Error);
            Assert.Empty(fake.Requests);
        }

        [Fact]
        public async Task UpdateAsync_KeepsAnnouncedRetirements_ForModelsThatLeftTheList()
        {
            // gpt-5.5 has dropped out of the live list after its retirement; a
            // config still naming it must land on gpt-5.6-sol, not the top model.
            var previous = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-5.6-sol", 4, new[] { "low" }),
                    Entry("gpt-5.5", 12, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(-2))),
                    Entry("gpt-early", 13, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(20))),
                    Entry("gpt-ancient", 15, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(-400))),
                },
                // Newer than the built-in list, so its metadata wins.
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraAndSolOnly);

            var result = await RunAsync(fake, previous);

            var catalog = result.Catalog;
            Assert.True(result.Succeeded, result.Error);
            Assert.False(catalog.Find("gpt-5.5").Listed);
            Assert.Equal("gpt-5.6-sol", catalog.ResolveEffectiveModel("gpt-5.5", FixedNow));
            // Missing ahead of its announced date is not evidence of retirement:
            // the date stands.
            Assert.Equal("gpt-early", catalog.ResolveEffectiveModel("gpt-early", FixedNow));
            Assert.Equal("gpt-5.6-sol", catalog.ResolveEffectiveModel("gpt-early", FixedNow.AddDays(21)));
            // Retired too long ago: dropped.
            Assert.Null(catalog.Find("gpt-ancient"));
            Assert.Contains("gpt-5.5", result.Removed);
            // And the carried entry survives the cache round trip.
            Assert.Equal("gpt-5.6-sol", _store.Load().ResolveEffectiveModel("gpt-5.5", FixedNow));
        }

        [Fact]
        public async Task UpdateAsync_KeepsConfiguredModels_ThatTheListLeavesOut_WithoutRetiringThem()
        {
            // client_version filtering (or another plan) can omit a model that
            // still answers requests. Absence alone must not reroute traffic.
            var previous = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-6-sol", 2, new[] { "low", "max" }),
                    Entry("gpt-soft", 3, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-astra", null)),
                    Entry("gpt-unused", 4, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-astra", null)),
                },
                // Newer than the built-in list, so its gpt-6-sol entry is the one kept.
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraAndSolOnly);

            var result = await RunAsync(fake, previous, keep: new[] { "gpt-6-sol", "GPT-SOFT" });

            var catalog = result.Catalog;
            Assert.False(catalog.Find("gpt-6-sol").Listed);
            Assert.Equal("gpt-6-sol", catalog.ResolveEffectiveModel("gpt-6-sol", FixedNow));
            Assert.Equal(new[] { "None", "Low", "Max" }, catalog.EffortsFor("gpt-6-sol"));
            Assert.Equal("gpt-soft", catalog.ResolveEffectiveModel("gpt-soft", FixedNow));
            Assert.Null(catalog.Find("gpt-unused"));   // not configured, no retirement date
            Assert.Equal(new[] { "gpt-6-sol", "gpt-soft" }, result.KeptUnlisted);
        }

        [Fact]
        public async Task UpdateAsync_UsesTheNewestKnownMetadata_ForACarriedModel()
        {
            // The running session still has an older replacement for gpt-x;
            // another admin's newer cache corrected it.
            var olderSession = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-x", 9, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-astra", FixedNow.AddDays(-1))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddDays(-5), "0.155.1");
            _store.Save(new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-5.6-sol", 4, new[] { "low" }),
                    Entry("gpt-x", 9, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(-1))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-2), "0.155.1"));
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraAndSolOnly);

            var result = await RunAsync(fake, olderSession);

            Assert.Equal("gpt-5.6-sol", result.Catalog.ResolveEffectiveModel("gpt-x", FixedNow));
        }

        private const string AstraOnly =
            "{\"models\":[{\"slug\":\"gpt-6-astra\",\"visibility\":\"list\",\"priority\":1,\"supported_reasoning_levels\":[{\"effort\":\"low\"}]}]}";

        private static HttpResponseMessage Json(HttpStatusCode status, string json)
        {
            return new HttpResponseMessage(status) { Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json") };
        }

        [Fact]
        public async Task UpdateAsync_KeepsTheReplacementAnAnnouncedRetirementRoutesTo()
        {
            // The list omits both the retiring model and its replacement; keeping
            // only the former would send it to the top model on its date.
            var previous = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-5.6-sol", 4, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-sol", null)),
                    Entry("gpt-5.5", 12, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(20))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraOnly);

            var result = await RunAsync(fake, previous, keep: new[] { "gpt-5.5" });

            var catalog = result.Catalog;
            Assert.NotNull(catalog.Find("gpt-5.6-sol"));
            Assert.Equal("gpt-5.5", catalog.ResolveEffectiveModel("gpt-5.5", FixedNow));
            Assert.Equal("gpt-5.6-sol", catalog.ResolveEffectiveModel("gpt-5.5", FixedNow.AddDays(21)));
        }

        [Fact]
        public async Task UpdateAsync_ANewerEntryWithoutARetirement_BeatsAnOlderRetirement()
        {
            // The newest knowledge of gpt-x has no retirement (withdrawn); an older
            // list's retirement must not come back just because gpt-x isn't kept.
            _store.Save(new ModelCatalog(
                new[] { Entry("gpt-6-astra", 1, new[] { "low" }), Entry("gpt-x", 9, new[] { "low" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-2), "0.155.1"));
            var olderSession = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-x", 9, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-astra", FixedNow.AddDays(-1))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddDays(-5), "0.155.1");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraAndSolOnly);

            var result = await RunAsync(fake, olderSession);

            // Kept as the withdrawal (hidden, no retirement), never rerouted.
            Assert.Null(result.Catalog.Find("gpt-x").Upgrade);
            Assert.Equal("gpt-x", result.Catalog.ResolveEffectiveModel("gpt-x", FixedNow));
        }

        [Fact]
        public async Task UpdateAsync_CarriesTheLatestMetadata_EvenIfItWasSavedDuringTheFetch()
        {
            // Another session saves a corrected retirement for gpt-x while this
            // refresh waits on the network; the commit must build on that.
            var olderSession = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-x", 9, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-6-astra", FixedNow.AddDays(10))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var corrected = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-5.6-sol", 4, new[] { "low" }),
                    Entry("gpt-x", 9, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-5.6-sol", FixedNow.AddDays(10))),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddMinutes(-1), "0.156.0");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.Queue(request => { _store.Save(corrected); return Json(HttpStatusCode.OK, AstraOnly); });
            fake.QueueJson(HttpStatusCode.BadRequest, InvalidEffortErrorBody);

            var result = await RunAsync(fake, olderSession);

            Assert.False(result.SupersededByNewer);   // our fetch is the newer one
            Assert.Equal("gpt-5.6-sol", result.Catalog.Find("gpt-x").Upgrade.Model);
            Assert.Equal("gpt-5.6-sol", _store.Load().Find("gpt-x").Upgrade.Model);
        }

        [Fact]
        public async Task UpdateAsync_AdoptingANewerList_KeepsThisSessionsConfiguredModel()
        {
            var session = new ModelCatalog(
                new[] { Entry("gpt-6-astra", 1, new[] { "low" }), Entry("gpt-6-sol", 2, new[] { "low", "max" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var theirs = new ModelCatalog(new[] { Entry("gpt-7", 1, new[] { "low" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddMinutes(1), "0.156.0");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK, AstraOnly);
            fake.Queue(request => { _store.Save(theirs); return Json(HttpStatusCode.BadRequest, InvalidEffortErrorBody); });

            var result = await RunAsync(fake, session, keep: new[] { "gpt-6-sol" });

            Assert.True(result.SupersededByNewer);
            Assert.Equal(new[] { "gpt-7" }, result.Catalog.ListedSlugs);
            Assert.Equal("gpt-6-sol", result.Catalog.ResolveEffectiveModel("gpt-6-sol", FixedNow));
            Assert.Equal(new[] { "gpt-6-sol" }, result.KeptUnlisted);
            Assert.NotNull(_store.Load().Find("gpt-6-sol"));
        }

        [Fact]
        public async Task UpdateAsync_AWithdrawnRetirement_StaysWithdrawnAcrossRefreshes()
        {
            // The built-in list announces gpt-5.5's retirement; newer data withdrew
            // it. Nothing configures gpt-5.5 here, and the listing omits it, yet
            // the withdrawal must outlive the next refresh (cache round trip) or
            // the built-in announcement comes back and reroutes other users.
            var withdrawn = new ModelCatalog(
                new[]
                {
                    Entry("gpt-6-astra", 1, new[] { "low" }),
                    Entry("gpt-5.6-sol", 4, new[] { "low" }),
                    Entry("gpt-5.5", 12, new[] { "low" }, upgrade: new ModelUpgradeInfo("gpt-5.6-sol", null)),
                },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.155.1");
            var afterRetirementDate = new DateTimeOffset(2026, 10, 15, 0, 0, 0, TimeSpan.Zero);

            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake, AstraAndSolOnly);
            var first = await RunAsync(fake, withdrawn);

            var fake2 = new FakeHttpMessageHandler();
            QueueGitHub(fake2, "rust-v0.156.0");
            QueueCatalogAndProbe(fake2, AstraAndSolOnly);
            var second = await RunAsync(fake2, _store.Load());

            Assert.Equal("gpt-5.5", first.Catalog.ResolveEffectiveModel("gpt-5.5", afterRetirementDate));
            Assert.Equal("gpt-5.5", second.Catalog.ResolveEffectiveModel("gpt-5.5", afterRetirementDate));
            Assert.Equal("gpt-5.5", _store.Load().ResolveEffectiveModel("gpt-5.5", afterRetirementDate));
        }

        [Fact]
        public async Task UpdateAsync_AdoptingANewerList_KeepsAConfiguredModelThisFetchJustFound()
        {
            var theirs = new ModelCatalog(new[] { Entry("gpt-7", 1, new[] { "low" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddMinutes(1), "0.156.0");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK,
                "{\"models\":["
                + "{\"slug\":\"gpt-6-astra\",\"visibility\":\"list\",\"priority\":1,\"supported_reasoning_levels\":[{\"effort\":\"low\"}]},"
                + "{\"slug\":\"gpt-new\",\"visibility\":\"list\",\"priority\":2,\"supported_reasoning_levels\":[{\"effort\":\"high\"}]}"
                + "]}");
            fake.Queue(request => { _store.Save(theirs); return Json(HttpStatusCode.BadRequest, InvalidEffortErrorBody); });

            var result = await RunAsync(fake, Current("0.155.1"), keep: new[] { "gpt-new" });

            Assert.True(result.SupersededByNewer);
            Assert.Equal("gpt-new", result.Catalog.ResolveEffectiveModel("gpt-new", FixedNow));
            Assert.Equal(new[] { "None", "High" }, result.Catalog.EffortsFor("gpt-new"));
        }

        [Fact]
        public async Task UpdateAsync_AdoptsANewerListSavedMeanwhile()
        {
            // Another session finished a refresh after ours started: keep theirs
            // rather than overwrite newer shared data with our older snapshot.
            var theirs = new ModelCatalog(new[] { Entry("gpt-7", 1, new[] { "low" }) },
                null, ModelCatalog.SourceChatGpt, FixedNow.AddMinutes(1), "0.156.0");
            _store.Save(theirs);
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake);

            var result = await RunAsync(fake, Current("0.155.1"));

            Assert.True(result.Succeeded, result.Error);
            Assert.True(result.SupersededByNewer);
            Assert.Equal(new[] { "gpt-7" }, result.Catalog.ListedSlugs);
            Assert.Equal(new[] { "gpt-7" }, _store.Load().ListedSlugs);
        }

        [Fact]
        public async Task UpdateAsync_FloorsTheVersionAtTheNewestCacheOnDisk()
        {
            // Another admin refreshed with a newer Codex version while this
            // session was running on an older in-memory list.
            _store.Save(new ModelCatalog(new[] { Entry("gpt-7", 1, new[] { "low" }) }, null,
                ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.158.0", clientVersionSource: "github"));
            var fake = new FakeHttpMessageHandler();
            fake.QueueText(HttpStatusCode.Forbidden, "rate limited");
            QueueCatalogAndProbe(fake);

            var result = await RunAsync(fake, BuiltInModelCatalog.Instance);

            Assert.True(result.Succeeded, result.Error);
            Assert.Equal("0.158.0", result.Catalog.ClientVersion);
            Assert.Equal("cache", result.ClientVersionSource);
        }

        [Fact]
        public async Task UpdateAsync_AConfigPinnedVersion_DoesNotOutliveThePin()
        {
            var pinned = new ModelCatalog(new[] { Entry("gpt-7", 1, new[] { "low" }) }, null,
                ModelCatalog.SourceChatGpt, FixedNow.AddHours(-1), "0.160.0", clientVersionSource: "config");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            QueueCatalogAndProbe(fake);

            var result = await RunAsync(fake, pinned, overrideVersion: "");

            Assert.Equal("0.156.0", result.Catalog.ClientVersion);
            Assert.Equal("github", result.ClientVersionSource);
        }

        [Fact]
        public async Task UpdateAsync_StalledGitHubLookup_FallsBackInsteadOfFailing()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueHang();
            QueueCatalogAndProbe(fake);

            var result = await RunAsync(fake, Current("0.157.0"), gitHubTimeout: TimeSpan.FromMilliseconds(200));

            Assert.True(result.Succeeded, result.Error);
            Assert.Equal("0.157.0", result.Catalog.ClientVersion);
            Assert.Equal("cache", result.ClientVersionSource);
        }

        [Fact]
        public async Task UpdateAsync_ProbeFailure_ReusesTheLastProbedEffortSet()
        {
            var probedBefore = new ModelCatalog(new[] { Entry("gpt-6-astra", 1, new[] { "low", "extreme" }) },
                new[] { "none", "low", "medium", "high", "xhigh", "max", "extreme" },
                ModelCatalog.SourceChatGpt, FixedNow.AddDays(-1), "0.156.0");
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK,
                "{\"models\":[{\"slug\":\"gpt-6-astra\",\"visibility\":\"list\",\"priority\":1,"
                + "\"supported_reasoning_levels\":[{\"effort\":\"low\"},{\"effort\":\"extreme\"}]}]}");
            fake.QueueText(HttpStatusCode.ServiceUnavailable, "busy");

            var result = await RunAsync(fake, probedBefore);

            Assert.True(result.Succeeded, result.Error);
            Assert.False(result.ServerEffortsProbed);
            Assert.Equal(probedBefore.ServerEfforts, result.Catalog.ServerEfforts);
            Assert.Equal(new[] { "None", "Low", "Extreme" }, result.Catalog.EffortsFor("gpt-6-astra"));
        }

        [Fact]
        public async Task UpdateAsync_IgnoresAProbedSetThatSharesNothingWithTheCatalog()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueJson(HttpStatusCode.BadRequest,
                "{\"error\":{\"message\":\"Invalid value: 'x'. Supported values are: 'alpha', 'beta'.\",\"param\":\"reasoning.effort\"}}");

            var result = await RunAsync(fake, Current("0.155.1"));

            Assert.False(result.ServerEffortsProbed);
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh", "Max" }, result.Catalog.EffortsFor("gpt-6-astra"));
        }

        [Fact]
        public async Task UpdateAsync_ProbeFailure_StillSucceeds_WithTheKnownEffortSet()
        {
            var fake = new FakeHttpMessageHandler();
            QueueGitHub(fake, "rust-v0.156.0");
            fake.QueueJson(HttpStatusCode.OK, LiveModelsResponse);
            fake.QueueText(HttpStatusCode.ServiceUnavailable, "busy");

            var result = await RunAsync(fake, Current("0.155.1"));

            Assert.True(result.Succeeded, result.Error);
            Assert.False(result.ServerEffortsProbed);
            Assert.Null(result.Catalog.ServerEfforts);
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh", "Max" }, result.Catalog.EffortsFor("gpt-6-astra"));
        }
    }
}
