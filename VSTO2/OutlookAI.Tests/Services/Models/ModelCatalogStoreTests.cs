using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OutlookAI.Services.Models;
using Xunit;
using static OutlookAI.Tests.Helpers.TestCatalogs;

namespace OutlookAI.Tests.Services.Models
{
    public class ModelCatalogStoreTests : IDisposable
    {
        private readonly string _dir;
        private readonly string _shared;
        private readonly string _user;

        public ModelCatalogStoreTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "outlookai-models", Path.GetRandomFileName());
            _shared = Path.Combine(_dir, "shared", "models.json");
            _user = Path.Combine(_dir, "user", "models.json");
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private static ModelCatalog Fetched(string slug, DateTimeOffset fetchedAt)
        {
            return new ModelCatalog(new[] { Entry(slug, 1, new[] { "low", "max" }) },
                new[] { "low", "max" }, ModelCatalog.SourceChatGpt, fetchedAt, "0.155.1");
        }

        [Fact]
        public void Save_WritesBothCopies_AndLoadRoundTrips()
        {
            var store = new ModelCatalogStore(_shared, _user);

            var result = store.Save(Fetched("gpt-7", FixedNow));

            Assert.True(result.SharedSaved);
            Assert.True(result.UserSaved);
            Assert.True(File.Exists(_shared));
            Assert.True(File.Exists(_user));
            var loaded = store.Load();
            Assert.Equal(new[] { "gpt-7" }, loaded.ListedSlugs);
            Assert.Equal(FixedNow, loaded.FetchedAt);
        }

        [Fact]
        public void Load_PrefersTheNewerCopy()
        {
            new ModelCatalogStore(_shared, null).Save(Fetched("older-shared", FixedNow));
            new ModelCatalogStore(null, _user).Save(Fetched("newer-user", FixedNow.AddHours(1)));

            Assert.Equal("newer-user", new ModelCatalogStore(_shared, _user).Load().ListedSlugs.Single());

            new ModelCatalogStore(_shared, null).Save(Fetched("newest-shared", FixedNow.AddHours(2)));

            Assert.Equal("newest-shared", new ModelCatalogStore(_shared, _user).Load().ListedSlugs.Single());
        }

        [Fact]
        public void Load_SkipsCorruptOrMissingCopies()
        {
            var store = new ModelCatalogStore(_shared, _user);
            Assert.Null(store.Load());

            new ModelCatalogStore(null, _user).Save(Fetched("good", FixedNow));
            Directory.CreateDirectory(Path.GetDirectoryName(_shared));
            File.WriteAllText(_shared, "{ this is not json");

            Assert.Equal("good", store.Load().ListedSlugs.Single());
        }

        [Fact]
        public void Load_SkipsCopiesDatedInTheFuture()
        {
            // A copy stamped ahead of the clock would otherwise outrank every
            // later refresh and the built-in list indefinitely.
            new ModelCatalogStore(_shared, null).Save(Fetched("from-the-future", FixedNow.AddDays(30)));
            new ModelCatalogStore(null, _user).Save(Fetched("current", FixedNow.AddHours(-1)));

            var store = new ModelCatalogStore(_shared, _user, () => FixedNow);

            Assert.Equal("current", store.Load().ListedSlugs.Single());
        }

        [Fact]
        public void Load_SkipsOversizedFiles()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_shared));
            File.WriteAllText(_shared, new string(' ', (int)ModelCatalogStore.MaxFileBytes + 1));

            Assert.Null(new ModelCatalogStore(_shared, null).Load());
        }

        [Fact]
        public void Save_ReportsEachCopyIndependently()
        {
            // A directory squatting on the shared file path makes that write fail.
            Directory.CreateDirectory(_shared);
            var store = new ModelCatalogStore(_shared, _user);

            var result = store.Save(Fetched("m", FixedNow));

            Assert.False(result.SharedSaved);
            Assert.False(string.IsNullOrEmpty(result.SharedError));
            Assert.True(result.UserSaved);
            Assert.Null(result.UserError);
        }

        [Fact]
        public void Save_DoesNotReplaceANewerCopy_AndHandsItBack()
        {
            var store = new ModelCatalogStore(_shared, _user);
            store.Save(Fetched("newer", FixedNow.AddMinutes(5)));

            var result = store.Save(Fetched("older", FixedNow));

            Assert.True(result.SharedSaved);   // nothing failed; there was just nothing to write
            Assert.Equal("newer", result.NewerOnDisk.ListedSlugs.Single());
            Assert.Equal("newer", store.Load().ListedSlugs.Single());
        }

        [Fact]
        public void Save_ReportsTheNewestCommittedCopy_EvenWhenOnlyThePerUserCopyIsNewer()
        {
            // Another refresh reached only the per-user fallback: the shared copy
            // gets ours, but the list in effect (what Load returns) is theirs.
            new ModelCatalogStore(_shared, null).Save(Fetched("old-shared", FixedNow.AddHours(-1)));
            new ModelCatalogStore(null, _user).Save(Fetched("newer-user", FixedNow.AddMinutes(1)));
            var store = new ModelCatalogStore(_shared, _user);

            var result = store.Save(Fetched("ours", FixedNow));

            Assert.Equal("newer-user", result.Committed.ListedSlugs.Single());
            Assert.Equal("newer-user", store.Load().ListedSlugs.Single());
        }

        [Fact]
        public void Save_WaitsForAnotherWriterToFinish()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_shared));
            var lockPath = _shared + ".lock";
            var held = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose);
            var release = Task.Run(async () => { await Task.Delay(300); held.Dispose(); });
            var store = new ModelCatalogStore(_shared, null);

            var result = store.Save(Fetched("m", FixedNow));

            release.Wait();
            Assert.True(result.SharedSaved, result.SharedError);
            Assert.Equal("m", store.Load().ListedSlugs.Single());
            Assert.False(File.Exists(lockPath));
        }

        [Fact]
        public void Save_GivesUpOnALockThatIsNeverReleased()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_shared));
            var store = new ModelCatalogStore(_shared, _user, lockTimeout: TimeSpan.FromMilliseconds(200));
            using (new FileStream(_shared + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose))
            {
                var result = store.Save(Fetched("m", FixedNow));

                Assert.False(result.SharedSaved);
                Assert.Contains("busy", result.SharedError);
                Assert.True(result.UserSaved);
            }
        }

        [Fact]
        public void Save_Overwrites_AndLeavesNoTempFiles()
        {
            var store = new ModelCatalogStore(_shared, _user);
            store.Save(Fetched("first", FixedNow));
            store.Save(Fetched("second", FixedNow.AddMinutes(1)));

            Assert.Equal("second", store.Load().ListedSlugs.Single());
            Assert.Equal(new[] { "models.json" }, Directory.GetFiles(Path.GetDirectoryName(_shared)).Select(Path.GetFileName));
            Assert.Equal(new[] { "models.json" }, Directory.GetFiles(Path.GetDirectoryName(_user)).Select(Path.GetFileName));
        }

        [Fact]
        public void DefaultPaths_AreProgramDataAndLocalAppData()
        {
            Assert.Equal(
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "OutlookAI", "models.json"),
                ModelCatalogStore.DefaultSharedPath);
            Assert.Equal(
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", "models.json"),
                ModelCatalogStore.DefaultUserPath);
        }
    }
}
