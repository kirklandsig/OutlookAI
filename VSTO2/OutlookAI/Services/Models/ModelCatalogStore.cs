using System;
using System.IO;
using System.Text;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.Models
{
    /// <summary>Outcome of writing both cache copies; an error is null when that copy was written.</summary>
    public sealed class ModelCatalogSaveResult
    {
        public ModelCatalogSaveResult(
            string sharedError, string userError, ModelCatalog committed, ModelCatalog newerOnDisk = null)
        {
            SharedError = sharedError;
            UserError = userError;
            Committed = committed;
            NewerOnDisk = newerOnDisk;
        }

        public string SharedError { get; }
        public string UserError { get; }
        public bool SharedSaved => SharedError == null;
        public bool UserSaved => UserError == null;

        /// <summary>
        /// The catalog in effect afterwards: the newer of the two copies as
        /// committed (the one <see cref="ModelCatalogStore.Load"/> will return), or
        /// the one built in memory if neither could be written.
        /// </summary>
        public ModelCatalog Committed { get; }

        /// <summary>
        /// <see cref="ModelCatalogStore.Save"/> only: a copy already on disk that was
        /// newer than the one being saved and was left in place; null otherwise.
        /// </summary>
        public ModelCatalog NewerOnDisk { get; }
    }

    /// <summary>
    /// The models.json cache. The shared copy in ProgramData serves every user on
    /// an RDS host (the installer grants Authenticated Users Modify there); the
    /// per-user copy in LocalAppData is the fallback when the shared write is
    /// refused. <see cref="Load"/> returns the newer readable copy.
    /// </summary>
    public sealed class ModelCatalogStore
    {
        public const string FileName = "models.json";
        public const long MaxFileBytes = 2 * 1024 * 1024;

        // A copy stamped further ahead than this (clock skew aside) would outrank
        // every later refresh and the built-in list; it is ignored instead.
        private static readonly TimeSpan MaxClockSkew = TimeSpan.FromDays(1);
        private static readonly TimeSpan DefaultLockTimeout = TimeSpan.FromSeconds(5);

        private readonly Func<DateTimeOffset> _clock;
        private readonly TimeSpan _lockTimeout;

        public ModelCatalogStore()
            : this(DefaultSharedPath, DefaultUserPath)
        {
        }

        public ModelCatalogStore(
            string sharedPath, string userPath, Func<DateTimeOffset> clock = null, TimeSpan? lockTimeout = null)
        {
            SharedPath = sharedPath;
            UserPath = userPath;
            _clock = clock ?? (() => DateTimeOffset.UtcNow);
            _lockTimeout = lockTimeout ?? DefaultLockTimeout;
        }

        public static string DefaultSharedPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "OutlookAI", FileName);

        public static string DefaultUserPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", FileName);

        public string SharedPath { get; }
        public string UserPath { get; }

        /// <summary>The newer of the two readable copies (by fetched_at), or null.</summary>
        public ModelCatalog Load()
        {
            return ModelCatalog.Newest(TryLoad(SharedPath), TryLoad(UserPath));
        }

        /// <summary>
        /// Writes both copies, except where the copy on disk is already newer:
        /// that one is kept and returned in <see cref="ModelCatalogSaveResult.NewerOnDisk"/>.
        /// </summary>
        public ModelCatalogSaveResult Save(ModelCatalog catalog)
        {
            ModelCatalog newerOnDisk = null;
            var result = Commit(existing =>
            {
                if (existing != null && existing.FetchedAt > catalog.FetchedAt)
                {
                    newerOnDisk = ModelCatalog.Newest(newerOnDisk, existing);
                    return existing;
                }
                return catalog;
            });
            return new ModelCatalogSaveResult(result.SharedError, result.UserError, result.Committed, newerOnDisk);
        }

        /// <summary>
        /// For each copy: under its cross-process lock, reads what is on disk now,
        /// asks <paramref name="build"/> for the catalog to store (returning the
        /// existing one leaves the file untouched) and writes it. Merging at this
        /// point means a refresh another session committed meanwhile is built on,
        /// not overwritten. <paramref name="build"/> must be quick and side-effect free.
        /// </summary>
        public ModelCatalogSaveResult Commit(Func<ModelCatalog, ModelCatalog> build)
        {
            ModelCatalog shared;
            ModelCatalog user;
            var sharedError = TryCommit(SharedPath, build, out shared);
            var userError = TryCommit(UserPath, build, out user);
            // Same choice Load makes, so this session and its next start agree.
            return new ModelCatalogSaveResult(sharedError, userError, ModelCatalog.Newest(shared, user) ?? build(null));
        }

        private ModelCatalog TryLoad(string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            try
            {
                var info = new FileInfo(path);
                if (!info.Exists) return null;
                if (info.Length > MaxFileBytes)
                {
                    Trace("ignoring oversized " + path + " (" + info.Length + " bytes)");
                    return null;
                }
                var catalog = ModelCatalogJson.ParseCache(File.ReadAllText(path, Encoding.UTF8));
                if (catalog == null)
                {
                    Trace("ignoring unreadable " + path);
                    return null;
                }
                if (catalog.FetchedAt > _clock() + MaxClockSkew)
                {
                    Trace("ignoring " + path + ": fetched_at is in the future");
                    return null;
                }
                return catalog;
            }
            catch (Exception ex)
            {
                Trace("could not read '" + path + "': " + ex.Message);
                return null;
            }
        }

        // Read, merge and replace all happen under a cross-process lock, so two
        // RDS sessions refreshing at once can't lose each other's data. Unique
        // temp name for the same reason.
        private string TryCommit(string path, Func<ModelCatalog, ModelCatalog> build, out ModelCatalog committed)
        {
            committed = null;
            if (string.IsNullOrEmpty(path)) return "No path configured.";
            var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                using (FileLock.Acquire(path + ".lock", _lockTimeout))
                {
                    var existing = TryLoad(path);
                    var catalog = build(existing);
                    if (ReferenceEquals(catalog, existing))
                    {
                        Trace("kept " + path + " as is");
                    }
                    else
                    {
                        File.WriteAllText(temp, ModelCatalogJson.SerializeCache(catalog), new UTF8Encoding(false));
                        if (File.Exists(path))
                        {
                            File.Replace(temp, path, null);
                        }
                        else
                        {
                            File.Move(temp, path);
                        }
                    }
                    committed = catalog;
                }
                return null;
            }
            catch (TimeoutException)
            {
                Trace("gave up waiting to write '" + path + "'");
                return "The model list file is busy (another Outlook session is saving it). Try again.";
            }
            catch (Exception ex)
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                Trace("could not write '" + path + "': " + ex.Message);
                return ex.Message;
            }
        }

        private static void Trace(string message)
        {
            TraceLog.Write("ModelCatalogStore: " + message, "Models");
        }
    }
}
