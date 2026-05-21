using System;
using System.IO;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    [Collection("UpdatePaths")]
    public class UpdateStartupReconcilerTests : IDisposable
    {
        private readonly string _tempBase;
        private readonly string _originalBaseDir;
        private readonly string _originalInstalledVersionJson;

        public UpdateStartupReconcilerTests()
        {
            _tempBase = Path.Combine(Path.GetTempPath(), "updater-startup-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(_tempBase);

            _originalBaseDir = UpdatePaths.BaseUpdatesDir;
            _originalInstalledVersionJson = UpdatePaths.InstalledVersionJson;

            UpdatePaths.BaseUpdatesDir = Path.Combine(_tempBase, "Updates");
            UpdatePaths.InstalledVersionJson = Path.Combine(_tempBase, "version.json");
            Directory.CreateDirectory(UpdatePaths.BaseUpdatesDir);
        }

        public void Dispose()
        {
            UpdatePaths.BaseUpdatesDir = _originalBaseDir;
            UpdatePaths.InstalledVersionJson = _originalInstalledVersionJson;
            try { Directory.Delete(_tempBase, true); } catch { }
        }

        private void WriteSentinel(string tag, DateTime? lastWriteUtc = null)
        {
            File.WriteAllText(UpdatePaths.InProgressSentinel, tag);
            if (lastWriteUtc.HasValue) File.SetLastWriteTimeUtc(UpdatePaths.InProgressSentinel, lastWriteUtc.Value);
        }

        private void WriteInstalledTag(string tag)
        {
            File.WriteAllText(UpdatePaths.InstalledVersionJson,
                "{\"tag\":\"" + tag + "\",\"commit\":\"-\",\"build_date\":\"2026-01-01T00:00:00Z\",\"repo\":\"x/y\"}");
        }

        [Fact]
        public void Reconcile_InstalledMatchesSentinel_ClearsAndLogsSuccess()
        {
            WriteSentinel("v2.1.0");
            WriteInstalledTag("v2.1.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            var entries = log.ReadAll();
            Assert.Contains(entries, e => e.Action == "install" && e.Result == "succeeded" && e.Tag == "v2.1.0");
        }

        [Fact]
        public void Reconcile_InstalledStillOld_SentinelStaleMoreThan30Min_ClearsAndLogsAborted()
        {
            WriteSentinel("v2.1.0", DateTime.UtcNow.AddMinutes(-31));
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            var entries = log.ReadAll();
            Assert.Contains(entries, e => e.Action == "install" && e.Result == "aborted" && e.Tag == "v2.1.0");
        }

        [Fact]
        public void Reconcile_InstalledStillOld_SentinelFresh_LeavesSentinelAlone()
        {
            WriteSentinel("v2.1.0", DateTime.UtcNow.AddMinutes(-5));
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.True(File.Exists(UpdatePaths.InProgressSentinel));
            Assert.Empty(log.ReadAll());
        }

        [Fact]
        public void Reconcile_NoSentinel_NoOp()
        {
            WriteInstalledTag("v2.0.0");
            var log = new UpdateHistoryLog(Path.Combine(_tempBase, "history.json"));

            UpdateStartupReconciler.Reconcile(log);

            Assert.False(File.Exists(UpdatePaths.InProgressSentinel));
            Assert.Empty(log.ReadAll());
        }
    }
}
