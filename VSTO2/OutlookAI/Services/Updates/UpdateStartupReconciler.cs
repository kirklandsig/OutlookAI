using System;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Runs on every Outlook startup. Reconciles the in-progress sentinel
    /// with the actually-installed version.json:
    ///  - if installed tag matches sentinel tag -> install succeeded, clear + log
    ///  - if installed tag differs and sentinel is older than 30 min -> assume aborted, clear + log
    ///  - otherwise leave the sentinel alone
    /// </summary>
    public static class UpdateStartupReconciler
    {
        public static readonly TimeSpan StaleAfter = TimeSpan.FromMinutes(30);

        public static void Reconcile(UpdateHistoryLog history)
        {
            try
            {
                if (!File.Exists(UpdatePaths.InProgressSentinel)) return;
                var sentinelTag = (File.ReadAllText(UpdatePaths.InProgressSentinel) ?? "").Trim();
                var installed = UpdateManifest.LoadFromInstallDir();

                if (!installed.IsDevBuild && string.Equals(installed.Tag, sentinelTag, StringComparison.Ordinal))
                {
                    File.Delete(UpdatePaths.InProgressSentinel);
                    history?.Append("install", "succeeded", sentinelTag, "");
                    return;
                }

                var age = DateTime.UtcNow - File.GetLastWriteTimeUtc(UpdatePaths.InProgressSentinel);
                if (age > StaleAfter)
                {
                    File.Delete(UpdatePaths.InProgressSentinel);
                    history?.Append("install", "aborted", sentinelTag, "sentinel stale > 30 min");
                }
            }
            catch (Exception ex)
            {
                // Best-effort; never break Outlook startup over a stale sentinel.
                // Log to TraceLog so the maintainer at least sees the failure mode
                // if a corrupt version.json or locked sentinel ever trips this path.
                try { OutlookAI.Diagnostics.TraceLog.Write("Reconcile error: " + ex, "Updater"); } catch { }
            }
        }
    }
}
