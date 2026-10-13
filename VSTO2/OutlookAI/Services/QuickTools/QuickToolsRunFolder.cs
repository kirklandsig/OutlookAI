using System;
using System.Globalization;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// One temp folder per run under %LOCALAPPDATA%\OutlookAI\QuickTools, deleted
    /// when the run ends; leftovers older than a day are purged on the next run.
    /// </summary>
    internal static class QuickToolsRunFolder
    {
        // Test seam.
        public static string Root { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", "QuickTools");

        public static string Create(DateTime nowUtc)
        {
            try
            {
                PurgeOlderThan(nowUtc.AddDays(-1));
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            var dir = Path.Combine(Root, nowUtc.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)
                + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            return dir;
        }

        /// <summary>Deletes a run folder; anything outside Root is left alone.</summary>
        public static void Delete(string dir)
        {
            try
            {
                if (!IsUnderRoot(dir)) return;
                if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static bool IsUnderRoot(string dir)
        {
            if (string.IsNullOrEmpty(dir)) return false;
            var root = Path.GetFullPath(Root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return Path.GetFullPath(dir).StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }

        internal static void PurgeOlderThan(DateTime cutoffUtc)
        {
            if (!Directory.Exists(Root)) return;
            foreach (var dir in Directory.GetDirectories(Root))
            {
                if (Directory.GetCreationTimeUtc(dir) < cutoffUtc) Delete(dir);
            }
        }
    }
}
