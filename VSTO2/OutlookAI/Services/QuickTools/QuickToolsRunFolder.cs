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
            PurgeOlderThan(nowUtc.AddDays(-1));
            var dir = Path.Combine(Root, nowUtc.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)
                + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static void Delete(string dir)
        {
            try
            {
                if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
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
