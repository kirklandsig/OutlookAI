using System;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Single source of truth for the on-disk locations the updater uses.
    /// Tests can override BaseUpdatesDir to point at a temp folder.
    /// </summary>
    public static class UpdatePaths
    {
        /// <summary>
        /// Root for staged downloads, one subdir per release tag.
        /// Defaults to %LOCALAPPDATA%\OutlookAI\Updates.
        /// </summary>
        public static string BaseUpdatesDir { get; set; } =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "Updates");

        /// <summary>
        /// Location of the installed version.json. Read at runtime to tell the
        /// updater what version is live. Uses %ProgramW6432% so a 32-bit Outlook
        /// process still resolves to C:\Program Files\OutlookAI (matching the
        /// installer's hard-coded install path), instead of C:\Program Files (x86)\.
        /// </summary>
        public static string InstalledVersionJson { get; set; } =
            Path.Combine(
                Environment.GetEnvironmentVariable("ProgramW6432")
                    ?? Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "OutlookAI",
                "version.json");

        /// <summary>
        /// Sentinel file written when an install is launched and cleared on
        /// next successful Outlook startup.
        /// </summary>
        public static string InProgressSentinel =>
            Path.Combine(BaseUpdatesDir, ".in-progress");

        /// <summary>
        /// Append-only structured log of update activity.
        /// </summary>
        public static string HistoryLog =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "update-history.json");
    }
}
