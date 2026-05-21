namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Pure static comparator used by the updater to decide whether the
    /// GitHub-reported release tag is newer than what is installed.
    /// </summary>
    public static class VersionComparator
    {
        public static UpdateAvailability Compare(string installedTag, string latestTag)
        {
            if (!SemVer.TryParse(installedTag, out var a)) return UpdateAvailability.NotComparable;
            if (!SemVer.TryParse(latestTag,    out var b)) return UpdateAvailability.NotComparable;

            var c = a.CompareTo(b);
            if (c < 0) return UpdateAvailability.NewerAvailable;
            if (c > 0) return UpdateAvailability.OlderThanInstalled;
            return UpdateAvailability.NoUpdate;
        }
    }
}
