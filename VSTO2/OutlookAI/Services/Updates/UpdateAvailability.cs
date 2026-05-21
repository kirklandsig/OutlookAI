namespace OutlookAI.Services.Updates
{
    public enum UpdateAvailability
    {
        NoUpdate,
        NewerAvailable,
        OlderThanInstalled,
        NotComparable,
        NoReleases,
    }
}
