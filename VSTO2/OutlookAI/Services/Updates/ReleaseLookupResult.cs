using System;

namespace OutlookAI.Services.Updates
{
    public abstract class ReleaseLookupResult { }

    public sealed class ReleaseFound : ReleaseLookupResult
    {
        public ReleaseInfo Info { get; set; }
    }

    public sealed class NoReleasesAvailable : ReleaseLookupResult { }

    public sealed class RateLimited : ReleaseLookupResult
    {
        public DateTimeOffset ResetAt { get; set; }
    }

    public sealed class NetworkError : ReleaseLookupResult
    {
        public string Detail { get; set; }
    }
}
