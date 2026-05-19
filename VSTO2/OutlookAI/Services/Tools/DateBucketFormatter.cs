using System;
using System.Globalization;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Produces a stable bucket key for messages grouped by calendar day
    /// (UTC). Outlook's ReceivedTime arrives as a local DateTime which
    /// our ToOffset wrapper turns into a DateTimeOffset; this helper
    /// normalizes to UTC, takes the date component, and formats as
    /// ISO-8601 (yyyy-MM-dd) so buckets sort lexically.
    /// </summary>
    public static class DateBucketFormatter
    {
        public const string UnknownDate = "(unknown date)";

        public static string Format(DateTimeOffset receivedAt)
        {
            if (receivedAt == DateTimeOffset.MinValue) return UnknownDate;
            return receivedAt.UtcDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
    }
}
