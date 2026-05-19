namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Produces a stable bucket key for messages grouped by sender.
    /// Prefers the display name; falls back to a lowercased email so
    /// case differences do not split the same correspondent into two
    /// buckets. Returns a sentinel string when neither is available so
    /// downstream code never has to handle null bucket keys.
    /// </summary>
    public static class SenderKeyNormalizer
    {
        public const string UnknownSender = "(unknown sender)";

        public static string Normalize(string senderName, string senderEmail)
        {
            var trimmedName = (senderName ?? "").Trim();
            if (trimmedName.Length > 0) return trimmedName;
            var trimmedEmail = (senderEmail ?? "").Trim();
            if (trimmedEmail.Length > 0) return trimmedEmail.ToLowerInvariant();
            return UnknownSender;
        }
    }
}
