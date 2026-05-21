using System;

namespace OutlookAI.Services.Updates
{
    public sealed class ReleaseInfo
    {
        public string Tag { get; set; }
        public string ReleasePageUrl { get; set; }
        public DateTimeOffset PublishedAt { get; set; }
        public string Body { get; set; }
        public string ZipAssetName { get; set; }
        public string ZipUrl { get; set; }
        public string ShaUrl { get; set; }
    }
}
