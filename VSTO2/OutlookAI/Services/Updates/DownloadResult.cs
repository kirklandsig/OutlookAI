namespace OutlookAI.Services.Updates
{
    public abstract class DownloadResult { }

    public sealed class DownloadSuccess : DownloadResult
    {
        public string StagingDir { get; set; }
        public string ExtractedDir { get; set; }
        public string InstallerScriptPath { get; set; }
        public string ExpectedSha256 { get; set; }
    }

    public sealed class HashMismatch : DownloadResult
    {
        public string Expected { get; set; }
        public string Actual { get; set; }
    }

    public sealed class DownloadFailed : DownloadResult
    {
        public string Detail { get; set; }
    }

    public sealed class MissingInstallerScript : DownloadResult { }

    public sealed class Cancelled : DownloadResult { }
}
