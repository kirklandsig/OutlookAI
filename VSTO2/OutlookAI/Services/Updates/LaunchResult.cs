namespace OutlookAI.Services.Updates
{
    public abstract class LaunchResult { }

    public sealed class Launched : LaunchResult
    {
        public int Pid { get; set; }
    }

    public sealed class UacDeclined : LaunchResult { }

    public sealed class LaunchFailed : LaunchResult
    {
        public string Detail { get; set; }
    }
}
