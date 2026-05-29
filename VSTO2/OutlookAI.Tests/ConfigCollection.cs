using Xunit;

namespace OutlookAI.Tests
{
    /// <summary>
    /// xUnit collection used to serialize test classes that read or mutate the
    /// shared static state on <c>Config</c> (e.g. <c>ReasoningEffort</c>,
    /// <c>Model</c>, <c>MaxBulkExportRows</c>) or call <c>Config.ResetDefaults</c>
    /// / <c>Config.LoadConfigFromPaths</c>. Without this, <c>ConfigTests</c> and
    /// <c>CodexChatServiceMultiRoundTests</c> can race: one class's load/reset
    /// stomps the static field another class is asserting on. Mirrors
    /// <c>Services/Updates/UpdatePathsCollection.cs</c>. Members opt in via
    /// <c>[Collection("Config")]</c>; <c>DisableParallelization</c> runs them
    /// one at a time with no shared fixture.
    /// </summary>
    [CollectionDefinition("Config", DisableParallelization = true)]
    public sealed class ConfigCollection
    {
    }
}
