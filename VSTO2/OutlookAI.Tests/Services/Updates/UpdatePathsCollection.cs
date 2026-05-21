using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    /// <summary>
    /// xUnit collection used to serialize test classes that mutate the shared
    /// static state on <c>UpdatePaths</c> (e.g. <c>BaseUpdatesDir</c>,
    /// <c>InstalledVersionJson</c>). Members opt in via
    /// <c>[Collection("UpdatePaths")]</c> on the class. The
    /// <c>DisableParallelization</c> flag is the explicit way to ask xUnit
    /// to run them one at a time; we don't need any shared fixture.
    /// </summary>
    [CollectionDefinition("UpdatePaths", DisableParallelization = true)]
    public sealed class UpdatePathsCollection
    {
    }
}
