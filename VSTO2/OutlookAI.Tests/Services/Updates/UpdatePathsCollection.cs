using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    /// <summary>
    /// xUnit collection that serializes test classes which mutate the shared
    /// static <see cref="OutlookAI.Services.Updates.UpdatePaths"/> properties.
    /// By default xUnit runs test classes in parallel, which would race on
    /// <c>UpdatePaths.BaseUpdatesDir</c> between e.g. UpdateDownloaderTests and
    /// UpdateStartupReconcilerTests. Classes opting into this collection run
    /// serially with respect to each other while still parallel across other
    /// collections.
    /// </summary>
    [CollectionDefinition("UpdatePaths")]
    public sealed class UpdatePathsCollection : ICollectionFixture<object>
    {
        // Empty: this exists solely to opt classes into a shared collection so
        // they run serially. The fixture object is not used.
    }
}
