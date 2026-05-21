using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class VersionComparatorTests
    {
        [Theory]
        [InlineData("v2.0.0",  "v2.1.0",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.0.0",  "v2.0.1",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.0.0",  "v3.0.0",  UpdateAvailability.NewerAvailable)]
        [InlineData("v2.1.0",  "v2.1.0",  UpdateAvailability.NoUpdate)]
        [InlineData("v2.1.0",  "v2.0.9",  UpdateAvailability.OlderThanInstalled)]
        [InlineData("2.1.0",   "v2.1.0",  UpdateAvailability.NoUpdate)]                // missing v prefix
        [InlineData("v2.1.0",  "2.1.0",   UpdateAvailability.NoUpdate)]                // missing v prefix
        [InlineData("v2.1.0-beta.1", "v2.1.0", UpdateAvailability.NewerAvailable)]     // beta < release
        [InlineData("v2.1.0", "v2.1.0-beta.1", UpdateAvailability.OlderThanInstalled)] // release > beta
        [InlineData("v2.1.0-beta.1", "v2.1.0-beta.2", UpdateAvailability.NewerAvailable)]
        public void Compare_KnownPairs_ReturnsExpected(string installed, string latest, UpdateAvailability expected)
        {
            Assert.Equal(expected, VersionComparator.Compare(installed, latest));
        }

        [Theory]
        [InlineData("(dev build)", "v2.1.0")]
        [InlineData("v2.1.0",      "garbage")]
        [InlineData(null,          "v2.1.0")]
        [InlineData("v2.1.0",      "")]
        public void Compare_UnparseableInputs_ReturnsNotComparable(string installed, string latest)
        {
            Assert.Equal(UpdateAvailability.NotComparable, VersionComparator.Compare(installed, latest));
        }
    }
}
