using OutlookAI.Services;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class IdResolverTests
    {
        [Fact]
        public void Shorten_ProducesDeterministicShortIdPerEntryId()
        {
            var r = new IdResolver();
            var a = r.Shorten("00000000000000000000000000000000ABCDEF");
            var b = r.Shorten("00000000000000000000000000000000ABCDEF");
            Assert.Equal(a, b);
            Assert.True(a.Length <= 12);
        }

        [Fact]
        public void Shorten_DifferentEntryIdsProduceDifferentShortIds()
        {
            var r = new IdResolver();
            Assert.NotEqual(r.Shorten("AAAA"), r.Shorten("BBBB"));
        }

        [Fact]
        public void Resolve_RoundTripsKnownShortId()
        {
            var r = new IdResolver();
            var entryId = "00000000000000000000000000000000DEADBEEF";
            var shortId = r.Shorten(entryId);
            Assert.Equal(entryId, r.Resolve(shortId));
        }

        [Fact]
        public void Resolve_ThrowsOnUnknownShortId()
        {
            var r = new IdResolver();
            Assert.Throws<System.Collections.Generic.KeyNotFoundException>(
                () => r.Resolve("not-a-real-id"));
        }
    }
}
