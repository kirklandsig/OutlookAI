using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class DateBucketFormatterTests
    {
        [Fact]
        public void Format_UtcDateTime_ReturnsIsoDateOnly()
        {
            var dt = new DateTimeOffset(2026, 5, 14, 18, 32, 0, TimeSpan.Zero);
            Assert.Equal("2026-05-14", DateBucketFormatter.Format(dt));
        }

        [Fact]
        public void Format_NonUtcOffset_NormalizedToUtcBeforeDateExtract()
        {
            // 2026-05-14 02:00 +04:00 == 2026-05-13 22:00 UTC. Bucket should be 2026-05-13.
            var dt = new DateTimeOffset(2026, 5, 14, 2, 0, 0, TimeSpan.FromHours(4));
            Assert.Equal("2026-05-13", DateBucketFormatter.Format(dt));
        }

        [Fact]
        public void Format_MinValue_ReturnsSentinel()
        {
            // Items we cannot date should not pollute buckets; sentinel
            // keeps them visible without colliding with real days.
            Assert.Equal("(unknown date)", DateBucketFormatter.Format(DateTimeOffset.MinValue));
        }
    }
}
