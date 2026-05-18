using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class TopNBucketSelectorTests
    {
        private static AggregationBucket B(string label, int count) =>
            new AggregationBucket { Label = label, Count = count };

        [Fact]
        public void TakeTop_OrdersByCountDescending()
        {
            var input = new[] { B("a", 3), B("b", 10), B("c", 1) };
            var result = TopNBucketSelector.TakeTop(input, 5);
            Assert.Equal(new[] { "b", "a", "c" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_ClampsToN()
        {
            var input = new[] { B("a", 5), B("b", 4), B("c", 3), B("d", 2) };
            var result = TopNBucketSelector.TakeTop(input, 2);
            Assert.Equal(2, result.Count);
            Assert.Equal(new[] { "a", "b" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_TiesBreakAlphabetically()
        {
            var input = new[] { B("zeta", 5), B("alpha", 5), B("mu", 5) };
            var result = TopNBucketSelector.TakeTop(input, 3);
            Assert.Equal(new[] { "alpha", "mu", "zeta" }, result.Select(r => r.Label).ToArray());
        }

        [Fact]
        public void TakeTop_NullOrEmptyInput_ReturnsEmpty()
        {
            Assert.Empty(TopNBucketSelector.TakeTop(null, 5));
            Assert.Empty(TopNBucketSelector.TakeTop(new AggregationBucket[0], 5));
        }

        [Fact]
        public void TakeTop_NonPositiveN_ReturnsEmpty()
        {
            var input = new[] { B("a", 5) };
            Assert.Empty(TopNBucketSelector.TakeTop(input, 0));
            Assert.Empty(TopNBucketSelector.TakeTop(input, -3));
        }
    }
}
