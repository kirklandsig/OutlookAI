using OutlookAI.TaskPane.InboxCopilot;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxCopilot
{
    public class QuickActionChipTests
    {
        [Fact]
        public void NoSelection_ReturnsThreeStaticChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(0);
            Assert.Equal(3, chips.Count);
            Assert.Contains(chips, c => c.Label == "What needs my attention?");
            Assert.Contains(chips, c => c.Label == "Summarize unread");
            Assert.Contains(chips, c => c.Label == "Today's emails");
        }

        [Fact]
        public void SingleSelection_AddsSingleSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(1);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Summarize this thread");
            Assert.Contains(chips, c => c.Label == "Draft a reply");
            // Static three still present:
            Assert.Contains(chips, c => c.Label == "Today's emails");
        }

        [Fact]
        public void MultiSelection_AddsMultiSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(3);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Summarize all selected");
            Assert.Contains(chips, c => c.Label == "Triage selected");
            Assert.Contains(chips, c => c.Label == "What needs my attention?");
            Assert.DoesNotContain(chips, c => c.Label == "Summarize this thread");
            Assert.DoesNotContain(chips, c => c.Label == "Draft a reply");
        }

        [Fact]
        public void Prompts_AreNotEmpty()
        {
            foreach (var n in new[] { 0, 1, 5 })
            {
                var chips = QuickActionChip.ComputeChipsForSelectionCount(n);
                Assert.All(chips, c => Assert.False(string.IsNullOrWhiteSpace(c.Prompt)));
                Assert.All(chips, c => Assert.False(string.IsNullOrWhiteSpace(c.Label)));
            }
        }
    }
}
