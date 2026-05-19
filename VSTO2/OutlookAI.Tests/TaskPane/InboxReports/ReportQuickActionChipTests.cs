using System.Linq;
using OutlookAI.TaskPane.InboxReports;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxReports
{
    public class ReportQuickActionChipTests
    {
        private readonly System.Collections.Generic.IReadOnlyList<ReportQuickActionChip> _chips
            = ReportQuickActionChip.Defaults();

        [Fact]
        public void Defaults_ReturnsSixChips()
        {
            Assert.Equal(6, _chips.Count);
        }

        [Fact]
        public void Defaults_EachChipHasLabelAndTemplate()
        {
            foreach (var c in _chips)
            {
                Assert.False(string.IsNullOrWhiteSpace(c.Label));
                Assert.False(string.IsNullOrWhiteSpace(c.TemplateText));
                Assert.True(c.TemplateText.Length > 40,
                    "Chip template too short: " + c.Label);
            }
        }

        [Fact]
        public void Defaults_OrderingMatchesSpec()
        {
            // Spec defines the order: Digest, Conversation, Action items,
            // Project status, Stats, Out-of-office.
            Assert.Contains("digest", _chips[0].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("conversation", _chips[1].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("action", _chips[2].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("project", _chips[3].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stats", _chips[4].Label, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("out", _chips[5].Label, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ConversationChip_TemplateMentionsPersonPlaceholder()
        {
            Assert.Contains("[name or email]", _chips[1].TemplateText);
        }

        [Fact]
        public void ProjectChip_TemplateMentionsTopicPlaceholder()
        {
            Assert.Contains("[topic", _chips[3].TemplateText);
        }

        [Fact]
        public void OutOfOfficeChip_TemplateMentionsDatePlaceholders()
        {
            Assert.Contains("[start date]", _chips[5].TemplateText);
            Assert.Contains("[end date]", _chips[5].TemplateText);
        }

        [Fact]
        public void StatsChip_TemplateMentionsAggregateTool()
        {
            Assert.Contains("outlook_aggregate_messages", _chips[4].TemplateText);
        }

        [Fact]
        public void Defaults_LabelsAreUnique()
        {
            Assert.Equal(_chips.Count, _chips.Select(c => c.Label).Distinct().Count());
        }
    }
}
