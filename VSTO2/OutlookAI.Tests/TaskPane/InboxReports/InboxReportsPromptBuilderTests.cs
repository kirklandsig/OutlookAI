using OutlookAI.TaskPane.InboxReports;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxReports
{
    public class InboxReportsPromptBuilderTests
    {
        private readonly string _prompt = new InboxReportsPromptBuilder().Build();

        [Fact]
        public void Prompt_AlwaysIncludesRolePreamble()
        {
            Assert.Contains("mailbox reports assistant", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_MentionsBulkReadTool()
        {
            Assert.Contains("outlook_read_messages", _prompt);
        }

        [Fact]
        public void Prompt_MentionsAggregateTool()
        {
            Assert.Contains("outlook_aggregate_messages", _prompt);
        }

        [Fact]
        public void Prompt_TellsModelToAskBeforeUnresolvedPlaceholders()
        {
            // The chip templates include placeholders like "[name or email]";
            // the model must ask for clarification rather than calling tools
            // with literal brackets in the args.
            Assert.Contains("placeholder", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("ask", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_DescribesMarkdownAndConciseFormat()
        {
            Assert.Contains("markdown", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("concise", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Prompt_DescribesHeaderWithScopeAndCount()
        {
            Assert.Contains("header", _prompt, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("scope", _prompt, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
