using System;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.InboxCopilot;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxCopilot
{
    public class InboxCopilotPromptBuilderTests
    {
        [Fact]
        public void NoSelection_PromptIncludesFolderAndUnreadCountOnly()
        {
            var prompt = InboxCopilotPromptBuilder.Build(
                folderName: "Inbox",
                unreadCount: 47,
                totalCount: 1284,
                selection: null);

            Assert.Contains("Inbox", prompt);
            Assert.Contains("47 unread", prompt);
            Assert.Contains("1284 total", prompt);
            Assert.DoesNotContain("Selected:", prompt);
        }

        [Fact]
        public void SingleSelection_PromptIncludesSelectedMessageBlock()
        {
            var sel = new CurrentSelectionResult
            {
                Folder = "Inbox",
                FolderId = "fld",
                Count = 1,
                Messages = new[]
                {
                    new MessageDetail
                    {
                        Id = "m1",
                        Subject = "Re: Q4 plan",
                        From = "Jane Doe <jane@acme.com>",
                        ReceivedAt = new DateTimeOffset(2026, 5, 17, 9, 14, 0, TimeSpan.Zero),
                        BodyPlaintext = "Hi team - thoughts on regional split. ",
                    }
                }
            };
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 47, 1284, sel);

            Assert.Contains("Selected:", prompt);
            Assert.Contains("Re: Q4 plan", prompt);
            Assert.Contains("Jane Doe", prompt);
        }

        [Fact]
        public void MultiSelection_PromptSummarizesCount()
        {
            var sel = new CurrentSelectionResult
            {
                Folder = "Inbox",
                FolderId = "fld",
                Count = 4,
                Messages = new MessageDetail[0],
            };
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 47, 1284, sel);

            Assert.Contains("4 messages selected", prompt);
            Assert.DoesNotContain("Re: Q4 plan", prompt);
        }

        [Fact]
        public void Prompt_AlwaysIncludesRolePreamble()
        {
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 0, 0, null);
            Assert.Contains("Inbox Copilot", prompt);
        }
    }
}
