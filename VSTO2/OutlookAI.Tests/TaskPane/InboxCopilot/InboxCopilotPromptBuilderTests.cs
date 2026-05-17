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

        /// <summary>
        /// Pin the search-Tips steering block. After the SSE-delta bug fix
        /// the model CAN finally pass structured args, but it still needs
        /// explicit conversation-level guidance to map natural language
        /// onto from/date_from/date_to/etc. instead of dumping everything
        /// into 'query'. Regressing this block would re-introduce the
        /// "only finds the newest 25 messages" UX failure.
        /// </summary>
        [Fact]
        public void Prompt_IncludesSearchTipsBlock_SteersTowardStructuredFields()
        {
            var prompt = InboxCopilotPromptBuilder.Build("Inbox", 0, 0, null);

            Assert.Contains("Tips for searching", prompt);
            Assert.Contains("structured fields", prompt);
            // Tells the model NOT to send empty args (the cause of the
            // "always-Uber-email" symptom).
            Assert.Contains("empty argument object", prompt);
            // Steers sender, dates, and keyword placement explicitly.
            Assert.Contains("'from'", prompt);
            Assert.Contains("ISO-8601 UTC", prompt);
            Assert.Contains("date_to=2020-01-01T00:00:00Z", prompt);
            // Reminds model to follow up with read_message for full body.
            Assert.Contains("outlook_read_message", prompt);
        }
    }
}
