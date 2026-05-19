using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookCreateDraftToolTests
    {
        [Fact]
        public async Task Execute_PassesArgsAndProjectsDraftId()
        {
            CreateDraftArgs observed = null;
            var surface = new Surface
            {
                OnCreate = a =>
                {
                    observed = a;
                    return new CreatedDraft { DraftId = "d1", Location = "Drafts" };
                }
            };
            var tool = new OutlookCreateDraftTool();
            var json = await tool.ExecuteAsync(
                "{\"subject\":\"Hi\",\"body_plaintext\":\"body\",\"to\":[\"a@b.com\"],\"in_reply_to_message_id\":\"m1\"}",
                surface, CancellationToken.None);
            Assert.Contains("\"draft_id\":\"d1\"", json);
            Assert.Contains("\"location\":\"Drafts\"", json);
            Assert.Equal("Hi", observed.Subject);
            Assert.Equal("body", observed.BodyPlaintext);
            Assert.Equal("a@b.com", observed.To[0]);
            Assert.Equal("m1", observed.InReplyToMessageId);
        }

        [Fact]
        public async Task Execute_ReturnsErrorWhenRequiredFieldsMissing()
        {
            var surface = new Surface { OnCreate = _ => new CreatedDraft() };
            var tool = new OutlookCreateDraftTool();
            var json = await tool.ExecuteAsync(
                "{\"subject\":\"\"}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<CreateDraftArgs, CreatedDraft> OnCreate { get; set; }
            public override CreatedDraft CreateDraft(CreateDraftArgs args) => OnCreate(args);
        }
    }
}
