using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    [Collection("Config")]
    public class PagePlannerLiveTests : IDisposable
    {
        private readonly ConfigStateScope _scope = new ConfigStateScope();

        public void Dispose()
        {
            _scope.Dispose();
        }

        [Fact]
        public async Task Live_FindsTheCornersOfAPhotographedPage()
        {
            if (Environment.GetEnvironmentVariable("OUTLOOKAI_LIVE_TESTS") != "1") return;
            var authPath = Environment.GetEnvironmentVariable("OUTLOOKAI_LIVE_AUTH")
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "auth.json");
            Config.Model = "gpt-5.6-sol";
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };

            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, corners, SyntheticPages.Desk))
            using (var auth = new CodexAuthService(authPath))
            using (var chat = new CodexChatService(auth))
            {
                var plan = await new PagePlanner(chat).PlanAsync(PagePlanner.PrepareForAi(photo), CancellationToken.None);

                Assert.True(plan.Analyzed);
                Assert.Equal(PageKind.DocumentPage, plan.Kind);
                for (var i = 0; i < 4; i++)
                {
                    Assert.InRange(plan.Corners[i].X, corners[i].X / 1600 - 0.03, corners[i].X / 1600 + 0.03);
                    Assert.InRange(plan.Corners[i].Y, corners[i].Y / 1200 - 0.03, corners[i].Y / 1200 + 0.03);
                }
            }
        }
    }
}
