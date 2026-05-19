using System;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class PrintTemplateRendererTests
    {
        private static readonly DateTimeOffset FixedGeneratedAt = new DateTimeOffset(2026, 5, 18, 14, 30, 0, TimeSpan.Zero);

        [Fact]
        public void Render_SubstitutesTitleSubtitleMarkdownAndGeneratedAtTokens()
        {
            var renderer = new PrintTemplateRenderer(Template());

            var html = renderer.Render("Quarterly Report", "Inbox summary", "# Heading", FixedGeneratedAt);

            Assert.Contains("<title>Quarterly Report</title>", html);
            Assert.Contains("<p>Inbox summary</p>", html);
            Assert.Contains("May 18, 2026", html);
            Assert.Contains("window.__OUTLOOKAI_MD__ = \"# Heading\";", html);
            Assert.DoesNotContain("__TITLE_TEXT__", html);
            Assert.DoesNotContain("__SUBTITLE_TEXT__", html);
            Assert.DoesNotContain("__GENERATED_AT__", html);
            Assert.DoesNotContain("__MD_INJECT__", html);
        }

        [Fact]
        public void Render_EncodesTitleAndSubtitleHtml()
        {
            var renderer = new PrintTemplateRenderer(Template());

            var html = renderer.Render("<script>alert('x')</script>", "A & B", "", FixedGeneratedAt);

            Assert.Contains("&lt;script&gt;alert(&#39;x&#39;)&lt;/script&gt;", html);
            Assert.Contains("A &amp; B", html);
            Assert.DoesNotContain("<title><script>", html);
        }

        [Fact]
        public void Render_JsonMarkdownInjectionEscapesScriptBreakingContent()
        {
            var renderer = new PrintTemplateRenderer(Template());
            var markdown = "Quote: \"hello\" / slash\n</script><p>owned</p>";

            var html = renderer.Render("Title", "Subtitle", markdown, FixedGeneratedAt);

            Assert.Contains("window.__OUTLOOKAI_MD__ = ", html);
            Assert.Contains("\\u0022hello\\u0022", html);
            Assert.Contains("/ slash", html);
            Assert.Contains("\\n", html);
            Assert.Contains("\\u003c/script\\u003e", html);
            Assert.Equal(1, CountOccurrences(html, "</script>"));
        }

        [Fact]
        public void Render_StripsInlineImagesOutsideFencedCodeBlocks()
        {
            var renderer = new PrintTemplateRenderer(Template());

            var html = renderer.Render("Title", "Subtitle", "Before ![chart](https://example.test/chart.png) after", FixedGeneratedAt);

            Assert.Contains("Before [image: chart] after", html);
            Assert.DoesNotContain("![chart](https://example.test/chart.png)", html);
        }

        [Fact]
        public void Render_DoesNotStripImageSyntaxInsideFencedCodeBlocks()
        {
            var renderer = new PrintTemplateRenderer(Template());
            var markdown = "```\n![keep](image.png)\n```\n![strip](image.png)";

            var html = renderer.Render("Title", "Subtitle", markdown, FixedGeneratedAt);

            Assert.Contains("![keep](image.png)", html);
            Assert.Contains("[image: strip]", html);
        }

        [Fact]
        public void Render_NullSubtitleProducesEmptyParagraphAndNoTokenLeftovers()
        {
            var renderer = new PrintTemplateRenderer(Template());

            var html = renderer.Render("Title", null, "Body", FixedGeneratedAt);

            Assert.Contains("<p></p>", html);
            AssertNoTemplateTokens(html);
        }

        [Fact]
        public void Render_NullTitleAndMarkdownAreSafeEmptyStrings()
        {
            var renderer = new PrintTemplateRenderer(Template());

            var html = renderer.Render(null, "Subtitle", null, FixedGeneratedAt);

            Assert.Contains("<title></title>", html);
            Assert.Contains("window.__OUTLOOKAI_MD__ = \"\";", html);
        }

        [Fact]
        public void Constructor_NullTemplateThrows()
        {
            Assert.Throws<ArgumentNullException>(() => new PrintTemplateRenderer(null));
        }

        private static string Template()
        {
            return "<html><head><title>__TITLE_TEXT__</title></head><body><h1>__TITLE_TEXT__</h1><p>__SUBTITLE_TEXT__</p><time>__GENERATED_AT__</time><script>__MD_INJECT__</script></body></html>";
        }

        private static void AssertNoTemplateTokens(string html)
        {
            Assert.DoesNotContain("__TITLE_TEXT__", html);
            Assert.DoesNotContain("__SUBTITLE_TEXT__", html);
            Assert.DoesNotContain("__GENERATED_AT__", html);
            Assert.DoesNotContain("__MD_INJECT__", html);
        }

        private static int CountOccurrences(string value, string text)
        {
            var count = 0;
            var index = 0;
            while ((index = value.IndexOf(text, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += text.Length;
            }

            return count;
        }
    }
}
