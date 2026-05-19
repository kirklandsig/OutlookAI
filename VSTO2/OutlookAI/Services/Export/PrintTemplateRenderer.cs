using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace OutlookAI.Services.Export
{
    public sealed class PrintTemplateRenderer
    {
        private static readonly Regex ImageRegex = new Regex(@"!\[([^\]]*)\]\([^)]+\)", RegexOptions.Compiled);
        private readonly string templateHtml;

        public PrintTemplateRenderer(string templateHtml)
        {
            if (templateHtml == null) throw new ArgumentNullException(nameof(templateHtml));
            this.templateHtml = templateHtml;
        }

        public static PrintTemplateRenderer LoadFromFile(string templatePath)
        {
            return new PrintTemplateRenderer(File.ReadAllText(templatePath));
        }

        public string Render(string title, string subtitle, string markdown, DateTimeOffset generatedAt)
        {
            var safeMarkdown = StripInlineImagesOutsideCodeBlocks(markdown ?? "");
            var markdownLiteral = JsonConvert.ToString(safeMarkdown, '"', StringEscapeHandling.EscapeHtml);
            var markdownInjection = "window.__OUTLOOKAI_MD__ = " + markdownLiteral + ";";

            return templateHtml
                .Replace("__TITLE_TEXT__", WebUtility.HtmlEncode(title ?? ""))
                .Replace("__SUBTITLE_TEXT__", WebUtility.HtmlEncode(subtitle ?? ""))
                .Replace("__GENERATED_AT__", generatedAt.ToLocalTime().ToString("MMMM d, yyyy h:mm tt", CultureInfo.InvariantCulture))
                .Replace("__MD_INJECT__", markdownInjection);
        }

        private static string StripInlineImagesOutsideCodeBlocks(string markdown)
        {
            var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var inFence = false;
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    inFence = !inFence;
                }
                else if (!inFence)
                {
                    lines[i] = ImageRegex.Replace(lines[i], "[image: $1]");
                }
            }

            return string.Join("\n", lines);
        }
    }
}
