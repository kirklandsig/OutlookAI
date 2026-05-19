using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class ExportPdfArgsParser
    {
        private const int MaxContentMarkdownLength = 250000;
        private const int MaxTitleLength = 200;
        private const int MaxSubtitleLength = 400;
        private const string DefaultFilenameHint = "OutlookAI-Report";

        public static ExportPdfArgs Parse(string argsJson)
        {
            JObject args;
            try
            {
                args = JObject.Parse(argsJson);
            }
            catch (JsonException ex)
            {
                throw InvalidArgs("Invalid JSON args: " + ex.Message);
            }

            var contentMarkdown = ParseContentMarkdown(args["content_markdown"]);
            var title = CleanOptionalString(args["title"], "title", MaxTitleLength);
            var subtitle = CleanOptionalString(args["subtitle"], "subtitle", MaxSubtitleLength);
            var filenameHint = CleanOptionalString(args["filename_hint"], "filename_hint", null) ?? title ?? DefaultFilenameHint;

            return new ExportPdfArgs
            {
                FilenameHint = filenameHint,
                ContentMarkdown = contentMarkdown,
                Title = title,
                Subtitle = subtitle,
            };
        }

        private static string ParseContentMarkdown(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) throw InvalidArgs("content_markdown is required");
            if (token.Type != JTokenType.String) throw InvalidArgs("content_markdown must be a string");

            var value = token.Value<string>();
            if (string.IsNullOrWhiteSpace(value)) throw InvalidArgs("content_markdown is required");
            if (value.Length > MaxContentMarkdownLength)
            {
                throw new ToolArgValidationException("content_too_large", "content_markdown must be <= " + MaxContentMarkdownLength + " characters");
            }

            return value;
        }

        private static string CleanOptionalString(JToken token, string fieldName, int? maxLength)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            if (token.Type != JTokenType.String) throw InvalidArgs(fieldName + " must be a string");

            var value = token.Value<string>()?.Trim();
            if (string.IsNullOrEmpty(value)) return null;

            if (maxLength.HasValue && value.Length > maxLength.Value)
            {
                return value.Substring(0, maxLength.Value);
            }

            return value;
        }

        private static ToolArgValidationException InvalidArgs(string message)
        {
            return new ToolArgValidationException("invalid_args", message);
        }
    }
}
