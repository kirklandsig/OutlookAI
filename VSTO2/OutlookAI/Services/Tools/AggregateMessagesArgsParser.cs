using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class AggregateMessagesArgsParser
    {
        private const int TopNFloor = 1;
        private const int TopNCap = 100;

        public static AggregateMessagesArgs Parse(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            return new AggregateMessagesArgs
            {
                Scope = EnumOrDefault(args["scope"], "auto", "auto", "current_folder", "all_mail"),
                FolderId = Clean(args["folder_id"]),
                DateFrom = ParseDate(args["date_from"]),
                DateTo = ParseDate(args["date_to"]),
                From = Clean(args["from"]),
                SubjectContains = Clean(args["subject_contains"]),
                BodyContains = Clean(args["body_contains"]),
                GroupBy = EnumOrDefault(args["group_by"], "sender", "sender", "day", "folder"),
                TopN = ClampTopN(args["top_n"]?.Value<int>() ?? 10),
            };
        }

        private static int ClampTopN(int value)
        {
            if (value < TopNFloor) return TopNFloor;
            if (value > TopNCap) return TopNCap;
            return value;
        }

        private static string Clean(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            var value = ((string)token)?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string EnumOrDefault(JToken token, string fallback, params string[] allowed)
        {
            var value = Clean(token);
            if (value == null) return fallback;
            value = value.ToLowerInvariant();
            foreach (var allowedValue in allowed)
            {
                if (value == allowedValue) return value;
            }
            return fallback;
        }

        private static DateTimeOffset? ParseDate(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
            if (!DateTimeOffset.TryParse(
                (string)token,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out value))
            {
                return null;
            }
            return value;
        }
    }
}
