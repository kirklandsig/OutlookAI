using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Variants
{
    /// <summary>
    /// Extracts the Variants envelope out of an assistant turn's final text.
    /// The system prompt for the Variants flow requires:
    ///   ```json
    ///   { "variants": [ { "tone", "rationale", "subject", "body" }, ... ] }
    ///   ```
    /// but the parser tolerates code fences with or without newlines, missing
    /// language tag, and bare JSON (no fence at all).
    /// </summary>
    public sealed class VariantParser
    {
        // Permissive code-fence pattern:
        //   ``` optional 'json' tag, optional whitespace/newline, lazy capture,
        //   optional newline, closing ```. Lets us catch both well-formed and
        //   single-line outputs.
        private static readonly Regex FencedJson = new Regex(
            @"```(?:json)?\s*\r?\n?([\s\S]*?)\r?\n?\s*```",
            RegexOptions.Compiled);

        public IReadOnlyList<Variant> Parse(string assistantText)
        {
            if (string.IsNullOrWhiteSpace(assistantText))
            {
                return new Variant[0];
            }

            string json = assistantText;
            var m = FencedJson.Match(assistantText);
            if (m.Success)
            {
                json = m.Groups[1].Value;
            }

            JObject root;
            try
            {
                root = JObject.Parse(json);
            }
            catch
            {
                return new Variant[0];
            }

            var arr = root["variants"] as JArray;
            if (arr == null)
            {
                return new Variant[0];
            }

            var list = new List<Variant>();
            foreach (var v in arr)
            {
                list.Add(new Variant
                {
                    Tone = ToneExtensions.ClosestTo((string)v["tone"]),
                    Rationale = (string)v["rationale"] ?? "",
                    Subject = (string)v["subject"] ?? "",
                    Body = (string)v["body"] ?? "",
                });
            }
            return list;
        }
    }
}
