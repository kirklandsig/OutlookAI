using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal enum PageKind { DocumentPage, NotDocument }

    internal enum PageColor { Grayscale, Color }

    internal enum CleanupStrength { Light, Normal, Strong }

    /// <summary>
    /// What the AI decided about one image, after validation. Corners are
    /// fractions of the image width and height: TL, TR, BR, BL as the text reads.
    /// </summary>
    internal sealed class PagePlan
    {
        public const string SchemaName = "page_plan";
        private const double Slack = 0.02;
        private const double MinAreaFraction = 0.05;

        private const string SchemaJson =
            "{\"type\":\"object\",\"additionalProperties\":false,"
            + "\"required\":[\"kind\",\"reason\",\"corners\",\"color\",\"strength\"],"
            + "\"properties\":{"
            + "\"kind\":{\"type\":\"string\",\"enum\":[\"document_page\",\"not_document\"]},"
            + "\"reason\":{\"type\":\"string\"},"
            + "\"corners\":{\"type\":\"array\",\"minItems\":4,\"maxItems\":4,\"items\":"
            + "{\"type\":\"array\",\"minItems\":2,\"maxItems\":2,\"items\":{\"type\":\"number\"}}},"
            + "\"color\":{\"type\":\"string\",\"enum\":[\"grayscale\",\"color\"]},"
            + "\"strength\":{\"type\":\"string\",\"enum\":[\"light\",\"normal\",\"strong\"]}}}";

        private PagePlan()
        {
        }

        public PageKind Kind { get; private set; }
        public string Reason { get; private set; }
        public PointD[] Corners { get; private set; }
        public PageColor Color { get; private set; }
        public CleanupStrength Strength { get; private set; }

        /// <summary>False when the AI failed or its outline was implausible; the image is then used uncropped.</summary>
        public bool Analyzed { get; private set; }

        public static PointD[] FullImage =>
            new[] { new PointD(0, 0), new PointD(1, 0), new PointD(1, 1), new PointD(0, 1) };

        public static JObject Schema() => JObject.Parse(SchemaJson);

        public static PagePlan Fallback() => new PagePlan
        {
            Kind = PageKind.DocumentPage,
            Reason = "",
            Corners = FullImage,
            Color = PageColor.Grayscale,
            Strength = CleanupStrength.Normal,
            Analyzed = false,
        };

        public static PagePlan FromJson(string text)
        {
            var o = ParseObject(text);
            if (o == null) return Fallback();

            PageKind kind;
            PageColor color;
            CleanupStrength strength;
            if (!TryKind(StringOf(o["kind"]), out kind)
                || !TryColor(StringOf(o["color"]), out color)
                || !TryStrength(StringOf(o["strength"]), out strength))
            {
                return Fallback();
            }

            var plan = new PagePlan
            {
                Kind = kind,
                Reason = (StringOf(o["reason"]) ?? "").Trim(),
                Color = color,
                Strength = strength,
            };
            var corners = ReadCorners(o["corners"] as JArray);
            if (corners != null && Quad.IsConvexClockwise(corners) && Quad.Area(corners) >= MinAreaFraction)
            {
                plan.Corners = corners;
                plan.Analyzed = true;
            }
            else
            {
                plan.Corners = FullImage;
                plan.Analyzed = false;
            }
            return plan;
        }

        // Strict output is plain JSON; a fence or a sentence around it is tolerated.
        private static JObject ParseObject(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try
            {
                return JObject.Parse(text.Substring(start, end - start + 1));
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static PointD[] ReadCorners(JArray array)
        {
            if (array == null || array.Count != 4) return null;
            var corners = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                var pair = array[i] as JArray;
                if (pair == null || pair.Count != 2) return null;
                double x, y;
                if (!TryNumber(pair[0], out x) || !TryNumber(pair[1], out y)) return null;
                if (x < -Slack || x > 1 + Slack || y < -Slack || y > 1 + Slack) return null;
                corners[i] = new PointD(Clamp(x), Clamp(y));
            }
            return corners;
        }

        // Only a JSON string counts; an object or array in a string field is not a value.
        private static string StringOf(JToken token) =>
            token != null && token.Type == JTokenType.String ? (string)token : null;

        private static bool TryNumber(JToken token, out double value)
        {
            value = 0;
            if (token == null || (token.Type != JTokenType.Float && token.Type != JTokenType.Integer)) return false;
            value = token.Value<double>();
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static double Clamp(double v) => v < 0 ? 0 : v > 1 ? 1 : v;

        private static bool TryKind(string s, out PageKind kind)
        {
            kind = PageKind.DocumentPage;
            if (s == "document_page") return true;
            if (s == "not_document") { kind = PageKind.NotDocument; return true; }
            return false;
        }

        private static bool TryColor(string s, out PageColor color)
        {
            color = PageColor.Grayscale;
            if (s == "grayscale") return true;
            if (s == "color") { color = PageColor.Color; return true; }
            return false;
        }

        private static bool TryStrength(string s, out CleanupStrength strength)
        {
            strength = CleanupStrength.Normal;
            if (s == "normal") return true;
            if (s == "light") { strength = CleanupStrength.Light; return true; }
            if (s == "strong") { strength = CleanupStrength.Strong; return true; }
            return false;
        }
    }
}
