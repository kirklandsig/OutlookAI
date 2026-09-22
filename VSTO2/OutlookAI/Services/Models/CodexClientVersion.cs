using System.Globalization;
using System.Text.RegularExpressions;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// The Codex CLI version sent as <c>client_version</c> on the model-catalog
    /// request. The backend only lists models whose minimum client version is at
    /// or below this value (0.154.0 hides gpt-6-sol/-luna; 0.155.x shows them), so
    /// it has to track real Codex releases rather than stay pinned.
    /// </summary>
    public static class CodexClientVersion
    {
        /// <summary>Latest released Codex CLI when this build was made (2026-09-22).</summary>
        public const string BuiltInFloor = "0.155.1";

        private static readonly Regex Pattern = new Regex(
            @"^(?:rust-)?v?(\d{1,9})\.(\d{1,9})\.(\d{1,9})(?:[-+][0-9A-Za-z.+\-]*)?\z",
            RegexOptions.CultureInvariant);

        /// <summary>
        /// <c>rust-v0.155.1</c> (the openai/codex tag shape), <c>v0.155.1</c>, or
        /// <c>0.155.1-alpha.2</c> → <c>0.155.1</c>. Anything else → null.
        /// </summary>
        public static string Normalize(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            var match = Pattern.Match(raw.Trim());
            if (!match.Success) return null;
            int major, minor, patch;
            if (!TryPart(match.Groups[1].Value, out major)
                || !TryPart(match.Groups[2].Value, out minor)
                || !TryPart(match.Groups[3].Value, out patch))
            {
                return null;
            }
            return major.ToString(CultureInfo.InvariantCulture) + "."
                + minor.ToString(CultureInfo.InvariantCulture) + "."
                + patch.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Highest valid version among the candidates; null if none is valid.</summary>
        public static string Max(params string[] candidates)
        {
            string best = null;
            int[] bestParts = null;
            foreach (var candidate in candidates ?? new string[0])
            {
                var normalized = Normalize(candidate);
                if (normalized == null) continue;
                var parts = Parts(normalized);
                if (bestParts == null || Compare(parts, bestParts) > 0)
                {
                    best = normalized;
                    bestParts = parts;
                }
            }
            return best;
        }

        private static bool TryPart(string digits, out int value)
        {
            return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out value);
        }

        private static int[] Parts(string normalized)
        {
            var pieces = normalized.Split('.');
            return new[]
            {
                int.Parse(pieces[0], CultureInfo.InvariantCulture),
                int.Parse(pieces[1], CultureInfo.InvariantCulture),
                int.Parse(pieces[2], CultureInfo.InvariantCulture),
            };
        }

        private static int Compare(int[] a, int[] b)
        {
            for (int i = 0; i < 3; i++)
            {
                var c = a[i].CompareTo(b[i]);
                if (c != 0) return c;
            }
            return 0;
        }
    }
}
