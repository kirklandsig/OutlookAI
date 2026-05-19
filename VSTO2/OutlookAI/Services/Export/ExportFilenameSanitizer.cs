using System;
using System.IO;
using System.Text;

namespace OutlookAI.Services.Export
{
    public static class ExportFilenameSanitizer
    {
        private const string DefaultStem = "OutlookAI-Report";
        private const int MaxStemLength = 80;

        public static string Build(string hint, string extension, DateTimeOffset now, Func<string, bool> exists)
        {
            if (exists == null) throw new ArgumentNullException(nameof(exists));

            var stem = SanitizeStem(hint);
            var normalizedExtension = NormalizeExtension(extension);
            var timestamp = now.ToString("yyyy-MM-dd-HHmm");
            var baseName = stem + "-" + timestamp;
            var candidate = baseName + normalizedExtension;

            if (!exists(candidate)) return candidate;

            for (var suffix = 2; suffix <= 999; suffix++)
            {
                candidate = baseName + "-" + suffix + normalizedExtension;
                if (!exists(candidate)) return candidate;
            }

            return baseName + "-" + now.UtcDateTime.Ticks + normalizedExtension;
        }

        private static string SanitizeStem(string hint)
        {
            var builder = new StringBuilder();
            var previousDash = false;
            foreach (var c in hint ?? "")
            {
                var replacement = IsInvalidStemChar(c) ? '-' : c;
                if (replacement == '-')
                {
                    if (previousDash) continue;
                    previousDash = true;
                }
                else
                {
                    previousDash = false;
                }

                builder.Append(replacement);
            }

            var stem = builder.ToString().Trim('-', '.', ' ');
            if (stem.Length > MaxStemLength) stem = stem.Substring(0, MaxStemLength).Trim('-', '.', ' ');
            return stem.Length == 0 ? DefaultStem : stem;
        }

        private static bool IsInvalidStemChar(char c)
        {
            return char.IsControl(c) || char.IsWhiteSpace(c) || Array.IndexOf(Path.GetInvalidFileNameChars(), c) >= 0;
        }

        private static string NormalizeExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension)) return "";
            var trimmed = extension.Trim();
            return trimmed.StartsWith(".", StringComparison.Ordinal) ? trimmed : "." + trimmed;
        }
    }
}
