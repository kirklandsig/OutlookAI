using System;
using System.Collections.Generic;
using System.Globalization;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Minimal semver parser for tags shaped like
    /// `[v]MAJOR.MINOR.PATCH[-PRERELEASE]`.
    /// Prerelease suffix sorts LOWER than the same numeric tuple without it,
    /// per SemVer 2.0.0. Prerelease itself is compared dot-separated:
    /// numeric identifiers numerically, others lexicographically.
    /// </summary>
    internal sealed class SemVer : IComparable<SemVer>
    {
        public int Major { get; }
        public int Minor { get; }
        public int Patch { get; }
        public string PreRelease { get; }

        private SemVer(int major, int minor, int patch, string pre)
        {
            Major = major;
            Minor = minor;
            Patch = patch;
            PreRelease = pre ?? string.Empty;
        }

        public static bool TryParse(string raw, out SemVer version)
        {
            version = null;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            var s = raw.Trim();
            if (s.StartsWith("v", StringComparison.OrdinalIgnoreCase)) s = s.Substring(1);

            string pre = "";
            var dash = s.IndexOf('-');
            if (dash >= 0)
            {
                pre = s.Substring(dash + 1);
                s = s.Substring(0, dash);
            }

            var parts = s.Split('.');
            if (parts.Length != 3) return false;

            if (!int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var maj)) return false;
            if (!int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var min)) return false;
            if (!int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var pat)) return false;

            version = new SemVer(maj, min, pat, pre);
            return true;
        }

        public int CompareTo(SemVer other)
        {
            if (other == null) return 1;
            var c = Major.CompareTo(other.Major); if (c != 0) return c;
            c = Minor.CompareTo(other.Minor);     if (c != 0) return c;
            c = Patch.CompareTo(other.Patch);     if (c != 0) return c;

            // Per SemVer 2.0.0, a version with prerelease is LOWER than the
            // same version without one.
            if (PreRelease.Length == 0 && other.PreRelease.Length == 0) return 0;
            if (PreRelease.Length == 0) return 1;
            if (other.PreRelease.Length == 0) return -1;

            return ComparePreRelease(PreRelease, other.PreRelease);
        }

        private static int ComparePreRelease(string a, string b)
        {
            var ap = a.Split('.');
            var bp = b.Split('.');
            var n = Math.Min(ap.Length, bp.Length);
            for (var i = 0; i < n; i++)
            {
                var ai = ap[i];
                var bi = bp[i];
                var aNum = int.TryParse(ai, NumberStyles.None, CultureInfo.InvariantCulture, out var an);
                var bNum = int.TryParse(bi, NumberStyles.None, CultureInfo.InvariantCulture, out var bn);
                if (aNum && bNum) { var c = an.CompareTo(bn); if (c != 0) return c; continue; }
                if (aNum) return -1;   // numeric identifiers sort lower than alphanumeric
                if (bNum) return 1;
                var sc = string.CompareOrdinal(ai, bi); if (sc != 0) return sc;
            }
            return ap.Length.CompareTo(bp.Length);
        }
    }
}
