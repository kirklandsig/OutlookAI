using System;

namespace OutlookAI.Services.Variants
{
    /// <summary>
    /// Fixed nine-value tone enum (per Phase 2 spec). The model is asked to
    /// pick one per variant; <see cref="ToneExtensions.ClosestTo"/> clamps any
    /// free-form string the model emits into one of these values.
    /// </summary>
    public enum Tone
    {
        Formal,
        Brief,
        Persuasive,
        Friendly,
        Technical,
        Apologetic,
        Direct,
        Diplomatic,
        Enthusiastic
    }

    public static class ToneExtensions
    {
        /// <summary>
        /// Maps a free-form tone string (whatever the model returned in the
        /// JSON envelope) to the closest enum value. Order:
        ///   1) Case-insensitive exact match on enum name.
        ///   2) Small synonym map for common variants.
        ///   3) Fall back to <see cref="Tone.Direct"/>.
        /// </summary>
        public static Tone ClosestTo(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return Tone.Direct;
            foreach (Tone t in Enum.GetValues(typeof(Tone)))
            {
                if (string.Equals(t.ToString(), raw, StringComparison.OrdinalIgnoreCase))
                {
                    return t;
                }
            }
            var r = raw.ToLowerInvariant();
            if (r.Contains("polite") || r.Contains("warm")) return Tone.Friendly;
            if (r.Contains("short") || r.Contains("conc")) return Tone.Brief;
            if (r.Contains("sales") || r.Contains("convince")) return Tone.Persuasive;
            if (r.Contains("sorry") || r.Contains("apolog")) return Tone.Apologetic;
            if (r.Contains("excit") || r.Contains("eager")) return Tone.Enthusiastic;
            if (r.Contains("tech") || r.Contains("engineer")) return Tone.Technical;
            if (r.Contains("formal") || r.Contains("official")) return Tone.Formal;
            if (r.Contains("diplo") || r.Contains("tact")) return Tone.Diplomatic;
            return Tone.Direct;
        }
    }
}
