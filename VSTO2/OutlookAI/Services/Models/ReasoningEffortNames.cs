using System;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// Spelling of reasoning efforts. The Codex backend and model catalog use
    /// lowercase wire values ("xhigh"); dropdowns and config.xml use the display
    /// spelling ("XHigh"). <see cref="Auto"/> is the app's own option meaning
    /// "omit the reasoning field": the model then uses its default (medium on
    /// current models).
    /// </summary>
    public static class ReasoningEffortNames
    {
        public const string Auto = "Auto";

        /// <summary>
        /// What <see cref="Auto"/> was called before v2.2.1. It read like "no
        /// reasoning" but always meant "omit the field". It is accepted as a
        /// synonym, and config.xml still stores Auto under this name
        /// (<see cref="ToConfigValue"/>).
        /// </summary>
        public const string LegacyNone = "None";

        public static string ToDisplay(string wire)
        {
            if (string.IsNullOrEmpty(wire)) return wire;
            switch (wire)
            {
                case "none": return "None";
                case "minimal": return "Minimal";
                case "low": return "Low";
                case "medium": return "Medium";
                case "high": return "High";
                case "xhigh": return "XHigh";
                case "max": return "Max";
                default: return char.ToUpperInvariant(wire[0]) + wire.Substring(1);
            }
        }

        /// <summary>True for "omit the reasoning field": blank, Auto or the legacy None.</summary>
        public static bool IsAuto(string effort)
        {
            if (string.IsNullOrWhiteSpace(effort)) return true;
            var trimmed = effort.Trim();
            return string.Equals(trimmed, Auto, StringComparison.OrdinalIgnoreCase)
                || string.Equals(trimmed, LegacyNone, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The spelling config.xml stores. Auto is saved as "None", which every
        /// OutlookAI version reads as "omit the field"; v2.2.0 and older would
        /// ignore "Auto", e.g. in a roaming profile on a server not yet updated.
        /// </summary>
        public static string ToConfigValue(string effort)
        {
            return IsAuto(effort) ? LegacyNone : effort;
        }
    }
}
