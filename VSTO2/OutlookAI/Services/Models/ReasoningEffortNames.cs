using System;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// Spelling of reasoning efforts. The Codex backend and model catalog use
    /// lowercase wire values ("xhigh"); dropdowns and config.xml use the display
    /// spelling ("XHigh"). <see cref="None"/> is the app's own option meaning
    /// "omit the reasoning field" (the server then applies its default).
    /// </summary>
    public static class ReasoningEffortNames
    {
        public const string None = "None";

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

        public static bool IsNone(string effort)
        {
            return string.IsNullOrWhiteSpace(effort)
                || string.Equals(effort.Trim(), None, StringComparison.OrdinalIgnoreCase);
        }
    }
}
