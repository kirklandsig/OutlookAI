namespace OutlookAI.Services.QuickTools
{
    /// <summary>One page in the review window: files on disk plus the user's choices.</summary>
    internal sealed class ReviewPage
    {
        public ReviewPage(int number, string label)
        {
            Number = number;
            Label = label ?? "";
            Plan = PagePlan.Fallback();
        }

        public int Number { get; }

        /// <summary>The attachment's file name, plus "(page n)" for a multi-page file.</summary>
        public string Label { get; }

        public bool Readable { get; set; }
        public bool OverLimit { get; set; }
        public string UprightPath { get; set; }
        public string CleanedPath { get; set; }
        public PagePlan Plan { get; set; }
        public bool Include { get; set; }
        public int QuarterTurns { get; set; }
        public bool UseOriginal { get; set; }
        public PageColor? ColorOverride { get; set; }

        public PageColor EffectiveColor => ColorOverride ?? Plan.Color;

        public string Status
        {
            get
            {
                if (OverLimit) return "not processed (limit " + ImagesToPdfPipeline.MaxPages + ")";
                if (!Readable) return "can't read this file";
                if (!Plan.Analyzed) return "AI couldn't analyze this one";
                if (Plan.Kind == PageKind.NotDocument)
                    return string.IsNullOrEmpty(Plan.Reason) ? "not a document" : "not a document: " + Plan.Reason;
                return "page";
            }
        }
    }
}
