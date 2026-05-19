namespace OutlookAI.Services.Variants
{
    /// <summary>
    /// One drafting variant produced by the Variants tab. The model emits
    /// these as JSON inside a fenced code block; <see cref="VariantParser"/>
    /// shapes them into this POCO for the UI.
    /// </summary>
    public sealed class Variant
    {
        public Tone Tone { get; set; }
        public string Rationale { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
    }
}
