namespace OutlookAI.Services.Tools
{
    public sealed class ExportPdfArgs
    {
        public string FilenameHint { get; set; }
        public string ContentMarkdown { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
    }
}
