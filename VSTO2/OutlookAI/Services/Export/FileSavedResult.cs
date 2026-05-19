namespace OutlookAI.Services.Export
{
    public sealed class FileSavedResult
    {
        public string Path { get; set; }
        public string FileUrl { get; set; }
        public string Format { get; set; }
        public long Bytes { get; set; }
        public string Filename { get; set; }
    }
}
