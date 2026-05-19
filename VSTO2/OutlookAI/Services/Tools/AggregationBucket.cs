namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// One bucket in an aggregate-messages result. Label is the grouped
    /// key (sender name, "yyyy-MM-dd" date, or folder name depending on
    /// args.GroupBy). Count is the number of matching messages.
    /// </summary>
    public sealed class AggregationBucket
    {
        public string Label { get; set; }
        public int Count { get; set; }
    }
}
