namespace OutlookAI.Services.Export
{
    public interface IExportPathPolicy
    {
        void RequireInsideReportsDir(string path);
    }
}
