using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_list_folders. Returns folder tree (max depth 6, max 200 nodes).
    /// </summary>
    public sealed class OutlookListFoldersTool : IOutlookTool
    {
        public string Name => "outlook_list_folders";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var folders = surface.ListFolders() ?? new FolderResult[0];
            var json = new JObject(
                new JProperty("folders", new JArray(folders.Select(f =>
                    new JObject(
                        new JProperty("id", f.Id ?? ""),
                        new JProperty("name", f.Name ?? ""),
                        new JProperty("parent_id", f.ParentId),
                        new JProperty("item_count", f.ItemCount))))));
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }
    }
}
