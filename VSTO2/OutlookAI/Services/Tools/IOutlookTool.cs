using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// One Phase-2 Outlook tool. Implementations validate args (already
    /// JSON-schema validated by ToolDispatcher), perform the operation via
    /// IOutlookSurface, and return a JSON string the chat service inserts as
    /// a function_call_output. Errors should be returned as structured JSON
    /// ({"error":{"code","message"}}) rather than thrown, so the model can
    /// see and react to the failure.
    /// </summary>
    public interface IOutlookTool
    {
        string Name { get; }
        Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct);
    }
}
