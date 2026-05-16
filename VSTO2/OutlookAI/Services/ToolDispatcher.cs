using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services
{
    /// <summary>
    /// Routes the model's function_call to the matching <see cref="IOutlookTool"/>
    /// instance, validates the arguments JSON shape, and wraps any failure as a
    /// structured <c>{"error": {...}}</c> JSON response so the model can see and
    /// recover from it.
    /// </summary>
    public sealed class ToolDispatcher
    {
        private readonly Dictionary<string, IOutlookTool> _tools;
        private readonly IOutlookSurface _surface;

        public ToolDispatcher(IEnumerable<IOutlookTool> tools, IOutlookSurface surface)
        {
            _tools = new Dictionary<string, IOutlookTool>(StringComparer.Ordinal);
            foreach (var t in tools)
            {
                if (t == null) continue;
                _tools[t.Name] = t;
            }
            _surface = surface;
        }

        public async Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
        {
            if (!_tools.TryGetValue(toolName ?? string.Empty, out var tool))
            {
                return BuildError("unknown_tool", "Tool '" + toolName + "' is not registered.");
            }

            JObject parsed;
            try
            {
                parsed = string.IsNullOrWhiteSpace(argsJson) ? new JObject() : JObject.Parse(argsJson);
            }
            catch (Exception ex)
            {
                return BuildError("invalid_arguments", ex.Message);
            }

            try
            {
                return await tool.ExecuteAsync(
                    parsed.ToString(Newtonsoft.Json.Formatting.None),
                    _surface,
                    ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                return BuildError(ex.GetType().Name, ex.Message);
            }
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
