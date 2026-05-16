using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services
{
    /// <summary>
    /// IToolHost implementation that aggregates the 10 Phase 2
    /// <see cref="IOutlookTool"/>s and a single <see cref="IOutlookSurface"/>,
    /// dispatching via <see cref="ToolDispatcher"/>. Write tools are included
    /// conditionally based on the admin's tool-permission setting.
    /// </summary>
    public sealed class OutlookToolHost : IToolHost, IDisposable
    {
        private readonly ToolDispatcher _dispatcher;

        public OutlookToolHost(IOutlookSurface surface, bool includeWriteTools)
        {
            var tools = new List<IOutlookTool>
            {
                new OutlookGetCurrentComposeStateTool(),
                new OutlookListFoldersTool(),
                new OutlookSearchMessagesTool(),
                new OutlookReadMessageTool(),
                new OutlookCountMessagesTool(),
                new OutlookListRecentThreadsWithTool(),
            };
            if (includeWriteTools)
            {
                // Per-tool gate against Config.EnabledWriteTools. Mirrors the
                // gating in ToolCatalogSchema so the dispatcher never has a
                // tool registered that the catalog didn't advertise.
                var enabled = Config.EnabledWriteTools
                    ?? new HashSet<string>(Config.AllWriteTools);

                if (enabled.Contains("outlook_create_draft"))
                    tools.Add(new OutlookCreateDraftTool());
                if (enabled.Contains("outlook_mark_as_read"))
                    tools.Add(new OutlookMarkAsReadTool());
                if (enabled.Contains("outlook_flag_message"))
                    tools.Add(new OutlookFlagMessageTool());
                if (enabled.Contains("outlook_set_category"))
                    tools.Add(new OutlookSetCategoryTool());
            }
            _dispatcher = new ToolDispatcher(tools, surface);
        }

        public Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
            => _dispatcher.DispatchAsync(toolName, argsJson, ct);

        public void Dispose() { /* surface lifecycle is owned by ThisAddIn / pane */ }
    }
}
