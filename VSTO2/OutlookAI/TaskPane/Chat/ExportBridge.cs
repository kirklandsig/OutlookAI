using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.Chat
{
    public sealed class ExportBridge
    {
        private readonly IOutlookSurface _surface;
        private readonly IExportPathPolicy _pathPolicy;
        private readonly Func<string, Task> _runScript;
        private readonly Action<string> _openFileLauncher;
        private readonly Action<string> _revealInExplorerLauncher;

        public ExportBridge(
            IOutlookSurface surface,
            IExportPathPolicy pathPolicy,
            Func<string, Task> runScript,
            Action<string> openFileLauncher = null,
            Action<string> revealInExplorerLauncher = null)
        {
            _surface = surface ?? throw new ArgumentNullException(nameof(surface));
            _pathPolicy = pathPolicy ?? throw new ArgumentNullException(nameof(pathPolicy));
            _runScript = runScript ?? throw new ArgumentNullException(nameof(runScript));
            _openFileLauncher = openFileLauncher ?? OpenWithDefaultApp;
            _revealInExplorerLauncher = revealInExplorerLauncher ?? RevealWithExplorer;
        }

        public async Task<bool> HandleAsync(string type, JObject payload, CancellationToken ct)
        {
            switch (type)
            {
                case "export_pdf":
                    await ExportPdfAsync(payload, ct).ConfigureAwait(false);
                    return true;
                case "open_file":
                    await OpenFileAsync(payload).ConfigureAwait(false);
                    return true;
                case "reveal_in_explorer":
                    await RevealInExplorerAsync(payload).ConfigureAwait(false);
                    return true;
                default:
                    return false;
            }
        }

        private async Task ExportPdfAsync(JObject payload, CancellationToken ct)
        {
            var messageId = (string)payload?["message_id"] ?? "";
            try
            {
                var args = new ExportPdfArgs
                {
                    FilenameHint = (string)payload?["filename_hint"] ?? "",
                    ContentMarkdown = (string)payload?["content_markdown"] ?? "",
                };

                var saved = _surface.ExportPdf(args, ct);
                var fileInfo = new JObject(
                    new JProperty("path", saved?.Path ?? ""),
                    new JProperty("file_url", saved?.FileUrl ?? ""),
                    new JProperty("format", saved?.Format ?? "pdf"),
                    new JProperty("bytes", saved?.Bytes ?? 0L),
                    new JProperty("filename", saved?.Filename ?? ""));

                await _runScript("outlookai.onFileSaved(" + JsString(messageId) + ", " +
                    fileInfo.ToString(Formatting.None) + ");").ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                await PostErrorAsync(messageId, "cancelled", "Export cancelled.").ConfigureAwait(false);
            }
            catch (ExportException ex)
            {
                await PostErrorAsync(messageId, ex.Code, ex.Message).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                TraceLog.Write("ExportPdfAsync error: " + ex.Message, "ExportBridge");
                await PostErrorAsync(messageId, "pdf_render_failed", ex.Message).ConfigureAwait(false);
            }
        }

        private async Task OpenFileAsync(JObject payload)
        {
            var messageId = (string)payload?["message_id"] ?? "";
            var path = (string)payload?["path"] ?? "";
            try
            {
                _pathPolicy.RequireInsideReportsDir(path);
                _openFileLauncher(path);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OpenFileAsync error: " + ex.Message, "ExportBridge");
                await PostErrorAsync(messageId, "open_file_failed", ex.Message).ConfigureAwait(false);
            }
        }

        private async Task RevealInExplorerAsync(JObject payload)
        {
            var messageId = (string)payload?["message_id"] ?? "";
            var path = (string)payload?["path"] ?? "";
            try
            {
                _pathPolicy.RequireInsideReportsDir(path);
                _revealInExplorerLauncher(path);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RevealInExplorerAsync error: " + ex.Message, "ExportBridge");
                await PostErrorAsync(messageId, "reveal_in_explorer_failed", ex.Message).ConfigureAwait(false);
            }
        }

        private Task PostErrorAsync(string messageId, string code, string detail)
        {
            var error = new JObject(
                new JProperty("error", code ?? "export_failed"),
                new JProperty("code", code ?? "export_failed"),
                new JProperty("detail", detail ?? ""));

            return _runScript("outlookai.onExportError(" + JsString(messageId) + ", " +
                error.ToString(Formatting.None) + ");");
        }

        private static string JsString(string value)
            => JsonConvert.SerializeObject(value ?? "");

        private static void OpenWithDefaultApp(string path)
        {
            Process.Start(new ProcessStartInfo(path)
            {
                UseShellExecute = true,
                // Force a local working directory so ShellExecute does not try
                // to set CWD to a UNC path when ~\Documents is folder-redirected
                // to a network share. Target file path is unaffected.
                WorkingDirectory = Path.GetTempPath(),
            });
        }

        private static void RevealWithExplorer(string path)
        {
            Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + path.Replace("\"", "\\\"") + "\"")
            {
                UseShellExecute = true,
                WorkingDirectory = Path.GetTempPath(),
            });
        }
    }
}
