using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public sealed class OutlookExportExcelTool : IOutlookTool
    {
        public string Name => "outlook_export_excel";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = ExportExcelArgsParser.Parse(argsJson);
                var saved = surface.ExportExcel(args, ct);

                var json = new JObject(
                    new JProperty("result_type", "file_saved"),
                    new JProperty("path", saved.Path ?? ""),
                    new JProperty("file_url", saved.FileUrl ?? ""),
                    new JProperty("format", saved.Format ?? ""),
                    new JProperty("bytes", saved.Bytes),
                    new JProperty("filename", saved.Filename ?? ""));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Export cancelled by user."));
            }
            catch (ToolArgValidationException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (ArgumentException ex)
            {
                return Task.FromResult(BuildError("invalid_args", ex.Message));
            }
            catch (ExportException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(BuildError("excel_build_failed", ex.Message));
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
