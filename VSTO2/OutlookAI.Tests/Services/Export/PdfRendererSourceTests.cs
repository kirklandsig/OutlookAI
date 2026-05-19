using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class PdfRendererSourceTests
    {
        [Fact]
        public void EnsureInitialized_UsesPdfSpecificWebView2DataFolder()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Export", "PdfRenderer.cs"));

            Assert.Contains("WebView2Bootstrap.PdfWebView2DataFolder", source);
            Assert.DoesNotContain("CoreWebView2Environment.CreateAsync(null, WebView2Bootstrap.WebView2DataFolder, null)", source);
        }

        [Fact]
        public void EnsureInitialized_LogsWebView2InitializationStages()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Export", "PdfRenderer.cs"));

            Assert.Contains("Initializing WebView2 dataFolder=", source);
            Assert.Contains(">> CoreWebView2Environment.CreateAsync", source);
            Assert.Contains("<< CoreWebView2Environment.CreateAsync", source);
            Assert.Contains(">> _webView.EnsureCoreWebView2Async", source);
            Assert.Contains("<< _webView.EnsureCoreWebView2Async", source);
        }

        private static string FindSourceFile(params string[] relativeParts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(relativeParts));
                if (File.Exists(candidate)) return candidate;

                current = current.Parent;
            }

            throw new FileNotFoundException("Could not find source file.", Path.Combine(relativeParts));
        }
    }
}
