using System.IO;
using OutlookAI.TaskPane.Chat;
using Xunit;

namespace OutlookAI.Tests.TaskPane.Chat
{
    public class WebView2BootstrapTests
    {
        // Resource names use "." as both folder separator and extension
        // boundary. The helper has to disambiguate: last dot = extension,
        // earlier dots = folder separators, and ".min.js" is treated as a
        // single suffix on the filename (not a fake folder).

        [Fact]
        public void ResourceNameToRelativePath_SimpleFile()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("index.html");
            Assert.Equal("index.html", result);
        }

        [Fact]
        public void ResourceNameToRelativePath_SingleFolderPrefix()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.marked.js");
            Assert.Equal(Path.Combine("vendor", "marked.js"), result);
        }

        [Fact]
        public void ResourceNameToRelativePath_MinifiedFile_KeepsMinOnFilename()
        {
            // "vendor.marked.min.js" must become vendor\marked.min.js, NOT
            // vendor\marked\min.js.
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.marked.min.js");
            Assert.Equal(Path.Combine("vendor", "marked.min.js"), result);
        }

        [Fact]
        public void ResourceNameToRelativePath_NestedFolder()
        {
            var result = WebView2Bootstrap.ResourceNameToRelativePath("vendor.highlight.styles.github.css");
            Assert.Equal(
                Path.Combine("vendor", "highlight", "styles", "github.css"),
                result);
        }

        [Fact]
        public void ResourceNameToRelativePath_NoExtension_ReturnsUnchanged()
        {
            // Unusual case: no dot at all.
            var result = WebView2Bootstrap.ResourceNameToRelativePath("README");
            Assert.Equal("README", result);
        }

        [Fact]
        public void VirtualHost_IsExpectedValue()
        {
            // The virtual host name is part of the public contract - the
            // chat HTML/JS will resolve resources via this URL.
            Assert.Equal("outlookai.local", WebView2Bootstrap.VirtualHost);
        }
    }
}
