using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OutlookAI.Diagnostics;

namespace OutlookAI.TaskPane.Chat
{
    /// <summary>
    /// One-shot bootstrap helper for the WebView2 control that hosts the
    /// Phase 2 chat surface. Responsibilities:
    /// <list type="bullet">
    ///   <item>Resolve / create the per-machine WebView2 data folder under
    ///         <c>%LOCALAPPDATA%\OutlookAI\WebView2Data</c> so multiple Outlook
    ///         processes don't fight over the same singleton runtime instance.</item>
    ///   <item>Materialize the bundled <c>WebUI/</c> assets (HTML, CSS, JS,
    ///         vendored libs) into <c>%LOCALAPPDATA%\OutlookAI\WebUI</c>.</item>
    ///   <item>Map the <c>outlookai.local</c> virtual host to that folder so
    ///         relative <c>fetch()</c>s and <c>&lt;script src&gt;</c>s resolve
    ///         without disk-path leakage.</item>
    ///   <item>Surface <see cref="WebView2Initialized"/> /
    ///         <see cref="WebView2InitializationFailed"/> events the
    ///         <c>ChatController</c> awaits before pushing the first message.</item>
    /// </list>
    /// </summary>
    public static class WebView2Bootstrap
    {
        public const string VirtualHost = "outlookai.local";

        private static string LocalAppDataRoot =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI");

        public static string WebUiFolder => Path.Combine(LocalAppDataRoot, "WebUI");
        public static string WebView2DataFolder => Path.Combine(LocalAppDataRoot, "WebView2Data");
        public static string PdfWebView2DataFolder => Path.Combine(LocalAppDataRoot, "WebView2PdfData");

        /// <summary>
        /// Initialize a WebView2 control, extract embedded WebUI resources,
        /// and wire up the virtual host mapping. Caller must add the control
        /// to its parent before calling this. Throws on failure - caller
        /// is expected to fall back to a friendly "install WebView2" panel.
        /// </summary>
        public static async Task InitializeAsync(WebView2 host)
        {
            if (host == null) throw new ArgumentNullException(nameof(host));

            TraceLog.Write("Bootstrap.InitializeAsync entered", "WebView2Bootstrap");
            EnsureFolders();
            TraceLog.Write("Folders ensured", "WebView2Bootstrap");
            ExtractEmbeddedWebUi();
            TraceLog.Write("Embedded WebUI extracted", "WebView2Bootstrap");

            TraceLog.Write(">> CoreWebView2Environment.CreateAsync", "WebView2Bootstrap");
            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: WebView2DataFolder,
                options: null).ConfigureAwait(true);
            TraceLog.Write("<< CreateAsync returned", "WebView2Bootstrap");

            TraceLog.Write(">> host.EnsureCoreWebView2Async", "WebView2Bootstrap");
            await host.EnsureCoreWebView2Async(env).ConfigureAwait(true);
            TraceLog.Write("<< EnsureCoreWebView2Async returned", "WebView2Bootstrap");

            host.CoreWebView2.Settings.AreDevToolsEnabled = false;
            host.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            host.CoreWebView2.Settings.IsStatusBarEnabled = false;
            host.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
            host.CoreWebView2.Settings.IsZoomControlEnabled = false;
            TraceLog.Write("Settings locked down", "WebView2Bootstrap");

            host.CoreWebView2.SetVirtualHostNameToFolderMapping(
                VirtualHost,
                WebUiFolder,
                CoreWebView2HostResourceAccessKind.Allow);
            TraceLog.Write("Virtual host mapping set", "WebView2Bootstrap");
        }

        /// <summary>Returns true if the Evergreen WebView2 runtime is installed on this machine.</summary>
        public static bool IsRuntimeInstalled()
        {
            try
            {
                var version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(version);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void EnsureFolders()
        {
            Directory.CreateDirectory(WebUiFolder);
            Directory.CreateDirectory(WebView2DataFolder);
        }

        /// <summary>
        /// Copies every <c>WebUI\*</c> embedded resource from the OutlookAI
        /// assembly into <see cref="WebUiFolder"/>. Replaces any existing
        /// files so an upgrade always lands the latest UI. Resources are
        /// embedded with logical names like
        /// <c>OutlookAI.WebUI.index.html</c>.
        /// </summary>
        private static void ExtractEmbeddedWebUi()
        {
            var asm = Assembly.GetExecutingAssembly();
            const string prefix = "OutlookAI.WebUI.";
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                // Logical name -> relative file path. "." in resource names is
                // ambiguous (folder separator vs filename "."), so we only
                // treat the LAST "." before the extension as the file
                // extension boundary, and earlier "." as folder separators.
                var relative = ResourceNameToRelativePath(name.Substring(prefix.Length));
                var targetPath = Path.Combine(WebUiFolder, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(targetPath));

                using (var src = asm.GetManifestResourceStream(name))
                using (var dst = File.Create(targetPath))
                {
                    src.CopyTo(dst);
                }
            }
        }

        /// <summary>
        /// Convert a resource name like <c>styles.css</c> or
        /// <c>vendor.marked.min.js</c> into a relative file path. The last
        /// dot is treated as the extension separator; all earlier dots are
        /// folder separators. So <c>vendor.marked.min.js</c> becomes
        /// <c>vendor\marked.min.js</c>.
        /// </summary>
        public static string ResourceNameToRelativePath(string resourceTail)
        {
            if (string.IsNullOrEmpty(resourceTail)) return resourceTail;
            var lastDot = resourceTail.LastIndexOf('.');
            if (lastDot <= 0)
            {
                return resourceTail;
            }
            var stem = resourceTail.Substring(0, lastDot);
            var ext = resourceTail.Substring(lastDot); // includes the dot
            // Special case: if the stem has a hyphen-separated minified suffix
            // like ".min", we want to keep it on the filename, not on a
            // folder. So treat the LAST two dots as one logical extension.
            if (stem.EndsWith(".min", StringComparison.Ordinal))
            {
                var lastDot2 = stem.LastIndexOf('.');
                if (lastDot2 > 0)
                {
                    var stem2 = stem.Substring(0, lastDot2);
                    var minSegment = stem.Substring(lastDot2); // ".min"
                    return stem2.Replace('.', Path.DirectorySeparatorChar) + minSegment + ext;
                }
            }
            return stem.Replace('.', Path.DirectorySeparatorChar) + ext;
        }
    }
}
