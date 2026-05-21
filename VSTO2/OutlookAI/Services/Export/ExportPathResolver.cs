using System;
using System.IO;

namespace OutlookAI.Services.Export
{
    public sealed class ExportPathResolver
    {
        private readonly string _baseDir;

        public ExportPathResolver()
            : this(baseDirOverride: null)
        {
        }

        public ExportPathResolver(string baseDirOverride)
            : this(baseDirOverride, docsProvider: null, localAppDataProvider: null)
        {
        }

        // Test seam: lets unit tests inject fake MyDocuments / LocalAppData
        // values so we can exercise the UNC-fallback branch without
        // actually owning a redirected folder.
        internal ExportPathResolver(
            string baseDirOverride,
            Func<string> docsProvider,
            Func<string> localAppDataProvider)
        {
            if (baseDirOverride != null)
            {
                _baseDir = Path.GetFullPath(baseDirOverride);
                return;
            }

            var docs = (docsProvider != null
                ? docsProvider()
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            var candidate = Path.Combine(docs, "OutlookAI", "Reports");

            // On RDS hosts with Folder Redirection, MyDocuments can resolve
            // to a UNC share (e.g. \\fileserver\users\jdoe\Documents). The
            // export pipeline downstream (ExportPathPolicy + PdfRenderer +
            // ClosedXML tempfile-rename) doesn't tolerate UNC roots, so
            // fall back to LocalAppData which is always machine-local and
            // writable by the running user.
            if (candidate.StartsWith(@"\\", StringComparison.Ordinal))
            {
                var localAppData = (localAppDataProvider != null
                    ? localAppDataProvider()
                    : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
                candidate = Path.Combine(localAppData, "OutlookAI", "Reports");
            }

            _baseDir = candidate;
        }

        public string ResolveBaseDir()
        {
            return _baseDir;
        }

        public void EnsureExists()
        {
            if (File.Exists(_baseDir))
            {
                throw new IOException("Reports path '" + _baseDir + "' exists as a file, not a directory.");
            }

            Directory.CreateDirectory(_baseDir);
        }
    }
}
