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
        {
            if (baseDirOverride != null)
            {
                _baseDir = Path.GetFullPath(baseDirOverride);
                return;
            }

            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _baseDir = Path.Combine(docs, "OutlookAI", "Reports");
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
