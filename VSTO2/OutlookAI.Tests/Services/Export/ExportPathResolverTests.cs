using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportPathResolverTests
    {
        [Fact]
        public void ResolveBaseDir_ReturnsDocumentsOutlookAIReports()
        {
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var resolver = new ExportPathResolver();

            var path = resolver.ResolveBaseDir();

            Assert.Equal(Path.Combine(docs, "OutlookAI", "Reports"), path);
        }

        [Fact]
        public void EnsureExists_CreatesDirectoryIfMissing()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            try
            {
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);

                Assert.False(Directory.Exists(sandbox));
                resolver.EnsureExists();

                Assert.True(Directory.Exists(sandbox));
            }
            finally
            {
                if (Directory.Exists(sandbox)) Directory.Delete(sandbox, true);
            }
        }

        [Fact]
        public void EnsureExists_IsIdempotent()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(sandbox);
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);

                resolver.EnsureExists();
                resolver.EnsureExists();

                Assert.True(Directory.Exists(sandbox));
            }
            finally
            {
                if (Directory.Exists(sandbox)) Directory.Delete(sandbox, true);
            }
        }

        [Fact]
        public void EnsureExists_ThrowsWhenPathIsAFile()
        {
            var sandbox = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            File.WriteAllText(sandbox, "not a directory");
            try
            {
                var resolver = new ExportPathResolver(baseDirOverride: sandbox);

                Assert.Throws<IOException>(() => resolver.EnsureExists());
            }
            finally
            {
                if (File.Exists(sandbox)) File.Delete(sandbox);
            }
        }
    }
}
