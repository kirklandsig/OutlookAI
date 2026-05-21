using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportPathResolverTests
    {
        [Fact]
        public void ResolveBaseDir_ReturnsExpectedBranchBasedOnDocsRedirection()
        {
            // Verifies the production code-path against whatever the running
            // host happens to have. On a normal dev box MyDocuments is local and
            // we get the Documents branch; on an RDS host with Folder Redirection
            // we get the LocalAppData fallback. Either is correct.
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var resolver = new ExportPathResolver();

            var path = resolver.ResolveBaseDir();

            var expectedDocs = Path.Combine(docs, "OutlookAI", "Reports");
            var expectedLocal = Path.Combine(localAppData, "OutlookAI", "Reports");
            Assert.True(
                path == expectedDocs || path == expectedLocal,
                $"Expected '{expectedDocs}' or '{expectedLocal}' but got '{path}'.");
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

        [Fact]
        public void ResolveBaseDir_WhenMyDocumentsIsUnc_FallsBackToLocalAppData()
        {
            var fakeUnc = @"\\fileserver\users\jdoe\Documents";
            var fakeLocal = @"C:\Users\jdoe\AppData\Local";
            var resolver = new ExportPathResolver(
                baseDirOverride: null,
                docsProvider: () => fakeUnc,
                localAppDataProvider: () => fakeLocal);

            var path = resolver.ResolveBaseDir();

            Assert.Equal(Path.Combine(fakeLocal, "OutlookAI", "Reports"), path);
        }

        [Fact]
        public void ResolveBaseDir_WhenMyDocumentsIsLocal_UsesIt()
        {
            var fakeLocal = @"C:\Users\jdoe\Documents";
            var resolver = new ExportPathResolver(
                baseDirOverride: null,
                docsProvider: () => fakeLocal,
                localAppDataProvider: () => @"C:\Users\jdoe\AppData\Local");

            var path = resolver.ResolveBaseDir();

            Assert.Equal(Path.Combine(fakeLocal, "OutlookAI", "Reports"), path);
        }
    }
}
