using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateManifestTests
    {
        [Fact]
        public void LoadFromFile_ValidJson_ReturnsManifest()
        {
            var path = Path.GetTempFileName();
            try
            {
                File.WriteAllText(path,
                    "{\"tag\":\"v2.1.0\",\"commit\":\"abc1234\"," +
                    "\"build_date\":\"2026-06-02T19:14:00Z\"," +
                    "\"repo\":\"kirklandsig/OutlookAI\"}");
                var manifest = UpdateManifest.LoadFromFile(path);
                Assert.Equal("v2.1.0", manifest.Tag);
                Assert.Equal("abc1234", manifest.Commit);
                Assert.Equal("kirklandsig/OutlookAI", manifest.Repo);
                Assert.False(manifest.IsDevBuild);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void LoadFromFile_MissingFile_ReturnsDevSentinel()
        {
            var manifest = UpdateManifest.LoadFromFile(@"C:\definitely\does\not\exist\version.json");
            Assert.True(manifest.IsDevBuild);
            Assert.Equal(UpdateManifest.DevSentinel, manifest.Tag);
        }

        [Fact]
        public void LoadFromFile_MalformedJson_ReturnsDevSentinel()
        {
            var path = Path.GetTempFileName();
            try
            {
                File.WriteAllText(path, "this is not json");
                var manifest = UpdateManifest.LoadFromFile(path);
                Assert.True(manifest.IsDevBuild);
            }
            finally { File.Delete(path); }
        }

        [Fact]
        public void LoadFromZip_ReadsVersionJsonAtArchiveRoot()
        {
            var zipPath = Path.GetTempFileName();
            File.Delete(zipPath);
            try
            {
                using (var fs = File.Create(zipPath))
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    var entry = archive.CreateEntry("version.json");
                    using (var w = new StreamWriter(entry.Open(), Encoding.UTF8))
                    {
                        w.Write("{\"tag\":\"v2.2.0\",\"commit\":\"def5678\"," +
                                "\"build_date\":\"2026-07-01T00:00:00Z\"," +
                                "\"repo\":\"kirklandsig/OutlookAI\"}");
                    }
                }
                var manifest = UpdateManifest.LoadFromZip(zipPath);
                Assert.Equal("v2.2.0", manifest.Tag);
            }
            finally { if (File.Exists(zipPath)) File.Delete(zipPath); }
        }

        [Fact]
        public void LoadFromZip_NoVersionJson_ReturnsDevSentinel()
        {
            var zipPath = Path.GetTempFileName();
            File.Delete(zipPath);
            try
            {
                using (var fs = File.Create(zipPath))
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    archive.CreateEntry("something-else.txt");
                }
                var manifest = UpdateManifest.LoadFromZip(zipPath);
                Assert.True(manifest.IsDevBuild);
            }
            finally { if (File.Exists(zipPath)) File.Delete(zipPath); }
        }
    }
}
