using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportPathPolicyTests
    {
        private static (ExportPathPolicy policy, string baseDir) CreateInSandbox()
        {
            var baseDir = Path.Combine(Path.GetTempPath(), "OutlookAITest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(baseDir);
            return (new ExportPathPolicy(new ExportPathResolver(baseDirOverride: baseDir)), baseDir);
        }

        [Fact]
        public void AcceptsCanonicalPathInsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var ok = Path.Combine(baseDir, "Report.xlsx");
                File.WriteAllText(ok, "");

                policy.RequireInsideReportsDir(ok);
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void AcceptsPathCaseInsensitive()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var weirdCase = baseDir.ToUpperInvariant() + Path.DirectorySeparatorChar + "Report.xlsx";
                File.WriteAllText(Path.Combine(baseDir, "Report.xlsx"), "");

                policy.RequireInsideReportsDir(weirdCase);
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsTraversalAttempt()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                var evil = Path.Combine(baseDir, "..", "..", "Windows", "System32", "cmd.exe");

                Assert.Throws<UnauthorizedExportPathException>(() => policy.RequireInsideReportsDir(evil));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsPathOutsideReportsDir()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"C:\Windows\System32\cmd.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsUncPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(@"\\server\share\evil.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsNullPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(null));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsEmptyPath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir(""));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }

        [Fact]
        public void RejectsRelativePath()
        {
            var (policy, baseDir) = CreateInSandbox();
            try
            {
                Assert.Throws<UnauthorizedExportPathException>(
                    () => policy.RequireInsideReportsDir("evil.exe"));
            }
            finally
            {
                if (Directory.Exists(baseDir)) Directory.Delete(baseDir, true);
            }
        }
    }
}
