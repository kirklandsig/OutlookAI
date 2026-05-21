using System;
using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateInstallerSourceTests
    {
        private static string FindSourceFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate)) return candidate;
                current = current.Parent;
            }
            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }

        [Fact]
        public void LaunchElevatedInstall_UsesRunasVerb_WithNoExitAndExecutionPolicyBypass()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Updates", "UpdateInstaller.cs"));
            var methodStart = source.IndexOf("public LaunchResult LaunchElevatedInstall", StringComparison.Ordinal);
            Assert.True(methodStart >= 0, "LaunchElevatedInstall should exist.");
            var method = source.Substring(methodStart);

            Assert.Contains("Verb = \"runas\"", method);
            Assert.Contains("UseShellExecute = true", method);
            Assert.Contains("-NoExit", method);
            Assert.Contains("-NoProfile", method);
            Assert.Contains("-ExecutionPolicy Bypass", method);
            Assert.Contains("-File", method);
            Assert.Contains("-SourcePath", method);
            Assert.Contains("WorkingDirectory = Path.GetTempPath()", method);
        }

        [Fact]
        public void LaunchElevatedInstall_HandlesUacDeclinedExitCode1223()
        {
            var source = File.ReadAllText(FindSourceFile("OutlookAI", "Services", "Updates", "UpdateInstaller.cs"));
            Assert.Contains("NativeErrorCode == 1223", source);
            Assert.Contains("UacDeclined", source);
        }
    }
}
