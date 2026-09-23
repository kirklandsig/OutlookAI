using System.IO;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>Finds a repository file by walking up from the test output folder.</summary>
    internal static class RepoFiles
    {
        public static string Find(params string[] parts)
        {
            var relative = Path.Combine(parts);
            for (var dir = new DirectoryInfo(Directory.GetCurrentDirectory()); dir != null; dir = dir.Parent)
            {
                var candidate = Path.Combine(dir.FullName, relative);
                if (File.Exists(candidate)) return candidate;
            }
            throw new FileNotFoundException("Could not find " + relative + " above " + Directory.GetCurrentDirectory());
        }
    }
}
