using System;
using System.IO;
using System.Threading;

namespace OutlookAI.Services
{
    /// <summary>
    /// Cross-process lock for read-modify-write of a shared file (e.g. by
    /// several RDS sessions): an exclusive, delete-on-close lock file, released
    /// even if the holder's process dies, so it can never go stale.
    /// </summary>
    internal static class FileLock
    {
        /// <summary>Waits up to <paramref name="timeout"/>, then throws <see cref="TimeoutException"/>.</summary>
        public static FileStream Acquire(string lockPath, TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow + timeout;
            while (true)
            {
                try
                {
                    return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    // Held by another writer, or its delete still pending (which
                    // surfaces as access denied). A missing file that we may not
                    // create is a real permission problem: report it.
                    if (ex is UnauthorizedAccessException && !File.Exists(lockPath)) throw;
                    if (DateTime.UtcNow >= deadline) throw new TimeoutException("Timed out waiting for " + lockPath, ex);
                    Thread.Sleep(50);
                }
            }
        }
    }
}
