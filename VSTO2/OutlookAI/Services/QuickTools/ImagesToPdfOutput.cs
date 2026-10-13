using System;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    internal static class ImagesToPdfOutput
    {
        private const int ErrorSharingViolation = 0x20;
        private const int ErrorLockViolation = 0x21;

        /// <summary>
        /// Copies the finished PDF (written locally first) next to where the user
        /// chose, then swaps it in, so a copy that fails partway (a dropped share)
        /// never touches an existing file. Null when saved, otherwise a message
        /// for the user.
        /// </summary>
        public static string Publish(string tempPdf, string destination)
        {
            var id = Guid.NewGuid().ToString("N").Substring(0, 8);
            var partial = destination + "." + id + ".partial";
            var backup = destination + "." + id + ".bak";
            try
            {
                File.Copy(tempPdf, partial, overwrite: false);
                if (File.Exists(destination))
                {
                    return Replace(partial, destination, backup);
                }
                File.Move(partial, destination);
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return "You don't have permission to save in that folder. Choose another folder.";
            }
            catch (IOException ex)
            {
                return "The PDF couldn't be saved there. If the file is open in another program, close it and try again.\n\n" + ex.Message;
            }
            finally
            {
                TryDelete(partial);
            }
        }

        // The old file always survives as the destination or the backup.
        private static string Replace(string partial, string destination, string backup)
        {
            try
            {
                // Ignoring metadata errors: shares that withhold WRITE_DAC from the
                // owner can't take the old file's ACL, and the swap must still work.
                File.Replace(partial, destination, backup, ignoreMetadataErrors: true);
                TryDelete(backup);
                return null;
            }
            catch (IOException ex) when (IsSharingViolation(ex))
            {
                throw;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Fall through to the rename swap.
            }

            if (!File.Exists(destination) && File.Exists(backup))
            {
                File.Move(backup, destination);
            }
            TryDelete(backup);

            File.Move(destination, backup);
            try
            {
                File.Move(partial, destination);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                try
                {
                    File.Move(backup, destination);
                }
                catch (Exception restore) when (restore is IOException || restore is UnauthorizedAccessException)
                {
                    return "The PDF couldn't be saved there, and the previous file was left at:\n" + backup;
                }
                throw;
            }
            TryDelete(backup);
            return null;
        }

        private static bool IsSharingViolation(IOException ex)
        {
            var code = ex.HResult & 0xFFFF;
            return code == ErrorSharingViolation || code == ErrorLockViolation;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }
}
