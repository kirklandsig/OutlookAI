using System;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    internal static class ImagesToPdfOutput
    {
        /// <summary>
        /// Copies the finished PDF (written locally first) next to where the user
        /// chose, then swaps it in, so a copy that fails partway (a dropped share)
        /// never touches an existing file. Null when saved, otherwise a message
        /// for the user.
        /// </summary>
        public static string Publish(string tempPdf, string destination)
        {
            var partial = destination + "." + Guid.NewGuid().ToString("N").Substring(0, 8) + ".partial";
            try
            {
                File.Copy(tempPdf, partial, overwrite: false);
                if (File.Exists(destination))
                {
                    File.Replace(partial, destination, null);
                }
                else
                {
                    File.Move(partial, destination);
                }
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
