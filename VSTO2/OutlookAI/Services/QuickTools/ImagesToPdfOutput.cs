using System;
using System.IO;

namespace OutlookAI.Services.QuickTools
{
    internal static class ImagesToPdfOutput
    {
        /// <summary>
        /// Copies the finished PDF (written locally first) to where the user chose,
        /// so a network share never holds a half-written file. Null when saved,
        /// otherwise a message for the user.
        /// </summary>
        public static string Publish(string tempPdf, string destination)
        {
            try
            {
                File.Copy(tempPdf, destination, overwrite: true);
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
        }
    }
}
