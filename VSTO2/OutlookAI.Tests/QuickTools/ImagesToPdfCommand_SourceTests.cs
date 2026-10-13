using System;
using System.IO;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.QuickTools
{
    // The command drives Outlook COM and dialogs, so these pin its structure.
    public class ImagesToPdfCommand_SourceTests
    {
        private static string Source =>
            File.ReadAllText(RepoFiles.Find("OutlookAI", "QuickTools", "ImagesToPdfCommand.cs")).Replace("\r\n", "\n");

        // Decoding big photos on Outlook's UI thread would freeze it.
        [Fact]
        public void Pipeline_RunsOffTheUiThread()
        {
            Assert.Contains("await Task.Run(() => pipeline.RunAsync(", Source);
        }

        [Fact]
        public void RunFolder_IsDeletedHoweverTheRunEnds()
        {
            var src = Source;
            var finallyAt = src.LastIndexOf("finally", StringComparison.Ordinal);
            Assert.True(finallyAt > 0);
            Assert.True(src.IndexOf("QuickToolsRunFolder.Delete(runFolder)", finallyAt, StringComparison.Ordinal) > finallyAt);
        }

        [Fact]
        public void Cancel_EndsTheRunQuietly()
        {
            Assert.Contains("catch (OperationCanceledException) when (cts.IsCancellationRequested)", Source);
        }

        // The progress window is created before the first await, which installs
        // the WinForms synchronization context on a fresh Outlook UI thread.
        [Fact]
        public void ProgressWindow_ExistsBeforeTheFirstAwait()
        {
            var src = Source;
            Assert.True(src.IndexOf("new QuickToolsProgressForm(", StringComparison.Ordinal)
                < src.IndexOf("await ", StringComparison.Ordinal));
        }

        // The review form calls save from a Click handler; an exception there
        // would only reach Application.ThreadException and show the user nothing.
        [Fact]
        public void SavePdf_ReportsFailuresAndReturnsFalse()
        {
            var src = Source;
            var start = src.IndexOf("private static bool SavePdf(", StringComparison.Ordinal);
            Assert.True(start > 0);
            var end = src.IndexOf("private static string Producer(", start, StringComparison.Ordinal);
            var body = src.Substring(start, end - start);
            var write = body.IndexOf("ImagePdfWriter.WriteFile(", StringComparison.Ordinal);
            Assert.True(write > 0);
            var catchAt = body.IndexOf("catch (Exception ex)", write, StringComparison.Ordinal);
            Assert.True(catchAt > write);
            Assert.True(body.IndexOf("return false;", catchAt, StringComparison.Ordinal) > catchAt);
        }
    }
}
