using System;
using System.IO;
using System.Runtime.InteropServices;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class EmailImageCollectorTests
    {
        private const string RunPath = @"C:\Temp\run\image-01.jpg";

        private static Exception Failure(string kind)
        {
            switch (kind)
            {
                case "com": return new COMException("Outlook couldn't save it", unchecked((int)0x80004005));
                case "access": return new UnauthorizedAccessException("Access is denied.");
                case "io": return new IOException("There is not enough space on the disk.");
                default: return new ArgumentException("Illegal characters in path.");
            }
        }

        // A blocked or unsaveable attachment is still a page in the review ("can't read this file").
        [Theory]
        [InlineData("com")]
        [InlineData("access")]
        [InlineData("io")]
        [InlineData("argument")]
        public void SaveCopy_FailedSave_IsListedWithoutAPath(string kind)
        {
            var failure = Failure(kind);

            var copy = EmailImageCollector.SaveCopy(path => { throw failure; }, RunPath, "scan.jpg", isInline: true);

            Assert.NotNull(copy);
            Assert.Null(copy.Path);
            Assert.Equal("scan.jpg", copy.FileName);
            Assert.True(copy.IsInline);
        }

        [Fact]
        public void SaveCopy_Saved_KeepsThePath()
        {
            string savedTo = null;

            var copy = EmailImageCollector.SaveCopy(path => savedTo = path, RunPath, "scan.jpg", isInline: false);

            Assert.Equal(RunPath, savedTo);
            Assert.Equal(RunPath, copy.Path);
            Assert.False(copy.IsInline);
        }
    }
}
