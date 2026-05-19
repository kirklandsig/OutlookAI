using System;
using System.IO;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExportFilenameSanitizerTests
    {
        private static DateTimeOffset FixedNow => new DateTimeOffset(2026, 5, 18, 19, 47, 0, TimeSpan.Zero);

        [Theory]
        [InlineData("IT Creations Quotes", "IT-Creations-Quotes")]
        [InlineData("hello world", "hello-world")]
        [InlineData("Report: Q1", "Report-Q1")]
        [InlineData("a/b\\c:d*e?f\"g<h>i|j", "a-b-c-d-e-f-g-h-i-j")]
        public void Sanitize_StripsInvalidCharsAndSpaces(string input, string expectedStem)
        {
            var result = ExportFilenameSanitizer.Build(input, ".xlsx", FixedNow, _ => false);

            Assert.StartsWith(expectedStem + "-2026-05-18-1947", result);
            Assert.EndsWith(".xlsx", result);
        }

        [Fact]
        public void Sanitize_TrailingDotsAndSpacesRemoved()
        {
            var result = ExportFilenameSanitizer.Build("Report...   ", ".xlsx", FixedNow, _ => false);

            Assert.StartsWith("Report-2026-05-18-1947", result);
        }

        [Fact]
        public void Sanitize_EmptyHintFallsBackToDefault()
        {
            var result = ExportFilenameSanitizer.Build("   ", ".xlsx", FixedNow, _ => false);

            Assert.StartsWith("OutlookAI-Report-2026-05-18-1947", result);
        }

        [Fact]
        public void Sanitize_NullHintFallsBackToDefault()
        {
            var result = ExportFilenameSanitizer.Build(null, ".pdf", FixedNow, _ => false);

            Assert.StartsWith("OutlookAI-Report-2026-05-18-1947", result);
            Assert.EndsWith(".pdf", result);
        }

        [Fact]
        public void Sanitize_TruncatesHintTo80Chars()
        {
            var hint = new string('A', 200);
            var result = ExportFilenameSanitizer.Build(hint, ".xlsx", FixedNow, _ => false);
            var stem = Path.GetFileNameWithoutExtension(result);
            var aCount = 0;
            foreach (var c in stem)
            {
                if (c == 'A') aCount++;
                else break;
            }

            Assert.True(aCount <= 80, $"hint truncated to <=80 chars, got {aCount}");
        }

        [Fact]
        public void Sanitize_AppendsCollisionSuffixWhenExists()
        {
            int callCount = 0;
            bool Exists(string path)
            {
                callCount++;
                return callCount <= 2;
            }

            var result = ExportFilenameSanitizer.Build("Quotes", ".xlsx", FixedNow, Exists);

            Assert.EndsWith("Quotes-2026-05-18-1947-3.xlsx", result);
        }

        [Fact]
        public void Sanitize_DoesNotAddSuffixWhenNoCollision()
        {
            var result = ExportFilenameSanitizer.Build("Quotes", ".xlsx", FixedNow, _ => false);

            Assert.Equal("Quotes-2026-05-18-1947.xlsx", result);
        }

        [Fact]
        public void Sanitize_PreservesExtensionDot()
        {
            var resultPdf = ExportFilenameSanitizer.Build("x", "pdf", FixedNow, _ => false);
            var resultXlsx = ExportFilenameSanitizer.Build("x", ".xlsx", FixedNow, _ => false);

            Assert.EndsWith(".pdf", resultPdf);
            Assert.EndsWith(".xlsx", resultXlsx);
        }

        [Fact]
        public void Sanitize_AllInvalidHintFallsBackToDefault()
        {
            var result = ExportFilenameSanitizer.Build("\u0001<>|", ".xlsx", FixedNow, _ => false);

            Assert.Equal("OutlookAI-Report-2026-05-18-1947.xlsx", result);
        }

        [Fact]
        public void Sanitize_After999CollisionsFallsBackWithTicks()
        {
            var ticks = FixedNow.UtcDateTime.Ticks.ToString();
            var result = ExportFilenameSanitizer.Build("Quotes", ".xlsx", FixedNow, path => !path.Contains("-" + ticks));

            Assert.Equal("Quotes-2026-05-18-1947-" + ticks + ".xlsx", result);
        }
    }
}
