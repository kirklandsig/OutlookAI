using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImagePdfWriterTests : IDisposable
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-pdf", Path.GetRandomFileName());

        public ImagePdfWriterTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private string Jpeg(string name, int width, int height)
        {
            var path = Path.Combine(_dir, name);
            using (var bitmap = new Bitmap(width, height))
            {
                using (var g = Graphics.FromImage(bitmap)) g.Clear(Color.White);
                JpegFiles.Save(bitmap, path, 85);
            }
            return path;
        }

        private string Pdf(IReadOnlyList<string> pages, string title = "Lease docs")
        {
            using (var stream = new MemoryStream())
            {
                ImagePdfWriter.Write(pages, title, "OutlookAI v2.3.0", stream);
                return Latin1.GetString(stream.ToArray());
            }
        }

        [Fact]
        public void JpegHeader_ReadsSizeAndComponents()
        {
            int w, h, c;
            Assert.True(JpegHeader.TryRead(File.ReadAllBytes(Jpeg("a.jpg", 300, 400)), out w, out h, out c));
            Assert.Equal(300, w);
            Assert.Equal(400, h);
            Assert.Equal(3, c);
        }

        [Fact]
        public void JpegHeader_RejectsOtherBytes()
        {
            int w, h, c;
            Assert.False(JpegHeader.TryRead(Encoding.ASCII.GetBytes("not a jpeg"), out w, out h, out c));
            Assert.False(JpegHeader.TryRead(new byte[] { 0xFF, 0xD8, 0xFF }, out w, out h, out c));
        }

        [Fact]
        public void Write_HasOnePagePerImage_InTheRightOrientation()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 300, 400), Jpeg("l.jpg", 400, 300) });

            Assert.StartsWith("%PDF-1.4", pdf);
            Assert.Contains("/Count 2", pdf);
            Assert.Contains("/MediaBox [0 0 612 792]", pdf);
            Assert.Contains("/MediaBox [0 0 792 612]", pdf);
            Assert.Contains("/Width 300 /Height 400", pdf);
            // Fitted inside 18 pt margins and centered: 300x400 px at 1.89 pt/px.
            Assert.Contains("q 567 0 0 756 22.5 18 cm /Im1 Do Q", pdf);
            Assert.EndsWith("%%EOF\n", pdf);
        }

        [Fact]
        public void Write_XrefOffsetsPointAtTheirObjects()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 300, 400), Jpeg("l.jpg", 400, 300) });

            var startxref = int.Parse(Regex.Match(pdf, @"startxref\n(\d+)\n%%EOF").Groups[1].Value);
            Assert.Equal("xref", pdf.Substring(startxref, 4));
            var entries = Regex.Matches(pdf.Substring(startxref), @"(\d{10}) 00000 n ").Cast<Match>().ToList();
            Assert.Equal(9, entries.Count); // catalog, pages, info, 2 x (page, image, content)
            for (var i = 0; i < entries.Count; i++)
            {
                var offset = int.Parse(entries[i].Groups[1].Value);
                Assert.StartsWith((i + 1) + " 0 obj", pdf.Substring(offset));
            }
        }

        [Fact]
        public void Write_EmbedsTheJpegBytesAsIs()
        {
            var path = Jpeg("p.jpg", 300, 400);
            var bytes = File.ReadAllBytes(path);

            var pdf = Pdf(new[] { path });

            var start = pdf.IndexOf("/Length " + bytes.Length + " >>\nstream\n", StringComparison.Ordinal);
            Assert.True(start > 0);
            var data = pdf.IndexOf("stream\n", start, StringComparison.Ordinal) + "stream\n".Length;
            Assert.Equal(bytes, Latin1.GetBytes(pdf.Substring(data, bytes.Length)));
        }

        [Fact]
        public void Write_TitleCanHoldAnyCharacters()
        {
            var pdf = Pdf(new[] { Jpeg("p.jpg", 30, 40) }, title: "Résumé");

            Assert.Contains("/Title <FEFF005200E9", pdf);
        }

        [Fact]
        public void Write_NeedsAtLeastOnePage_AndOnlyJpegs()
        {
            Assert.Throws<ArgumentException>(() => Pdf(new string[0]));
            var notJpeg = Path.Combine(_dir, "x.jpg");
            File.WriteAllText(notJpeg, "nope");
            Assert.Throws<InvalidDataException>(() => Pdf(new[] { notJpeg }));
        }

        [Fact]
        public void Publish_CopiesTheFile()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var destination = Path.Combine(_dir, "out.pdf");

            Assert.Null(ImagesToPdfOutput.Publish(temp, destination));
            Assert.Equal("%PDF-1.4", File.ReadAllText(destination));
        }

        [Fact]
        public void Publish_LockedDestination_SaysCloseIt()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var destination = Path.Combine(_dir, "open-in-viewer.pdf");
            File.WriteAllText(destination, "old");
            using (new FileStream(destination, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var message = ImagesToPdfOutput.Publish(temp, destination);

                Assert.NotNull(message);
                Assert.Contains("close it", message);
            }
            Assert.Equal("old", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(_dir, "*.partial"));
        }

        [Fact]
        public void Publish_ReplacesAnExistingFile_AndLeavesNoPartial()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var destination = Path.Combine(_dir, "lease.pdf");
            File.WriteAllText(destination, "old");

            Assert.Null(ImagesToPdfOutput.Publish(temp, destination));
            Assert.Equal("%PDF-1.4", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(_dir, "*.partial"));
        }

        // A share that drops mid-copy: a locked region in the source makes the copy fail partway.
        [Fact]
        public void Publish_CopyFailsMidway_KeepsTheExistingFile()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllBytes(temp, new byte[8 * 1024 * 1024]);
            var destination = Path.Combine(_dir, "lease.pdf");
            File.WriteAllText(destination, "old");
            using (var holder = new FileStream(temp, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                holder.Lock(4 * 1024 * 1024, 1024 * 1024);
                try
                {
                    Assert.NotNull(ImagesToPdfOutput.Publish(temp, destination));
                }
                finally
                {
                    holder.Unlock(4 * 1024 * 1024, 1024 * 1024);
                }
            }
            Assert.Equal("old", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(_dir, "*.partial"));
        }

        // Hardened shares (an explicit OWNER RIGHTS ACE) withhold WRITE_DAC, so ReplaceFile
        // can't merge the old file's ACL unless told to ignore merge errors.
        [Fact]
        public void Publish_OverAFileWhereTheOwnerLacksWriteDac_Succeeds()
        {
            var temp = Path.Combine(_dir, "temp.pdf");
            File.WriteAllText(temp, "%PDF-1.4");
            var folder = Path.Combine(_dir, "hardened");
            Directory.CreateDirectory(folder);
            var security = new DirectorySecurity();
            security.SetAccessRuleProtection(true, false);
            var inherit = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
            security.AddAccessRule(new FileSystemAccessRule(WindowsIdentity.GetCurrent().User, FileSystemRights.Modify, inherit, PropagationFlags.None, AccessControlType.Allow));
            security.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier("S-1-3-4"), FileSystemRights.Modify, inherit, PropagationFlags.None, AccessControlType.Allow));
            new DirectoryInfo(folder).SetAccessControl(security);
            var destination = Path.Combine(folder, "lease.pdf");
            File.WriteAllText(destination, "old");

            Assert.Null(ImagesToPdfOutput.Publish(temp, destination));

            Assert.Equal("%PDF-1.4", File.ReadAllText(destination));
            Assert.Equal(new[] { destination }, Directory.GetFiles(folder));
        }

        [Fact]
        public void Publish_MissingSource_KeepsTheExistingFile()
        {
            var destination = Path.Combine(_dir, "lease.pdf");
            File.WriteAllText(destination, "old");

            Assert.NotNull(ImagesToPdfOutput.Publish(Path.Combine(_dir, "gone.pdf"), destination));
            Assert.Equal("old", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(_dir, "*.partial"));
        }
    }
}
