using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// A PDF 1.4 with one JPEG per Letter page (landscape when wider than tall),
    /// fitted inside 18 pt margins. The JPEG bytes are embedded as they are, one
    /// file in memory at a time.
    /// </summary>
    internal static class ImagePdfWriter
    {
        public const double LetterShort = 612;
        public const double LetterLong = 792;
        public const double Margin = 18;

        public static void WriteFile(IReadOnlyList<string> jpegPaths, string title, string producer, string path)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Write(jpegPaths, title, producer, stream);
            }
        }

        public static void Write(IReadOnlyList<string> jpegPaths, string title, string producer, Stream output)
        {
            if (jpegPaths == null || jpegPaths.Count == 0)
                throw new ArgumentException("At least one page is needed.", nameof(jpegPaths));

            var pdf = new PdfOut(output);
            var count = jpegPaths.Count;
            var size = 4 + 3 * count;
            var offsets = new long[size];

            pdf.Write("%PDF-1.4\n");
            pdf.Write(new byte[] { (byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n' });

            offsets[1] = pdf.Position;
            pdf.Write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

            var kids = new StringBuilder();
            for (var p = 0; p < count; p++) kids.Append(p == 0 ? "" : " ").Append(PageObject(p)).Append(" 0 R");
            offsets[2] = pdf.Position;
            pdf.Write("2 0 obj\n<< /Type /Pages /Kids [" + kids + "] /Count " + count + " >>\nendobj\n");

            offsets[3] = pdf.Position;
            pdf.Write("3 0 obj\n<< /Title " + TextString(title ?? "") + " /Producer " + TextString(producer ?? "") + " >>\nendobj\n");

            for (var p = 0; p < count; p++)
            {
                var jpeg = File.ReadAllBytes(jpegPaths[p]);
                int width, height, components;
                if (!JpegHeader.TryRead(jpeg, out width, out height, out components))
                    throw new InvalidDataException("Page " + (p + 1) + " is not a JPEG.");
                var colorSpace = components == 1 ? "/DeviceGray" : components == 3 ? "/DeviceRGB" : null;
                if (colorSpace == null)
                    throw new NotSupportedException("Page " + (p + 1) + " has " + components + " color components.");

                var landscape = width > height;
                var pageW = landscape ? LetterLong : LetterShort;
                var pageH = landscape ? LetterShort : LetterLong;
                var scale = Math.Min((pageW - 2 * Margin) / width, (pageH - 2 * Margin) / height);
                var drawW = width * scale;
                var drawH = height * scale;
                var content = "q " + N(drawW) + " 0 0 " + N(drawH) + " " + N((pageW - drawW) / 2) + " " + N((pageH - drawH) / 2) + " cm /Im1 Do Q";

                var pageObject = PageObject(p);
                offsets[pageObject] = pdf.Position;
                pdf.Write(pageObject + " 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + N(pageW) + " " + N(pageH)
                    + "] /Resources << /XObject << /Im1 " + (pageObject + 1) + " 0 R >> >> /Contents " + (pageObject + 2) + " 0 R >>\nendobj\n");

                offsets[pageObject + 1] = pdf.Position;
                pdf.Write((pageObject + 1) + " 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width + " /Height " + height
                    + " /ColorSpace " + colorSpace + " /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpeg.Length + " >>\nstream\n");
                pdf.Write(jpeg);
                pdf.Write("\nendstream\nendobj\n");

                offsets[pageObject + 2] = pdf.Position;
                pdf.Write((pageObject + 2) + " 0 obj\n<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream\nendobj\n");
            }

            var xref = pdf.Position;
            var table = new StringBuilder("xref\n0 " + size + "\n0000000000 65535 f \n");
            for (var o = 1; o < size; o++)
            {
                table.Append(offsets[o].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            }
            pdf.Write(table.ToString());
            pdf.Write("trailer\n<< /Size " + size + " /Root 1 0 R /Info 3 0 R >>\nstartxref\n"
                + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        }

        /// <summary>A PDF text string as UTF-16BE hex with a byte-order mark (any characters).</summary>
        internal static string TextString(string text)
        {
            var sb = new StringBuilder("<FEFF");
            foreach (var b in Encoding.BigEndianUnicode.GetBytes(text)) sb.Append(b.ToString("X2", CultureInfo.InvariantCulture));
            return sb.Append('>').ToString();
        }

        private static int PageObject(int index) => 4 + 3 * index;

        private static string N(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

        private sealed class PdfOut
        {
            private readonly Stream _stream;

            public PdfOut(Stream stream)
            {
                _stream = stream;
            }

            public long Position { get; private set; }

            public void Write(string ascii) => Write(Encoding.ASCII.GetBytes(ascii));

            public void Write(byte[] bytes)
            {
                _stream.Write(bytes, 0, bytes.Length);
                Position += bytes.Length;
            }
        }
    }
}
