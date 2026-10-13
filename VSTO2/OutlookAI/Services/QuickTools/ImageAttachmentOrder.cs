using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class AttachmentInfo
    {
        public int Index { get; set; }
        public string FileName { get; set; }
        public string MimeTag { get; set; }
        public string ContentId { get; set; }
        public bool IsByValue { get; set; }
    }

    /// <summary>Which attachments are images, and in what order they become pages.</summary>
    internal static class ImageAttachmentOrder
    {
        public static readonly string[] ImageExtensions =
            { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif" };

        public static bool IsImage(AttachmentInfo a) =>
            a.IsByValue
            && (ImageExtensions.Contains(Extension(a.FileName))
                || (a.MimeTag ?? "").StartsWith("image/", StringComparison.OrdinalIgnoreCase));

        public static bool IsInline(AttachmentInfo a, string html) => Position(a, html) >= 0;

        /// <summary>Images the body shows, in the order it shows them; then the rest by attachment order.</summary>
        public static List<AttachmentInfo> Order(IEnumerable<AttachmentInfo> attachments, string html)
        {
            var images = attachments.Where(IsImage).ToList();
            var inline = images.Where(a => Position(a, html) >= 0).OrderBy(a => Position(a, html));
            var rest = images.Where(a => Position(a, html) < 0).OrderBy(a => a.Index);
            return inline.Concat(rest).ToList();
        }

        public static string SafeExtension(string fileName)
        {
            var ext = Extension(fileName);
            return ImageExtensions.Contains(ext) ? ext : ".img";
        }

        private static int Position(AttachmentInfo a, string html)
        {
            var id = (a.ContentId ?? "").Trim().TrimStart('<').TrimEnd('>');
            if (id.Length == 0 || string.IsNullOrEmpty(html)) return -1;
            var cid = "cid:" + id;
            for (var at = html.IndexOf(cid, StringComparison.OrdinalIgnoreCase); at >= 0;
                 at = html.IndexOf(cid, at + 1, StringComparison.OrdinalIgnoreCase))
            {
                var end = at + cid.Length;
                if (end == html.Length || !ContinuesId(html[end])) return at;
            }
            return -1;
        }

        private static bool ContinuesId(char c) => char.IsLetterOrDigit(c) || "._-%+=$#~".IndexOf(c) >= 0;

        private static string Extension(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "";
            try
            {
                return Path.GetExtension(fileName).ToLowerInvariant();
            }
            catch (ArgumentException)
            {
                return "";
            }
        }
    }
}
