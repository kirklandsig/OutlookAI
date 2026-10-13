using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using OutlookAI.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class CollectedImage
    {
        public CollectedImage(string path, string fileName, bool isInline)
        {
            Path = path;
            FileName = fileName ?? "";
            IsInline = isInline;
        }

        public string Path { get; }
        public string FileName { get; }
        public bool IsInline { get; }
    }

    internal sealed class CollectedImages
    {
        public CollectedImages(IReadOnlyList<CollectedImage> images, bool isRichText)
        {
            Images = images;
            IsRichText = isRichText;
        }

        public IReadOnlyList<CollectedImage> Images { get; }

        /// <summary>RTF emails carry pasted images as OLE attachments, which this version skips.</summary>
        public bool IsRichText { get; }
    }

    internal static class EmailImageCollector
    {
        private const string MimeTagProperty = "http://schemas.microsoft.com/mapi/proptag/0x370E001F";
        private const string ContentIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

        /// <summary>Saves the email's image attachments into the run folder, in page order. UI thread only.</summary>
        public static CollectedImages Collect(Outlook.MailItem mail, string runFolder)
        {
            var attachments = mail.Attachments;
            var infos = new List<AttachmentInfo>();
            for (var i = 1; i <= attachments.Count; i++)
            {
                var attachment = attachments[i];
                infos.Add(new AttachmentInfo
                {
                    Index = i,
                    FileName = Safe(() => attachment.FileName),
                    IsByValue = Safe(() => attachment.Type) == Outlook.OlAttachmentType.olByValue,
                    MimeTag = Property(attachment, MimeTagProperty),
                    ContentId = Property(attachment, ContentIdProperty),
                });
            }

            var html = Safe(() => mail.HTMLBody);
            var images = new List<CollectedImage>();
            var n = 0;
            foreach (var info in ImageAttachmentOrder.Order(infos, html))
            {
                n++;
                var path = Path.Combine(runFolder, "image-" + n.ToString("D2", CultureInfo.InvariantCulture)
                    + ImageAttachmentOrder.SafeExtension(info.FileName));
                try
                {
                    attachments[info.Index].SaveAsFile(path);
                }
                catch (COMException ex)
                {
                    TraceLog.Write("QuickTools: SaveAsFile failed (0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")", "QuickTools");
                    continue;
                }
                images.Add(new CollectedImage(path, info.FileName, ImageAttachmentOrder.IsInline(info, html)));
            }

            var richText = Safe(() => mail.BodyFormat) == Outlook.OlBodyFormat.olFormatRichText;
            return new CollectedImages(images, richText);
        }

        private static string Property(Outlook.Attachment attachment, string tag)
        {
            try
            {
                return attachment.PropertyAccessor.GetProperty(tag) as string;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static T Safe<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (COMException)
            {
                return default(T);
            }
        }
    }
}
