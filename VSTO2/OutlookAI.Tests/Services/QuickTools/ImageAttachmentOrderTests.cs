using System.Linq;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImageAttachmentOrderTests
    {
        private static AttachmentInfo A(int index, string name, string cid = null, string mime = null, bool byValue = true) =>
            new AttachmentInfo { Index = index, FileName = name, ContentId = cid, MimeTag = mime, IsByValue = byValue };

        private const string Html =
            "<p>Here are the pages</p><img src=\"cid:image002.jpg@01DB1234\"><br><IMG SRC=\"CID:image001.png@01DB1234\">";

        [Fact]
        public void Order_InlineImagesByBodyPosition_ThenTheRestByIndex()
        {
            var ordered = ImageAttachmentOrder.Order(new[]
            {
                A(1, "scan-b.jpg"),
                A(2, "image001.png", "<image001.png@01DB1234>"),
                A(3, "image002.jpg", "image002.jpg@01DB1234"),
                A(4, "scan-a.jpg"),
            }, Html);

            Assert.Equal(new[] { 3, 2, 1, 4 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void Order_KeepsOnlyImagesByValue()
        {
            var ordered = ImageAttachmentOrder.Order(new[]
            {
                A(1, "contract.pdf"),
                A(2, "photo", mime: "image/jpeg"),
                A(3, "notes.docx"),
                A(4, "phone.HEIC"),
                A(5, "linked.jpg", byValue: false),
            }, null);

            Assert.Equal(new[] { 2, 4 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void Order_WithoutABody_UsesAttachmentOrder()
        {
            var ordered = ImageAttachmentOrder.Order(new[] { A(2, "b.png", "x@y"), A(1, "a.png", "z@y") }, null);

            Assert.Equal(new[] { 1, 2 }, ordered.Select(a => a.Index));
        }

        [Fact]
        public void IsInline_OnlyWhenTheBodyReferencesIt()
        {
            Assert.True(ImageAttachmentOrder.IsInline(A(2, "image001.png", "<image001.png@01DB1234>"), Html));
            Assert.False(ImageAttachmentOrder.IsInline(A(9, "other.png", "other@01DB1234"), Html));
        }

        [Theory]
        [InlineData("Scan.JPG", ".jpg")]
        [InlineData("photo.heic", ".heic")]
        [InlineData("odd.exe", ".img")]
        [InlineData(null, ".img")]
        public void SafeExtension_KeepsOnlyImageExtensions(string name, string expected)
        {
            Assert.Equal(expected, ImageAttachmentOrder.SafeExtension(name));
        }
    }
}
