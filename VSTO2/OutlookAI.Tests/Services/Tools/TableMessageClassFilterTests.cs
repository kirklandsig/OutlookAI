using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Phase 3b Table-API path: the Outlook Table for a default-mail folder
    /// returns rows for every item type stored in the folder (mail,
    /// meeting requests, calendar invites pinned to a mail folder, etc.).
    /// We need to keep only the actual mail messages so our oldest /
    /// newest / filter results are not contaminated by non-message items.
    /// MAPI's MessageClass property starts with "IPM.Note" for mail and
    /// its subclasses (signed, encrypted, conversation actions, etc.).
    /// </summary>
    public class TableMessageClassFilterTests
    {
        [Theory]
        [InlineData("IPM.Note")]
        [InlineData("IPM.Note.SMIME")]
        [InlineData("IPM.Note.SMIME.MultipartSigned")]
        [InlineData("IPM.Note.Microsoft.Conversation.Action")]
        [InlineData("ipm.note")]                          // case-insensitive
        [InlineData("IPM.NOTE")]
        public void IsMailMessage_MailClasses_ReturnsTrue(string messageClass)
        {
            Assert.True(TableMessageClassFilter.IsMailMessage(messageClass));
        }

        [Theory]
        [InlineData("IPM.Appointment")]
        [InlineData("IPM.Contact")]
        [InlineData("IPM.Task")]
        [InlineData("IPM.Activity")]
        [InlineData("IPM.Schedule.Meeting.Request")]
        [InlineData("IPM.Schedule.Meeting.Resp.Pos")]
        [InlineData("IPM.Document")]
        [InlineData("REPORT.IPM.Note.NDR")]               // not actually IPM.Note*
        public void IsMailMessage_NonMailClasses_ReturnsFalse(string messageClass)
        {
            Assert.False(TableMessageClassFilter.IsMailMessage(messageClass));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void IsMailMessage_NullEmptyOrWhitespace_ReturnsFalse(string messageClass)
        {
            Assert.False(TableMessageClassFilter.IsMailMessage(messageClass));
        }
    }
}
