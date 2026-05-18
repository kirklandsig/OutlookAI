using System;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Filters MAPI message classes to mail messages. Used by the
    /// Outlook Table API path in LiveOutlookSurface to drop non-mail
    /// rows (meetings, contacts, tasks, NDRs, etc.) that may live in a
    /// folder whose DefaultItemType is still olMailItem.
    ///
    /// MAPI convention: mail messages are MessageClass = "IPM.Note" or
    /// any subclass thereof ("IPM.Note.SMIME", "IPM.Note.Microsoft.Conversation.Action").
    /// Non-mail items use other prefixes (IPM.Appointment, IPM.Contact,
    /// IPM.Task, IPM.Schedule.Meeting.*) which we always reject.
    /// </summary>
    public static class TableMessageClassFilter
    {
        public static bool IsMailMessage(string messageClass)
        {
            if (string.IsNullOrWhiteSpace(messageClass)) return false;
            return messageClass.Trim().StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase);
        }
    }
}
