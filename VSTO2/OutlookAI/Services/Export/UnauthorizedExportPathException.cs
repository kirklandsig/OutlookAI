using System;

namespace OutlookAI.Services.Export
{
    public sealed class UnauthorizedExportPathException : Exception
    {
        public UnauthorizedExportPathException(string message)
            : base(message)
        {
        }
    }
}
