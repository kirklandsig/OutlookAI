using System;

namespace OutlookAI.Services.Tools
{
    public sealed class ToolArgValidationException : Exception
    {
        public ToolArgValidationException(string code, string detail)
            : base(detail)
        {
            Code = code;
        }

        public string Code { get; private set; }
    }
}
