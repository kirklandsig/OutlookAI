using System;
using System.Net;

namespace OutlookAI.Services
{
    /// <summary>The Codex backend answered with an HTTP error.</summary>
    public sealed class CodexBackendException : InvalidOperationException
    {
        public CodexBackendException(HttpStatusCode statusCode, TimeSpan? retryAfter, string message)
            : base(message)
        {
            StatusCode = statusCode;
            RetryAfter = retryAfter;
        }

        public HttpStatusCode StatusCode { get; }

        /// <summary>The server's Retry-After, when it sent one.</summary>
        public TimeSpan? RetryAfter { get; }
    }
}
