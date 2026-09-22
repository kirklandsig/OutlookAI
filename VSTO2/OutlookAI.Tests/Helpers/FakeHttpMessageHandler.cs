using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Scripted HTTP responder for unit tests. Callers queue responses in the
    /// order their code-under-test is expected to make requests. Records each
    /// outbound request and its body for assertions.
    /// </summary>
    public sealed class FakeHttpMessageHandler : HttpMessageHandler
    {
        private readonly Queue<Func<HttpRequestMessage, Task<HttpResponseMessage>>> _responses =
            new Queue<Func<HttpRequestMessage, Task<HttpResponseMessage>>>();

        public List<HttpRequestMessage> Requests { get; } = new List<HttpRequestMessage>();
        public List<string> RequestBodies { get; } = new List<string>();

        public void QueueJson(HttpStatusCode status, string json) =>
            Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json"),
            });

        public void QueueSse(HttpStatusCode status, string sseBody) =>
            Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(sseBody, System.Text.Encoding.UTF8, "text/event-stream"),
            });

        public void QueueText(HttpStatusCode status, string text) =>
            Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(text),
            });

        /// <summary>
        /// Queue a fully-formed <see cref="HttpContent"/> for cases where the
        /// caller needs precise control over the response body (e.g. a
        /// streaming <see cref="System.Net.Http.StreamContent"/> backed by a
        /// custom pausable stream). The content is taken as-is; the caller is
        /// responsible for setting any required headers (including
        /// <c>Content-Type</c>).
        /// </summary>
        public void QueueRaw(HttpStatusCode status, HttpContent content) =>
            Enqueue(_ => new HttpResponseMessage(status) { Content = content });

        /// <summary>Simulate a transport failure (DNS, TLS, proxy, reset).</summary>
        public void QueueException(Exception ex) =>
            Enqueue(_ => { throw ex; });

        /// <summary>
        /// Custom responder, e.g. to change shared state (another session saving)
        /// at the exact moment a request is in flight.
        /// </summary>
        public void Queue(Func<HttpRequestMessage, HttpResponseMessage> respond) =>
            Enqueue(respond);

        /// <summary>
        /// Never answers and ignores cancellation, like a proxy that accepts the
        /// connection and then stalls.
        /// </summary>
        public void QueueHang() =>
            _responses.Enqueue(_ => new TaskCompletionSource<HttpResponseMessage>().Task);

        private void Enqueue(Func<HttpRequestMessage, HttpResponseMessage> respond) =>
            _responses.Enqueue(request => Task.FromResult(respond(request)));

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            if (request.Content != null)
            {
                RequestBodies.Add(await request.Content.ReadAsStringAsync().ConfigureAwait(false));
            }
            else
            {
                RequestBodies.Add(string.Empty);
            }

            if (_responses.Count == 0)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent("{\"error\":\"no fake response queued\"}"),
                };
            }
            return await _responses.Dequeue()(request).ConfigureAwait(false);
        }
    }
}
