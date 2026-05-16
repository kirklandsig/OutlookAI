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
        private readonly Queue<Func<HttpRequestMessage, HttpResponseMessage>> _responses =
            new Queue<Func<HttpRequestMessage, HttpResponseMessage>>();

        public List<HttpRequestMessage> Requests { get; } = new List<HttpRequestMessage>();
        public List<string> RequestBodies { get; } = new List<string>();

        public void QueueJson(HttpStatusCode status, string json) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json"),
            });

        public void QueueSse(HttpStatusCode status, string sseBody) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(sseBody, System.Text.Encoding.UTF8, "text/event-stream"),
            });

        public void QueueText(HttpStatusCode status, string text) =>
            _responses.Enqueue(_ => new HttpResponseMessage(status)
            {
                Content = new StringContent(text),
            });

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
            return _responses.Dequeue()(request);
        }
    }
}
