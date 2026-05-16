using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Scripted dispatcher for chat-service tests. Caller queues (name &#x2192;
    /// response JSON) in the order the model is expected to call them. Used
    /// by <c>CodexChatServiceMultiRoundTests</c> first; later picked up by
    /// any test that needs to drive <c>RunTurnAsync</c> through tool rounds
    /// without spinning up the real Outlook surface.
    /// </summary>
    public sealed class FakeToolHost : OutlookAI.Services.IToolHost
    {
        private readonly Queue<ScriptedEntry> _scripted = new Queue<ScriptedEntry>();

        public List<CallEntry> Calls { get; } = new List<CallEntry>();

        public void Queue(string name, string responseJson) =>
            _scripted.Enqueue(new ScriptedEntry(name, responseJson, throwException: null));

        public void QueueThrow(string name, Exception exception) =>
            _scripted.Enqueue(new ScriptedEntry(name, null, throwException: exception));

        public Task<string> DispatchAsync(string toolName, string argsJson, CancellationToken ct)
        {
            Calls.Add(new CallEntry(toolName, argsJson));
            if (_scripted.Count == 0)
            {
                throw new InvalidOperationException(
                    "FakeToolHost: no scripted response for " + toolName);
            }
            var entry = _scripted.Dequeue();
            if (!string.Equals(entry.Name, toolName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "FakeToolHost: expected " + entry.Name + " but model called " + toolName);
            }
            if (entry.ThrowException != null)
            {
                throw entry.ThrowException;
            }
            return Task.FromResult(entry.ResponseJson);
        }

        public sealed class CallEntry
        {
            public string Name { get; }
            public string ArgsJson { get; }
            public CallEntry(string name, string argsJson) { Name = name; ArgsJson = argsJson; }
        }

        private sealed class ScriptedEntry
        {
            public string Name { get; }
            public string ResponseJson { get; }
            public Exception ThrowException { get; }
            public ScriptedEntry(string name, string responseJson, Exception throwException)
            {
                Name = name;
                ResponseJson = responseJson;
                ThrowException = throwException;
            }
        }
    }
}
