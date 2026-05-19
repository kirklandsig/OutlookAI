using System;
using System.Collections.Generic;
using OutlookAI.Services.Tools;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Manual-control test double. Tests drive <see cref="Start"/>,
    /// <see cref="Stop"/>, and the <see cref="IAdvancedSearchHost.Completed"/>
    /// event timing directly.
    /// </summary>
    internal sealed class FakeAdvancedSearchHost : IAdvancedSearchHost
    {
        public List<StartCall> StartCalls { get; } = new List<StartCall>();
        public List<string> StopCalls { get; } = new List<string>();

        /// <summary>If non-null, thrown from Start to simulate COM failure.</summary>
        public Func<string, Exception> ThrowOnStart { get; set; }

        public event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;

        public void Start(string scope, string filter, bool searchSubFolders, string tag)
        {
            var ex = ThrowOnStart?.Invoke(tag);
            if (ex != null) throw ex;
            StartCalls.Add(new StartCall
            {
                Scope = scope,
                Filter = filter,
                SearchSubFolders = searchSubFolders,
                Tag = tag,
            });
        }

        public void Stop(string tag) => StopCalls.Add(tag);

        public void RaiseCompleted(string tag, IReadOnlyList<MessageProjectionInput> items)
        {
            Completed?.Invoke(this, new AdvancedSearchHostCompleteEventArgs
            {
                Tag = tag,
                Items = items,
            });
        }

        internal sealed class StartCall
        {
            public string Scope { get; set; }
            public string Filter { get; set; }
            public bool SearchSubFolders { get; set; }
            public string Tag { get; set; }
        }
    }
}
