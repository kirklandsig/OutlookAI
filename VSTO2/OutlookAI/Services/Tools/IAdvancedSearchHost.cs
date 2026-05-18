using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Abstraction over <c>Application.AdvancedSearch</c> + the
    /// <c>AdvancedSearchComplete</c> event so the runner can be unit-tested
    /// without a live Outlook. Production implementation
    /// (<c>LiveAdvancedSearchHost</c>) wraps the real COM API.
    /// </summary>
    public interface IAdvancedSearchHost
    {
        /// <summary>
        /// Begin an AdvancedSearch. Must return immediately. The
        /// implementation is responsible for raising <see cref="Completed"/>
        /// later with the matching <paramref name="tag"/>.
        /// </summary>
        void Start(string scope, string filter, bool searchSubFolders, string tag);

        /// <summary>
        /// Best-effort stop. Implementations should not throw on unknown
        /// tags.
        /// </summary>
        void Stop(string tag);

        event EventHandler<AdvancedSearchHostCompleteEventArgs> Completed;
    }

    public sealed class AdvancedSearchHostCompleteEventArgs : EventArgs
    {
        public string Tag { get; set; }
        public IReadOnlyList<MessageProjectionInput> Items { get; set; }
    }
}
