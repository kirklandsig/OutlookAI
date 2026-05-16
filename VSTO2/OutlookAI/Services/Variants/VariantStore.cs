using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Variants
{
    /// <summary>
    /// Per-Inspector in-memory list of generated variants. The Variants tab
    /// uses one instance per compose pane; cleared when the Inspector closes.
    /// All access is locked so background generation tasks can swap the list
    /// while the UI is reading it.
    /// </summary>
    public sealed class VariantStore
    {
        private readonly object _lock = new object();
        private readonly List<Variant> _variants = new List<Variant>();

        /// <summary>Replace the entire list with a new set (e.g. after generation).</summary>
        public void Replace(IEnumerable<Variant> variants)
        {
            if (variants == null) throw new ArgumentNullException(nameof(variants));
            lock (_lock)
            {
                _variants.Clear();
                _variants.AddRange(variants);
            }
        }

        /// <summary>
        /// Replace a single entry in-place (e.g. user edits the body of card #2).
        /// Throws <see cref="ArgumentOutOfRangeException"/> if index is invalid.
        /// </summary>
        public void Update(int index, Variant variant)
        {
            if (variant == null) throw new ArgumentNullException(nameof(variant));
            lock (_lock)
            {
                if (index < 0 || index >= _variants.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }
                _variants[index] = variant;
            }
        }

        /// <summary>Returns a copy of the current variants (caller-owned).</summary>
        public IReadOnlyList<Variant> Snapshot()
        {
            lock (_lock)
            {
                return _variants.ToArray();
            }
        }

        /// <summary>Number of stored variants. Thread-safe.</summary>
        public int Count
        {
            get { lock (_lock) { return _variants.Count; } }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _variants.Clear();
            }
        }
    }
}
