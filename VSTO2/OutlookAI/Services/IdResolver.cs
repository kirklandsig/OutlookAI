using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.Services
{
    /// <summary>
    /// Maps Outlook EntryIDs (long hex strings, opaque, prompt-injection-friendly)
    /// to short stable IDs we hand to the model. Round-trips while the
    /// <see cref="IdResolver"/> instance is alive (process lifetime is fine
    /// because conversation state is per-Inspector and dies on close). Forged
    /// short IDs throw <see cref="KeyNotFoundException"/>.
    /// </summary>
    public sealed class IdResolver
    {
        private readonly ConcurrentDictionary<string, string> _shortToEntry =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, string> _entryToShort =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);

        public string Shorten(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) throw new ArgumentException(nameof(entryId));
            return _entryToShort.GetOrAdd(entryId, eid =>
            {
                using (var sha = SHA256.Create())
                {
                    var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(eid));
                    var b64 = Convert.ToBase64String(hash, 0, 9)
                                     .Replace('+', '-').Replace('/', '_');
                    _shortToEntry[b64] = eid;
                    return b64;
                }
            });
        }

        public string Resolve(string shortId)
        {
            if (_shortToEntry.TryGetValue(shortId, out var entry)) return entry;
            throw new KeyNotFoundException("Unknown OutlookAI message id: " + shortId);
        }
    }
}
