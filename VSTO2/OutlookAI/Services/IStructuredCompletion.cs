using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    /// <summary>One answer held to a JSON schema, from text and image parts.</summary>
    internal interface IStructuredCompletion
    {
        Task<string> CompleteStructuredAsync(
            string instructions,
            IReadOnlyList<JObject> contentParts,
            string schemaName,
            JObject schema,
            string effort,
            CancellationToken cancellationToken);
    }
}
