using System;
using System.IO;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    /// <summary>
    /// Speech-to-text via OpenAI Realtime WebSocket
    /// (<c>wss://api.openai.com/v1/realtime?model=&lt;voice-model&gt;</c>),
    /// authenticated with the same ChatGPT OAuth <c>access_token</c> the
    /// chat service uses. Replaces the legacy Whisper REST call.
    ///
    /// The OAuth bearer is accepted by the GA Realtime endpoint and bills
    /// against the user's ChatGPT consumer subscription, matching the
    /// pattern used in the AIReceptionist project.
    /// </summary>
    public sealed class RealtimeVoiceService : IDisposable
    {
        public const string RealtimeBaseUrl = "wss://api.openai.com/v1/realtime";

        // 16-kHz, 16-bit, mono PCM matches the AITaskPane WaveInEvent capture
        // and the format Realtime expects when input_audio_format = "pcm16".
        private const int AudioChunkBytes = 16 * 1024;

        private readonly CodexAuthService _auth;

        public RealtimeVoiceService(CodexAuthService auth)
        {
            _auth = auth ?? throw new ArgumentNullException(nameof(auth));
        }

        /// <summary>
        /// Streams 16-kHz / 16-bit / mono PCM bytes from <paramref name="pcm"/>
        /// to the Realtime session and returns the final user transcript.
        /// </summary>
        public async Task<string> TranscribeAsync(Stream pcm, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (pcm == null) throw new ArgumentNullException(nameof(pcm));

            var accessToken = await _auth.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
            var url = new Uri(RealtimeBaseUrl + "?model=" + Uri.EscapeDataString(Config.VoiceModel));

            using (var ws = new ClientWebSocket())
            {
                // Bearer-only. The GA Realtime endpoint rejects the
                // OpenAI-Beta: realtime=v1 header with beta_api_shape_disabled.
                ws.Options.SetRequestHeader("Authorization", "Bearer " + accessToken);

                using (var connectCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    connectCts.CancelAfter(TimeSpan.FromSeconds(20));
                    await ws.ConnectAsync(url, connectCts.Token).ConfigureAwait(false);
                }

                try
                {
                    await ConfigureSessionAsync(ws, cancellationToken).ConfigureAwait(false);
                    await StreamAudioAsync(ws, pcm, cancellationToken).ConfigureAwait(false);
                    var transcript = await ReadTranscriptAsync(ws, cancellationToken).ConfigureAwait(false);
                    await SafeCloseAsync(ws, WebSocketCloseStatus.NormalClosure, "transcribe complete", cancellationToken).ConfigureAwait(false);
                    return transcript;
                }
                catch
                {
                    await SafeCloseAsync(ws, WebSocketCloseStatus.InternalServerError, "transcribe error", cancellationToken).ConfigureAwait(false);
                    throw;
                }
            }
        }

        private static async Task ConfigureSessionAsync(ClientWebSocket ws, CancellationToken cancellationToken)
        {
            // Switch the Realtime session into transcription-only mode:
            //  - text-only output modality (no audio response)
            //  - PCM16 input audio
            //  - server-side speech-to-text via gpt-4o-mini-transcribe
            //  - turn_detection: null so we control commit/finalize manually
            var sessionUpdate = new JObject(
                new JProperty("type", "session.update"),
                new JProperty("session", new JObject(
                    new JProperty("modalities", new JArray("text")),
                    new JProperty("input_audio_format", "pcm16"),
                    new JProperty("input_audio_transcription", new JObject(
                        new JProperty("model", "gpt-4o-mini-transcribe"))),
                    new JProperty("turn_detection", JValue.CreateNull()))));

            await SendJsonAsync(ws, sessionUpdate, cancellationToken).ConfigureAwait(false);
        }

        private static async Task StreamAudioAsync(ClientWebSocket ws, Stream pcm, CancellationToken cancellationToken)
        {
            var buffer = new byte[AudioChunkBytes];
            int read;
            while ((read = await ReadFullChunkAsync(pcm, buffer, cancellationToken).ConfigureAwait(false)) > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var slice = read == buffer.Length ? buffer : CopyOf(buffer, read);
                var append = new JObject(
                    new JProperty("type", "input_audio_buffer.append"),
                    new JProperty("audio", Convert.ToBase64String(slice)));
                await SendJsonAsync(ws, append, cancellationToken).ConfigureAwait(false);
            }

            await SendJsonAsync(ws, new JObject(new JProperty("type", "input_audio_buffer.commit")), cancellationToken).ConfigureAwait(false);
            await SendJsonAsync(
                ws,
                new JObject(
                    new JProperty("type", "response.create"),
                    new JProperty("response", new JObject(
                        new JProperty("modalities", new JArray("text"))))),
                cancellationToken).ConfigureAwait(false);
        }

        private static async Task<int> ReadFullChunkAsync(Stream stream, byte[] buffer, CancellationToken cancellationToken)
        {
            int total = 0;
            while (total < buffer.Length)
            {
                var chunk = await stream.ReadAsync(buffer, total, buffer.Length - total, cancellationToken).ConfigureAwait(false);
                if (chunk == 0)
                {
                    break;
                }
                total += chunk;
            }
            return total;
        }

        private static byte[] CopyOf(byte[] source, int length)
        {
            var copy = new byte[length];
            Buffer.BlockCopy(source, 0, copy, 0, length);
            return copy;
        }

        private static async Task<string> ReadTranscriptAsync(ClientWebSocket ws, CancellationToken cancellationToken)
        {
            string transcript = null;
            bool responseCompleted = false;
            var rawBuffer = new byte[32 * 1024];
            var accumulator = new StringBuilder();

            while (ws.State == WebSocketState.Open)
            {
                cancellationToken.ThrowIfCancellationRequested();

                accumulator.Clear();
                WebSocketReceiveResult result;
                do
                {
                    result = await ws.ReceiveAsync(new ArraySegment<byte>(rawBuffer), cancellationToken).ConfigureAwait(false);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        return transcript ?? "";
                    }
                    if (result.Count > 0)
                    {
                        accumulator.Append(Encoding.UTF8.GetString(rawBuffer, 0, result.Count));
                    }
                } while (!result.EndOfMessage);

                if (accumulator.Length == 0)
                {
                    continue;
                }

                JObject evt;
                try
                {
                    evt = JObject.Parse(accumulator.ToString());
                }
                catch
                {
                    continue;
                }

                var type = (string)evt["type"];
                if (string.IsNullOrEmpty(type))
                {
                    continue;
                }

                if (type == "conversation.item.input_audio_transcription.completed")
                {
                    var t = (string)evt["transcript"];
                    if (!string.IsNullOrEmpty(t))
                    {
                        transcript = t.Trim();
                    }
                    if (responseCompleted)
                    {
                        return transcript ?? "";
                    }
                }
                else if (type == "response.completed" || type == "response.done")
                {
                    responseCompleted = true;
                    if (transcript != null)
                    {
                        return transcript;
                    }
                }
                else if (type == "error")
                {
                    var err = evt["error"] as JObject;
                    var message = err != null ? (string)err["message"] : null;
                    throw new InvalidOperationException("Realtime voice error: " + (message ?? accumulator.ToString()));
                }
            }

            if (transcript == null)
            {
                throw new InvalidOperationException(
                    "Realtime session ended without producing a transcript. Try speaking again or check your microphone.");
            }
            return transcript;
        }

        private static async Task SendJsonAsync(ClientWebSocket ws, JObject payload, CancellationToken cancellationToken)
        {
            var bytes = Encoding.UTF8.GetBytes(payload.ToString(Newtonsoft.Json.Formatting.None));
            await ws.SendAsync(
                new ArraySegment<byte>(bytes),
                WebSocketMessageType.Text,
                endOfMessage: true,
                cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        private static async Task SafeCloseAsync(ClientWebSocket ws, WebSocketCloseStatus status, string description, CancellationToken cancellationToken)
        {
            if (ws.State != WebSocketState.Open && ws.State != WebSocketState.CloseReceived)
            {
                return;
            }
            try
            {
                using (var closeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    closeCts.CancelAfter(TimeSpan.FromSeconds(5));
                    await ws.CloseAsync(status, description, closeCts.Token).ConfigureAwait(false);
                }
            }
            catch
            {
                // Best-effort close.
            }
        }

        public void Dispose()
        {
            // No long-lived resources; sockets are owned per TranscribeAsync call.
        }
    }
}
