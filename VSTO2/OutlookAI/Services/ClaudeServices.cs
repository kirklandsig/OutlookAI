using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAI.Services
{
    public class ClaudeService
    {
        private const string ApiUrl = "https://api.anthropic.com/v1/messages";

        public enum ActionType
        {
            Proofread, Revise, Draft, Shorten, Lengthen, Formal, Friendly, Custom
        }

        public async Task<string> ProcessEmailAsync(ActionType action, string emailContent, string customPrompt = "")
        {
            try
            {
                // Enable TLS 1.2 and 1.3
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                var systemPrompt = GetSystemPrompt(action);
                var userMessage = BuildUserMessage(action, emailContent, customPrompt);
                var requestBody = CreateRequestJson(systemPrompt, userMessage);

                var request = (HttpWebRequest)WebRequest.Create(ApiUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Headers.Add("x-api-key", Config.ApiKey);
                request.Headers.Add("anthropic-version", "2023-06-01");

                using (var streamWriter = new StreamWriter(await request.GetRequestStreamAsync()))
                {
                    await streamWriter.WriteAsync(requestBody);
                }

                using (var response = await request.GetResponseAsync())
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    var responseText = await streamReader.ReadToEndAsync();
                    return ParseResponse(responseText);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var streamReader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        var errorText = streamReader.ReadToEnd();
                        throw new Exception("API Error: " + errorText);
                    }
                }
                throw new Exception("Connection error: " + ex.Message);
            }
        }

        private string CreateRequestJson(string systemPrompt, string userMessage)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"model\":\"" + EscapeJson(Config.Model) + "\",");
            sb.Append("\"max_tokens\":" + Config.MaxTokens + ",");
            sb.Append("\"system\":\"" + EscapeJson(systemPrompt) + "\",");
            sb.Append("\"messages\":[{");
            sb.Append("\"role\":\"user\",");
            sb.Append("\"content\":\"" + EscapeJson(userMessage) + "\"");
            sb.Append("}]}");
            return sb.ToString();
        }

        private string ParseResponse(string json)
        {
            var contentMarker = "\"text\":\"";
            var startIndex = json.LastIndexOf(contentMarker);
            if (startIndex == -1) throw new Exception("Could not parse API response");

            startIndex += contentMarker.Length;
            var endIndex = startIndex;
            while (endIndex < json.Length)
            {
                endIndex = json.IndexOf("\"", endIndex);
                if (endIndex == -1) break;
                if (json[endIndex - 1] != '\\') break;
                endIndex++;
            }
            if (endIndex == -1) throw new Exception("Could not parse API response");

            return UnescapeJson(json.Substring(startIndex, endIndex - startIndex));
        }

        private string EscapeJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private string UnescapeJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\n", "\n")
                .Replace("\\r", "\r")
                .Replace("\\t", "\t")
                .Replace("\\\"", "\"")
                .Replace("\\\\", "\\");
        }

        private string GetSystemPrompt(ActionType action)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return GetProofreadPrompt();
                case ActionType.Revise:
                    return GetRevisePrompt();
                case ActionType.Draft:
                    return GetDraftPrompt();
                case ActionType.Shorten:
                    return GetShortenPrompt();
                case ActionType.Lengthen:
                    return GetLengthenPrompt();
                case ActionType.Formal:
                    return GetFormalPrompt();
                case ActionType.Friendly:
                    return GetFriendlyPrompt();
                default:
                    return GetDefaultPrompt();
            }
        }

        private string GetProofreadPrompt()
        {
            return "You are a professional editor. Review the email for grammar, spelling, punctuation, and clarity issues. Return the corrected email text only. Do not add any explanations.";
        }

        private string GetRevisePrompt()
        {
            return "You are a professional writing assistant. Improve the email clarity, flow, and impact. Return only the revised email text without any explanations.";
        }

        private string GetDraftPrompt()
        {
            return "You are a professional email writer. Write a clear, professional email based on the instructions. If replying to an email thread, write only your reply - do not include the previous messages. Return only the email text you are composing.";
        }

        private string GetShortenPrompt()
        {
            return "You are a professional editor. Condense this email to be more concise while keeping essential information. Return only the shortened email text.";
        }

        private string GetLengthenPrompt()
        {
            return "You are a professional writer. Expand this email with more detail while maintaining professionalism. Return only the expanded email text.";
        }

        private string GetFormalPrompt()
        {
            return "You are a professional editor. Rewrite this email in a more formal tone suitable for business. Return only the rewritten email text.";
        }

        private string GetFriendlyPrompt()
        {
            return "You are a professional editor. Rewrite this email in a warmer, friendlier tone while remaining professional. Return only the rewritten email text.";
        }

        private string GetDefaultPrompt()
        {
            return "You are a professional email writing assistant. Help the user with their email based on their instructions. Return only the result.";
        }

        private string BuildUserMessage(ActionType action, string emailContent, string customPrompt)
        {
            if (action == ActionType.Draft)
            {
                if (!string.IsNullOrWhiteSpace(emailContent))
                {
                    // Replying to an existing email chain
                    return "Write a reply email based on these instructions:\n\n" + customPrompt +
                           "\n\n--- Email thread for context (do NOT include this in your response, just use it for context) ---\n\n" + emailContent;
                }
                // New email with no context
                return "Write an email based on these instructions:\n\n" + customPrompt;
            }
            if (action == ActionType.Custom)
            {
                return "Email content:\n\n" + emailContent + "\n\nInstructions: " + customPrompt;
            }
            return "Email to " + action.ToString().ToLower() + ":\n\n" + emailContent;
        }
    }
}
