using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Pins the steering examples baked into the tool descriptions. The
    /// Inbox Copilot used to fire 20+ <c>outlook_search_messages</c> calls
    /// with empty args, returning only the newest inbox messages. Two
    /// independent fixes prevent recurrence: (1) the SSE delta-accumulation
    /// fix in <c>CodexChatService</c> (so args actually reach us), and
    /// (2) explicit examples in the tool description that show the model
    /// how to map natural language to structured fields. These tests pin (2).
    /// </summary>
    public class ToolCatalogSchemaTests
    {
        private static JObject FindTool(JArray tools, string name)
        {
            return tools.OfType<JObject>()
                .FirstOrDefault(t => (string)t["name"] == name);
        }

        [Fact]
        public void SearchMessages_Description_IncludesSteeringExamples()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var search = FindTool(tools, "outlook_search_messages");
            Assert.NotNull(search);
            var desc = (string)search["description"];
            Assert.NotNull(desc);

            // Top-level steering: must tell the model to translate natural
            // language into structured fields and NOT to send empty args.
            Assert.Contains("structured fields", desc);
            Assert.Contains("almost never what the user asked for", desc);

            // Concrete examples covering the four most common patterns:
            //   sender + date range, keyword + date upper bound,
            //   keyword + unread, flagged + attachment.
            Assert.Contains("from Alice last week", desc);
            Assert.Contains("from before 2020 with the EIN", desc);
            Assert.Contains("unread invoices", desc);
            Assert.Contains("flagged messages with attachments", desc);

            // The ISO-8601 literal we want the model to emit for the
            // canonical "before 2020" example. If a future edit drops this
            // exact string the model is much more likely to fall back to
            // freeform 'query' and miss the date filter entirely.
            Assert.Contains("'2020-01-01T00:00:00Z'", desc);
        }

        [Fact]
        public void SearchMessages_QueryField_Description_TellsModelNotToDumpStructuredInfoIntoIt()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var search = FindTool(tools, "outlook_search_messages");
            var query = search["parameters"]["properties"]["query"];
            var desc = (string)query["description"];

            Assert.Contains("Do NOT put dates, sender names", desc);
        }

        [Fact]
        public void SearchMessages_DateFields_Description_IncludesRelativeExamples()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var search = FindTool(tools, "outlook_search_messages");
            var dateFrom = (string)search["parameters"]["properties"]["date_from"]["description"];
            var dateTo   = (string)search["parameters"]["properties"]["date_to"]["description"];

            Assert.Contains("ISO-8601 UTC", dateFrom);
            Assert.Contains("last week", dateFrom);
            Assert.Contains("ISO-8601 UTC", dateTo);
            Assert.Contains("before 2020", dateTo);
            Assert.Contains("2020-01-01T00:00:00Z", dateTo);
        }

        [Fact]
        public void CountMessages_Description_IncludesSteeringExamples()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var count = FindTool(tools, "outlook_count_messages");
            var desc = (string)count["description"];

            Assert.Contains("structured fields", desc);
            // Two canonical examples cover {from + is_unread} and {date_from}.
            Assert.Contains("how many unread from Bob", desc);
            Assert.Contains("how many emails this year", desc);
        }

        [Fact]
        public void SearchMessages_Schema_AdvertisesScopeSortAndTriStates_NotOldBooleans()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var search = FindTool(tools, "outlook_search_messages");
            var props = (JObject)search["parameters"]["properties"];

            Assert.NotNull(props["scope"]);
            Assert.NotNull(props["sort_order"]);
            Assert.NotNull(props["attachment_filter"]);
            Assert.NotNull(props["read_status"]);
            Assert.NotNull(props["flag_status"]);
            Assert.NotNull(props["importance_filter"]);
            Assert.Null(props["has_attachment"]);
            Assert.Null(props["is_unread"]);
            Assert.Null(props["is_flagged"]);
            Assert.Null(props["importance"]);
        }

        [Fact]
        public void SearchMessages_Description_IncludesOldestAndAllMailExamples()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var search = FindTool(tools, "outlook_search_messages");
            var desc = (string)search["description"];

            Assert.Contains("first email ever", desc);
            Assert.Contains("scope:'all_mail'", desc);
            Assert.Contains("sort_order:'oldest'", desc);
            Assert.Contains("EIN", desc);
        }

        [Fact]
        public void CountMessages_Schema_AdvertisesScopeAndTriStates()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var count = FindTool(tools, "outlook_count_messages");
            var props = (JObject)count["parameters"]["properties"];

            Assert.NotNull(props["scope"]);
            Assert.NotNull(props["read_status"]);
            Assert.Null(props["is_unread"]);
        }

        [Fact]
        public void ReadMessages_Schema_HasIdsArrayAndBodyToggle()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var tool = FindTool(tools, "outlook_read_messages");
            Assert.NotNull(tool);
            var props = (JObject)tool["parameters"]["properties"];

            Assert.NotNull(props["ids"]);
            Assert.Equal("array", (string)props["ids"]["type"]);
            Assert.NotNull(props["include_body"]);
            Assert.Equal("boolean", (string)props["include_body"]["type"]);
            Assert.NotNull(props["max_items"]);
        }

        [Fact]
        public void ReadMessages_Description_HintsAtUseInsteadOfManyReadCalls()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var tool = FindTool(tools, "outlook_read_messages");
            var desc = (string)tool["description"] ?? "";
            Assert.Contains("read_message", desc, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AggregateMessages_Schema_HasGroupByEnumAndTopN()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var tool = FindTool(tools, "outlook_aggregate_messages");
            Assert.NotNull(tool);
            var props = (JObject)tool["parameters"]["properties"];

            Assert.NotNull(props["group_by"]);
            var groupByEnum = props["group_by"]["enum"] as JArray;
            Assert.NotNull(groupByEnum);
            var enumValues = groupByEnum.Select(t => (string)t).ToArray();
            Assert.Contains("sender", enumValues);
            Assert.Contains("day", enumValues);
            Assert.Contains("folder", enumValues);
            Assert.NotNull(props["top_n"]);
        }

        [Fact]
        public void AggregateMessages_Description_HintsAtUseInsteadOfManyCountCalls()
        {
            var tools = ToolCatalogSchema.BuildResponsesToolsArray(includeWriteTools: false);
            var tool = FindTool(tools, "outlook_aggregate_messages");
            var desc = (string)tool["description"] ?? "";
            Assert.Contains("count", desc, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
