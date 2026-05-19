using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    public sealed class ExportExcelArgs
    {
        public string FilenameHint { get; set; }

        public string SheetName { get; set; }

        public IList<ExcelColumnSpec> Columns { get; set; }

        public IList<JToken[]> Rows { get; set; }
    }
}
