using System.Collections.Generic;
using JsonToWord.Models;
using Newtonsoft.Json;

namespace JsonToWord.Models.TestReporterModels
{
    public class MewpCoverageReporterModel : ITestReporterObject
    {
        public TestReporterObjectType Type { get; set; }
        public string TestPlanName { get; set; }
        public List<string> ColumnOrder { get; set; }
        public List<Dictionary<string, object>> Rows { get; set; }
        [JsonProperty("mergeDuplicateRequirementCells")]
        public bool MergeDuplicateRequirementCells { get; set; }
    }
}
