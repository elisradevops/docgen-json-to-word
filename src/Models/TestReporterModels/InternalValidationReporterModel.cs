using System.Collections.Generic;
using JsonToWord.Models;

namespace JsonToWord.Models.TestReporterModels
{
    public class InternalValidationReporterModel : ITestReporterObject
    {
        public TestReporterObjectType Type { get; set; }
        public string TestPlanName { get; set; }
        public List<string> ColumnOrder { get; set; }
        public List<Dictionary<string, object>> Rows { get; set; }
    }
}
