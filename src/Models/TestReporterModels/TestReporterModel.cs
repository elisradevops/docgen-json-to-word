
using System.Collections.Generic;

namespace JsonToWord.Models.TestReporterModels
{
    public class TestReporterModel : ITestReporterObject
    {
        public TestReporterObjectType Type { get; set; }
        public string TestPlanName { get; set; }
        public List<TestSuiteModel> TestSuites { get; set; }
    }
}
