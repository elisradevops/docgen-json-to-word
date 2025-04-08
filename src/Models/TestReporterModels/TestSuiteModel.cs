using System.Collections.Generic;

namespace JsonToWord.Models.TestReporterModels
{
    public class TestSuiteModel
    {
        public string SuiteName { get; set; }
        public List<TestCaseModel> TestCases { get; set; }

    }
}
