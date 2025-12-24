using System.Collections.Generic;
using Newtonsoft.Json;

namespace JsonToWord.Models.TestReporterModels
{
    public class TestCaseModel
    {
        public int TestCaseId { get; set; }
        public string TestCaseName { get; set; }
        public string TestCaseUrl { get; set; } 
        public string? Comment { get; set; }
        public TestCaseResultModel? TestCaseResult { get; set; }
        public List<TestStepModel>? TestSteps { get; set; }
        public string? RunBy { get; set; }
        public string FailureType { get; set; }
        public string? Configuration { get; set; }
        public string? ExecutionDate { get; set; }  
        public string? State { get; set; }  
        public string? StateChangeDate { get; set; }
        public List<string>? HistoryEntries { get; set; }
        public List<AssociatedItemModel>? AssociatedRequirements { get; set; }
        public List<AssociatedItemModel>? AssociatedBugs { get; set; }
        public List<AssociatedItemModel>? AssociatedCRs { get; set; }

        [JsonExtensionData]
        public Dictionary<string, object> CustomFields { get; set; }
    }
}
