﻿using System.Collections.Generic;

namespace JsonToWord.Models.TestReporterModels
{
    public class TestCaseModel
    {
        public int TestCaseId { get; set; }
        public string TestCaseName { get; set; }
        public string TestCaseUrl { get; set; } 
        public int? Priority { get; set; }
        public string FailureType { get; set; }
        public TestCaseResultModel? TestCaseResult { get; set; }
        public List<TestStepModel>? TestSteps { get; set; }
        public string? Comment { get; set; }
        public string? RunBy { get; set; }
        public string? Configuration { get; set; }
        public string? AutomationStatus { get; set; }
        public string? ExecutionDate { get; set; }  
        public string? AssignedTo { get; set; }
        public string? SubSystem { get; set; }
        public string? State { get; set; }  
        public string? StateChangeDate { get; set; }

        public List<AssociatedRequirementModel>? AssociatedRequirements { get; set; }
        public List<AssociatedBugModel>? AssociatedBugs { get; set; }
        public List<AssociatedCRModel>? AssociatedCRs { get; set; }
    }
}
