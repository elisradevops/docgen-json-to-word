using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces.ExcelServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JsonToWord.Services.ExcelServices
{
    public class ColumnService : IColumnService
    {

        private readonly IExcelHelperService _excelHelperService;
        public ColumnService(IExcelHelperService excelHelperService) { 
            _excelHelperService = excelHelperService;
        }

        public Columns CreateColumns(List<ColumnDefinition> columnDefinitions)
        {
            Columns columns = new Columns();
            uint columnIndex = 1;

            foreach (var col in columnDefinitions)
            {
                columns.Append(new Column
                {
                    Min = columnIndex,
                    Max = columnIndex++,
                    Width = col.Width,
                    CustomWidth = true
                });
            }

            return columns;
        }

        public List<ColumnDefinition> DefineColumns(TestReporterModel testReporterModel, bool groupBySuite)
        {
            var allColumns = new List<ColumnDefinition>
            {
                new ColumnDefinition{Name ="Test Case ID", Width = 15, Property= "TestCaseId", Group= "Test Cases" },
                new ColumnDefinition { Name = "Test Case Title", Width = 30, Property = "TestCaseName", Group= "Test Cases" },

                // TestCase fields - include additional properties
                new ColumnDefinition { Name = "Execution Date", Width = 16, Property = "ExecutionDate", Group= "Test Cases" },
                new ColumnDefinition { Name = "TC Actual Result", Width = 30, Property = "TestCaseResult", Group= "Test Cases" },
                new ColumnDefinition { Name = "Failure Type", Width = 15, Property = "FailureType", Group= "Test Cases" },
                new ColumnDefinition { Name = "Test Case Comment", Width = 30, Property = "TestCaseComment", Group= "Test Cases" },

                // TestStep fields
                new ColumnDefinition { Name = "Step #", Width = 10, Property = "StepNo", Group= "Test Cases" },
                new ColumnDefinition { Name = "Step Action", Width = 40, Property = "StepAction", Group= "Test Cases" },
                new ColumnDefinition { Name = "Step Expected Result", Width = 40, Property = "StepExpected", Group= "Test Cases" },
                new ColumnDefinition { Name = "Step Actual Result", Width = 30, Property = "StepErrorMessage", Group= "Test Cases" },
                new ColumnDefinition { Name = "Step Run Status", Width = 17, Property = "StepRunStatus", Group= "Test Cases" },

                // TestCase fields - include additional properties
                new ColumnDefinition { Name = "Run By", Width = 20, Property = "RunBy", Group= "Test Cases" },
                new ColumnDefinition { Name = "Configuration", Width = 15, Property = "Configuration", Group= "Test Cases" },

            };

            if (!groupBySuite)
            {
                allColumns.Insert(0, new ColumnDefinition { Name = "Suite Name", Width = 20, Property = "SuiteName", Group= "Test Cases" });
            }

            // Add dynamic custom fields before associated requirements
            if (testReporterModel.TestSuites != null && testReporterModel.TestSuites.Any())
            {
                // Collect all unique custom field keys from all test cases
                var testCaseCustomFields = testReporterModel.TestSuites
                    .SelectMany(suite => suite.TestCases)
                    .Where(tc => tc.CustomFields != null)
                    .SelectMany(tc => tc.CustomFields)
                    .Where(field => field.Value != null)
                    .Select(field => field.Key)
                    .ToHashSet();

                // Add a column for each custom field
                foreach (var fieldKey in testCaseCustomFields)
                {
                    // Format the display name with proper spacing and capitalization
                    string displayName = string.Concat(
                        fieldKey.Select((c, i) => i > 0 && char.IsUpper(c) ? " " + c.ToString() : c.ToString()))
                        .Replace("_", " ");
                    displayName = char.ToUpper(displayName[0]) + displayName.Substring(1);

                    // Convert the field name to a proper column name (camelCase to PascalCase for property name)
                    string columnName = char.ToUpper(fieldKey[0]) + fieldKey.Substring(1);

                    // Add to columns list with a reasonable default width
                    allColumns.Add(new ColumnDefinition { Name = displayName, Width = 25, Property = columnName, Group = "Test Cases" });
                }
            }

            // Need to iterate for each test case, iterate through each associated item and custom field, add it to a hashSet of had data, else skip

            var allTestCases = testReporterModel.TestSuites.SelectMany(suite => suite.TestCases);

            var reqCustomFields = allTestCases
                .Where(tc => tc.AssociatedRequirements != null)
                .SelectMany(tc => tc.AssociatedRequirements)
                .Where(req => req.CustomFields != null)
                .SelectMany(req => req.CustomFields)
                .Where(field => field.Value != null)
                .Select(field => field.Key)
                .ToHashSet();

            var bugCustomFields = allTestCases
                .Where(tc => tc.AssociatedBugs != null)
                .SelectMany(tc => tc.AssociatedBugs)
                .Where(bug => bug.CustomFields != null)
                .SelectMany(bug => bug.CustomFields)
                .Where(field => field.Value != null)
                .Select(field => field.Key)
                .ToHashSet();

            var crCustomFields = allTestCases
                .Where(tc => tc.AssociatedCRs != null)
                .SelectMany(tc => tc.AssociatedCRs)
                .Where(cr => cr.CustomFields != null)
                .SelectMany(cr => cr.CustomFields)
                .Where(field => field.Value != null)
                .Select(field => field.Key)
                .ToHashSet();

            bool hasAnyReq = testReporterModel.TestSuites.Any(suite => suite.TestCases.Any(tc => tc.AssociatedRequirements != null && tc.AssociatedRequirements.Any()));
            bool hasAnyBug = testReporterModel.TestSuites.Any(suite => suite.TestCases.Any(tc => tc.AssociatedBugs != null && tc.AssociatedBugs.Any()));
            bool hasAnyCR = testReporterModel.TestSuites.Any(suite => suite.TestCases.Any(tc => tc.AssociatedCRs != null && tc.AssociatedCRs.Any()));

            // Adding base fields:
            if (hasAnyReq)
            {
                allColumns.Add(new ColumnDefinition { Name = "ID", Width = 20, Property = "RequirementId", Group = "Requirements" });
                allColumns.Add(new ColumnDefinition { Name = "Title", Width = 30, Property = "RequirementName", Group = "Requirements" });
                // fetch the first associated requirement
                var firstSuiteTestCaseWithRequirement = testReporterModel.TestSuites.FirstOrDefault(suite => suite.TestCases.Any(tc => tc.AssociatedRequirements != null && tc.AssociatedRequirements.Any()));
                if (firstSuiteTestCaseWithRequirement != null)
                {
                    var firstTestCaseWithRequirement = firstSuiteTestCaseWithRequirement.TestCases.FirstOrDefault(tc => tc.AssociatedRequirements != null && tc.AssociatedRequirements.Any());
                    if (firstTestCaseWithRequirement != null)
                    {
                        var firstAssociatedRequirement = firstTestCaseWithRequirement.AssociatedRequirements.FirstOrDefault();
                        if (firstAssociatedRequirement != null)
                        {
                            if (firstAssociatedRequirement.CustomFields != null)
                            {
                                foreach (var field in firstAssociatedRequirement.CustomFields)
                                {
                                    if(!reqCustomFields.Contains(field.Key))
                                    {
                                        continue;
                                    }
                                    // Format the display name with proper spacing and capitalization
                                    string displayName = string.Concat(
                                        field.Key.Select((c, i) => i > 0 && char.IsUpper(c) ? " " + c.ToString() : c.ToString()))
                                        .Replace("_", " ");
                                    displayName = char.ToUpper(displayName[0]) + displayName.Substring(1);

                                    // Convert the field name to a proper column name (camelCase to PascalCase for property name)
                                    string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                  
                                    // Add to columns list with a reasonable default width
                                    allColumns.Add(new ColumnDefinition { Name = displayName, Width = 25, Property = columnName, Group = "Requirements" });
                                }
                            }
                        }
                    }
                }
                
            }

            if (hasAnyBug)
            {
                allColumns.Add(new ColumnDefinition { Name = "ID", Width = 20, Property = "BugId", Group = "Bugs" });
                allColumns.Add(new ColumnDefinition { Name = "Title", Width = 30, Property = "BugName", Group = "Bugs" });
                var firstSuiteTestCaseWithBug = testReporterModel.TestSuites.FirstOrDefault(suite => suite.TestCases.Any(tc => tc.AssociatedBugs != null && tc.AssociatedBugs.Any()));
                if (firstSuiteTestCaseWithBug != null)
                {
                    var firstTestCaseWithBug = firstSuiteTestCaseWithBug.TestCases.FirstOrDefault(tc => tc.AssociatedBugs != null && tc.AssociatedBugs.Any());
                    if (firstTestCaseWithBug != null)
                    {
                        var firstAssociatedBug = firstTestCaseWithBug.AssociatedBugs.FirstOrDefault();
                        if (firstAssociatedBug != null)
                        {
                            if (firstAssociatedBug.CustomFields != null)
                            {
                                foreach (var field in firstAssociatedBug.CustomFields)
                                {
                                    if(!bugCustomFields.Contains(field.Key))
                                    {
                                        continue;
                                    }
                                    // Format the display name with proper spacing and capitalization
                                    string displayName = string.Concat(
                                        field.Key.Select((c, i) => i > 0 && char.IsUpper(c) ? " " + c.ToString() : c.ToString()))
                                        .Replace("_", " ");
                                    displayName = char.ToUpper(displayName[0]) + displayName.Substring(1);

                                    // Convert the field name to a proper column name (camelCase to PascalCase for property name)
                                    string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);

                                    // Add to columns list with a reasonable default width
                                    allColumns.Add(new ColumnDefinition { Name = displayName, Width = 25, Property = columnName, Group = "Bugs" });
                                }
                            }
                        }
                    }
                }
            }

            if (hasAnyCR)
            {
                allColumns.Add(new ColumnDefinition { Name = "ID", Width = 20, Property = "CRId", Group = "CRs" });
                allColumns.Add(new ColumnDefinition { Name = "Title", Width = 30, Property = "CRName", Group = "CRs" });
                var firstSuiteTestCaseWithCR = testReporterModel.TestSuites.FirstOrDefault(suite => suite.TestCases.Any(tc => tc.AssociatedCRs != null && tc.AssociatedCRs.Any()));
                if (firstSuiteTestCaseWithCR != null)
                {
                    var firstTestCaseWithCR = firstSuiteTestCaseWithCR.TestCases.FirstOrDefault(tc => tc.AssociatedCRs != null && tc.AssociatedCRs.Any());
                    if (firstTestCaseWithCR != null)
                    {
                        var firstAssociatedCR = firstTestCaseWithCR.AssociatedCRs.FirstOrDefault();
                        if (firstAssociatedCR != null)
                        {
                            if (firstAssociatedCR.CustomFields != null)
                            {
                                foreach (var field in firstAssociatedCR.CustomFields)
                                {
                                    if(!crCustomFields.Contains(field.Key))
                                    {
                                        continue;
                                    }
                                    // Format the display name with proper spacing and capitalization
                                    string displayName = string.Concat(
                                        field.Key.Select((c, i) => i > 0 && char.IsUpper(c) ? " " + c.ToString() : c.ToString()))
                                        .Replace("_", " ");
                                    displayName = char.ToUpper(displayName[0]) + displayName.Substring(1);

                                    // Convert the field name to a proper column name (camelCase to PascalCase for property name)
                                    string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                    // Add to columns list with a reasonable default width
                                    allColumns.Add(new ColumnDefinition { Name = displayName, Width = 25, Property = columnName, Group = "CRs" });
                                }
                            }
                        }
                    }
                }
            }


            // Get list of columns that actually have data
            List<string> columnsWithData = GetColumnsWithData(testReporterModel);

            // Filter columns based on which ones have data
            return allColumns.Where(col => columnsWithData.Contains(col.Property)).ToList();
        }


        public Dictionary<string, int> GetColumnCountForeachGroup(List<ColumnDefinition> columnDefinitions)
        {
            Dictionary<string, int> groupCount = new Dictionary<string, int>();
            foreach (var column in columnDefinitions)
            {
                string groupKey = column.Group ?? "Ungrouped"; // Use "Ungrouped" as a default for null groups
                
                if(groupCount.ContainsKey(groupKey))
                {
                    groupCount[groupKey]++;
                }
                else
                {
                    groupCount.Add(groupKey, 1);
                }
            }
            return groupCount;
        }

        private List<string> GetColumnsWithData(TestReporterModel testReporterModel)
        {
            // Maintain sets of columns with data, organized by groups
            HashSet<string> testCaseColumnsWithData = new HashSet<string>
            {
                "SuiteName",
                "TestCaseId",
                "TestCaseName"
            };
            
            // Set for tracking which requirement, bug, and CR columns have data
            HashSet<string> reqColumnsWithData = new HashSet<string>();
            HashSet<string> bugColumnsWithData = new HashSet<string>();
            HashSet<string> crColumnsWithData = new HashSet<string>();

            foreach (var suite in testReporterModel.TestSuites)
            {
                foreach (var testCase in suite.TestCases)
                {
                    // Test Case group columns check
                    if (!string.IsNullOrEmpty(testCase.ExecutionDate))
                        testCaseColumnsWithData.Add("ExecutionDate");

                    if (testCase.TestCaseResult != null && !string.IsNullOrEmpty(testCase.TestCaseResult.ResultMessage))
                        testCaseColumnsWithData.Add("TestCaseResult");

                    if (!string.IsNullOrEmpty(testCase.FailureType))
                        testCaseColumnsWithData.Add("FailureType");

                    if (!string.IsNullOrEmpty(testCase.Comment))
                        testCaseColumnsWithData.Add("TestCaseComment");

                    // Check steps data - part of Test Case group
                    if (testCase.TestSteps != null)
                    {
                        foreach (var step in testCase.TestSteps)
                        {
                            if (!string.IsNullOrEmpty(step.StepNo))
                                testCaseColumnsWithData.Add("StepNo");
                            if (!string.IsNullOrEmpty(step.StepAction))
                                testCaseColumnsWithData.Add("StepAction");
                            if (!string.IsNullOrEmpty(step.StepExpected))
                                testCaseColumnsWithData.Add("StepExpected");
                            if (!string.IsNullOrEmpty(step.StepRunStatus))
                                testCaseColumnsWithData.Add("StepRunStatus");
                            if (!string.IsNullOrEmpty(step.StepErrorMessage))
                                testCaseColumnsWithData.Add("StepErrorMessage");
                        }
                    }

                    if (!string.IsNullOrEmpty(testCase.RunBy))
                        testCaseColumnsWithData.Add("RunBy");
                    if (!string.IsNullOrEmpty(testCase.Configuration))
                        testCaseColumnsWithData.Add("Configuration");

                    // Check for Requirements group columns data
                    if (testCase.AssociatedRequirements != null)
                    {
                        if (testCase.AssociatedRequirements.Count > 0)
                        {
                            testCaseColumnsWithData.Add("AssociatedRequirementCount");
                            reqColumnsWithData.Add("RequirementId");
                            reqColumnsWithData.Add("RequirementName");
                        }
                        for (int i = 0; i < testCase.AssociatedRequirements.Count; i++)
                        {
                            var req = testCase.AssociatedRequirements[i];
                            if (req != null && !string.IsNullOrEmpty(req.Title))
                            {
                                reqColumnsWithData.Add($"AssociatedRequirement_{i}");
                                
                                // Add custom fields from requirements
                                if (req.CustomFields != null)
                                {
                                    foreach (var field in req.CustomFields)
                                    {
                                        if (!string.IsNullOrEmpty(_excelHelperService.GetValueString(field.Value)))
                                        {
                                            string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                            reqColumnsWithData.Add(columnName);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Check for Bugs group columns data
                    if (testCase.AssociatedBugs != null)
                    {
                        if (testCase.AssociatedBugs.Count > 0)
                        {
                            testCaseColumnsWithData.Add("AssociatedBugCount");
                            bugColumnsWithData.Add("BugId");
                            bugColumnsWithData.Add("BugName");
                        }
                        for (int i = 0; i < testCase.AssociatedBugs.Count; i++)
                        {
                            var bug = testCase.AssociatedBugs[i];
                            if (bug != null && !string.IsNullOrEmpty(bug.Title))
                            {
                                bugColumnsWithData.Add($"AssociatedBug_{i}");
                                
                                // Add custom fields from bugs
                                if (bug.CustomFields != null)
                                {
                                    foreach (var field in bug.CustomFields)
                                    {
                                        if (!string.IsNullOrEmpty(_excelHelperService.GetValueString(field.Value)))
                                        {
                                            string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                            bugColumnsWithData.Add(columnName);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Check for CRs group columns data
                    if (testCase.AssociatedCRs != null)
                    {
                        if (testCase.AssociatedCRs.Count > 0)
                        {
                            testCaseColumnsWithData.Add("AssociatedCRCount");
                            crColumnsWithData.Add("CRId");
                            crColumnsWithData.Add("CRName");
                        }
                        for (int i = 0; i < testCase.AssociatedCRs.Count; i++)
                        {
                            var cr = testCase.AssociatedCRs[i];
                            if (cr != null && !string.IsNullOrEmpty(cr.Title))
                            {
                                crColumnsWithData.Add($"AssociatedCR_{i}");
                                
                                // Add custom fields from CRs
                                if (cr.CustomFields != null)
                                {
                                    foreach (var field in cr.CustomFields)
                                    {
                                        if (!string.IsNullOrEmpty(_excelHelperService.GetValueString(field.Value)))
                                        {
                                            string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                            crColumnsWithData.Add(columnName);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Dynamically check all custom fields - part of Test Case group
                    if (testCase.CustomFields != null)
                    {
                        foreach (var field in testCase.CustomFields)
                        {
                            if (field.Value != null && !string.IsNullOrEmpty(_excelHelperService.GetValueString(field.Value)))
                            {
                                // Convert the field name to a proper column name (camelCase to PascalCase)
                                string columnName = char.ToUpper(field.Key[0]) + field.Key.Substring(1);
                                testCaseColumnsWithData.Add(columnName);
                            }
                        }
                    }
                }
            }

            // Combine all column sets into a final list
            HashSet<string> columnsWithData = new HashSet<string>(testCaseColumnsWithData);
            columnsWithData.UnionWith(reqColumnsWithData);
            columnsWithData.UnionWith(bugColumnsWithData);
            columnsWithData.UnionWith(crColumnsWithData);

            return columnsWithData.ToList();
        }

    }
}
