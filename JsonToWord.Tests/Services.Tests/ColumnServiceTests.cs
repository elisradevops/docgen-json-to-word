using System.Collections.Generic;
using System.Linq;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.ExcelServices;
using Xunit;

namespace JsonToWord.Services.Tests
{
    public class ColumnServiceTests
    {
        [Fact]
        public void CreateColumns_UsesColumnDefinitions()
        {
            var service = new ColumnService(new ExcelHelperService());
            var definitions = new List<ColumnDefinition>
            {
                new ColumnDefinition { Name = "A", Width = 10, Property = "TestCaseId", Group = "Test Cases" },
                new ColumnDefinition { Name = "B", Width = 20, Property = "TestCaseName", Group = "Test Cases" }
            };

            var columns = service.CreateColumns(definitions);

            var columnList = columns.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>().ToList();
            Assert.Equal(2, columnList.Count);
            Assert.Equal(10d, columnList[0].Width.Value);
            Assert.Equal(20d, columnList[1].Width.Value);
        }

        [Fact]
        public void DefineColumns_IncludesHistoryColumn_WhenAnyStepHasHistory()
        {
            var model = new TestReporterModel
            {
                TestSuites = new List<TestSuiteModel>
                {
                    new TestSuiteModel
                    {
                        SuiteName = "S1",
                        TestCases = new List<TestCaseModel>
                        {
                            new TestCaseModel
                            {
                                TestCaseId = 1,
                                TestCaseName = "TC1",
                                HistoryEntries = new List<string> { "h1", "h2" }
                            }
                        }
                    }
                }
            };

            var svc = new ColumnService(new ExcelHelperService());
            var cols = svc.DefineColumns(model, groupBySuite: true);

            Assert.Contains(cols, c => c.Property == "History");
        }

        [Fact]
        public void DefineColumns_ExcludesHistoryColumn_WhenNoStepHasHistory()
        {
            var model = new TestReporterModel
            {
                TestSuites = new List<TestSuiteModel>
                {
                    new TestSuiteModel
                    {
                        SuiteName = "S1",
                        TestCases = new List<TestCaseModel>
                        {
                            new TestCaseModel
                            {
                                TestCaseId = 1,
                                TestCaseName = "TC1",
                                TestSteps = new List<TestStepModel> { new TestStepModel { StepNo = "" } }
                            }
                        }
                    }
                }
            };

            var svc = new ColumnService(new ExcelHelperService());
            var cols = svc.DefineColumns(model, groupBySuite: true);

            Assert.DoesNotContain(cols, c => c.Property == "History");
        }

        [Fact]
        public void DefineColumns_IncludesCustomAndAssociatedColumns_WhenDataExists()
        {
            var model = new TestReporterModel
            {
                TestSuites = new List<TestSuiteModel>
                {
                    new TestSuiteModel
                    {
                        SuiteName = "S1",
                        TestCases = new List<TestCaseModel>
                        {
                            new TestCaseModel
                            {
                                TestCaseId = 1,
                                TestCaseName = "TC1",
                                ExecutionDate = "2024-01-01",
                                FailureType = "Failure",
                                Comment = "Comment",
                                RunBy = "User",
                                Configuration = "Config",
                                State = "Active",
                                StateChangeDate = "2024-01-02",
                                TestCaseResult = new TestCaseResultModel { ResultMessage = "Pass" },
                                TestSteps = new List<TestStepModel>
                                {
                                    new TestStepModel
                                    {
                                        StepNo = "1",
                                        StepAction = "Do",
                                        StepExpected = "Expect",
                                        StepRunStatus = "Pass",
                                        StepErrorMessage = "Error"
                                    }
                                },
                                HistoryEntries = new List<string> { "History" },
                                CustomFields = new Dictionary<string, object>
                                {
                                    { "customField", "value" }
                                },
                                AssociatedRequirements = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel
                                    {
                                        Id = "R1",
                                        Title = "Req1",
                                        CustomFields = new Dictionary<string, object> { { "reqField", "val" } }
                                    },
                                    new AssociatedItemModel { Id = "R2", Title = "Req2" }
                                },
                                AssociatedBugs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel
                                    {
                                        Id = "B1",
                                        Title = "Bug1",
                                        CustomFields = new Dictionary<string, object> { { "bugField", "val" } }
                                    }
                                },
                                AssociatedCRs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel
                                    {
                                        Id = "C1",
                                        Title = "CR1",
                                        CustomFields = new Dictionary<string, object> { { "crField", "val" } }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var service = new ColumnService(new ExcelHelperService());
            var cols = service.DefineColumns(model, groupBySuite: false);

            Assert.Contains(cols, c => c.Property == "SuiteName");
            Assert.Contains(cols, c => c.Property == "CustomField");
            Assert.Contains(cols, c => c.Property == "RequirementId");
            Assert.Contains(cols, c => c.Property == "RequirementName");
            Assert.Contains(cols, c => c.Property == "ReqField");
            Assert.Contains(cols, c => c.Property == "BugId");
            Assert.Contains(cols, c => c.Property == "BugName");
            Assert.Contains(cols, c => c.Property == "BugField");
            Assert.Contains(cols, c => c.Property == "CRId");
            Assert.Contains(cols, c => c.Property == "CRName");
            Assert.Contains(cols, c => c.Property == "CrField");
        }

        [Fact]
        public void GetColumnCountForeachGroup_CountsPerGroup()
        {
            var service = new ColumnService(new ExcelHelperService());
            var columns = new List<ColumnDefinition>
            {
                new ColumnDefinition { Name = "A", Group = "Test Cases" },
                new ColumnDefinition { Name = "B", Group = "Test Cases" },
                new ColumnDefinition { Name = "C", Group = "Requirements" }
            };

            var counts = service.GetColumnCountForeachGroup(columns);

            Assert.Equal(2, counts["Test Cases"]);
            Assert.Equal(1, counts["Requirements"]);
        }
    }
}
