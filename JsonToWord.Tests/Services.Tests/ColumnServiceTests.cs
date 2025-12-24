using System.Collections.Generic;
using System.Linq;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.ExcelServices;
using Xunit;

namespace JsonToWord.Services.Tests
{
    public class ColumnServiceTests
    {
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
    }
}
