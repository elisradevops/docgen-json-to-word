using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using JsonToWord.Services.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class MewpCoverageReporterServiceTests
    {
        [Fact]
        public void Insert_WritesNumericColumnsAsNumberCells()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow("1001", "SR1001", "Req 1001", "Fail", bugId: 9001),
                        CreateCoverageRow("1002", "SR1002", "Uncovered", "Not Run"),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var rows = sheetData.Elements<Row>().ToList();

                // Bug ID column is index 6 (G).
                var firstDataRowCells = rows[1].Elements<Cell>().ToList();
                Assert.Equal(CellValues.Number, firstDataRowCells[6].DataType?.Value);

                var secondDataRowCells = rows[2].Elements<Cell>().ToList();
                Assert.Equal(CellValues.String, secondDataRowCells[6].DataType?.Value);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_MergesDuplicateL2Columns_WhenMergeFlagEnabled()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    MergeDuplicateRequirementCells = true,
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow(
                            "5303",
                            "SR5303",
                            "Mock Requirement SR5303",
                            "Fail",
                            bugId: 10003,
                            l3ReqId: "9003",
                            l3ReqTitle: "L3 Link 9003"
                        ),
                        CreateCoverageRow(
                            "5303",
                            "SR5303",
                            "Mock Requirement SR5303",
                            "Fail",
                            bugId: 20003,
                            l4ReqId: "9103",
                            l4ReqTitle: "L4 Link 9103"
                        ),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                Assert.NotNull(mergeCells);
                var refs = mergeCells!.Elements<MergeCell>()
                    .Select(x => x.Reference?.Value ?? string.Empty)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                Assert.Contains("A2:A3", refs);
                Assert.Contains("B2:B3", refs);
                Assert.Contains("C2:C3", refs);
                Assert.Contains("D2:D3", refs);
                Assert.Contains("E2:E3", refs);
                Assert.Contains("F2:F3", refs);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_MergesDuplicateL3Columns_WhenMergeFlagEnabled()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    MergeDuplicateRequirementCells = true,
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow(
                            "9000",
                            "SR9000",
                            "Req 9000",
                            "Fail",
                            bugId: 30001,
                            l3ReqId: "9003",
                            l3ReqTitle: "L3 9003",
                            l4ReqId: "9103",
                            l4ReqTitle: "L4 9103"
                        ),
                        CreateCoverageRow(
                            "9000",
                            "SR9000",
                            "Req 9000",
                            "Fail",
                            bugId: 30002,
                            l3ReqId: "9003",
                            l3ReqTitle: "L3 9003",
                            l4ReqId: "9104",
                            l4ReqTitle: "L4 9104"
                        ),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                Assert.NotNull(mergeCells);
                var refs = mergeCells!.Elements<MergeCell>()
                    .Select(x => x.Reference?.Value ?? string.Empty)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                Assert.Contains("J2:J3", refs);
                Assert.Contains("K2:K3", refs);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_AlternatesColorByL2Group_WhenMergeFlagEnabled()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    MergeDuplicateRequirementCells = true,
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow("5301", "SR5301", "Req 1", "Fail", bugId: 10001, l3ReqId: "9001", l3ReqTitle: "L3 9001"),
                        CreateCoverageRow("5301", "SR5301", "Req 1", "Fail", bugId: 20001, l4ReqId: "9101", l4ReqTitle: "L4 9101"),
                        CreateCoverageRow("5302", "SR5302", "Req 2", "Pass", l3ReqId: "9002", l3ReqTitle: "L3 9002"),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var rows = sheetData.Elements<Row>().ToList();

                // Bug ID column (G) is not merged, so its style index reflects zebra decision.
                var bugIdStyleFirstGroupRow1 = rows[1].Elements<Cell>().ElementAt(6).StyleIndex!.Value;
                var bugIdStyleFirstGroupRow2 = rows[2].Elements<Cell>().ElementAt(6).StyleIndex!.Value;
                var bugIdStyleSecondGroupRow1 = rows[3].Elements<Cell>().ElementAt(6).StyleIndex!.Value;

                Assert.Equal(bugIdStyleFirstGroupRow1, bugIdStyleFirstGroupRow2);
                Assert.NotEqual(bugIdStyleFirstGroupRow1, bugIdStyleSecondGroupRow1);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_UsesDifferentStyleGroups_ForL2BugAndLinkedColumns()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow(
                            "7001",
                            "SR7001",
                            "Req 7001",
                            "Fail",
                            bugId: 12345,
                            l3ReqId: "9001",
                            l3ReqTitle: "L3 9001",
                            l4ReqId: "9101",
                            l4ReqTitle: "L4 9101"
                        )
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var row = sheetData.Elements<Row>().ElementAt(1);
                var cells = row.Elements<Cell>().ToList();

                var l2Style = cells[0].StyleIndex!.Value;      // A - L2 REQ ID
                var bugStyle = cells[6].StyleIndex!.Value;     // G - Bug ID
                var linkedStyle = cells[9].StyleIndex!.Value;  // J - L3 REQ ID

                Assert.NotEqual(l2Style, bugStyle);
                Assert.NotEqual(l2Style, linkedStyle);
                Assert.NotEqual(bugStyle, linkedStyle);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_UsesDifferentStyleGroups_ForL3AndL4Columns()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow(
                            "8001",
                            "SR8001",
                            "Req 8001",
                            "Pass",
                            l3ReqId: "9301",
                            l3ReqTitle: "L3 9301",
                            l4ReqId: "9401",
                            l4ReqTitle: "L4 9401"
                        )
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var row = sheetData.Elements<Row>().ElementAt(1);
                var cells = row.Elements<Cell>().ToList();

                var l3Style = cells[9].StyleIndex!.Value;  // J - L3 REQ ID
                var l4Style = cells[11].StyleIndex!.Value; // L - L4 REQ ID

                Assert.NotEqual(l3Style, l4Style);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_CreatesAdditionalL2CoverageSummaryWorksheet()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow("5001", "SR5001", "Req 5001", "Fail", bugId: 111),
                        CreateCoverageRow("5001", "SR5001", "Req 5001", "Fail", bugId: 222),
                        CreateCoverageRow("5002", "SR5002", "Req 5002", "Pass"),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var workbook = document.WorkbookPart!.Workbook;
                var sheets = workbook.Descendants<Sheet>().ToList();
                Assert.Equal(2, sheets.Count);
                Assert.Contains(sheets, s => (s.Name?.Value ?? string.Empty).Contains("Summary", StringComparison.OrdinalIgnoreCase));

                var summarySheet = sheets.First(s => (s.Name?.Value ?? string.Empty).Contains("Summary", StringComparison.OrdinalIgnoreCase));
                var summaryPart = (WorksheetPart)document.WorkbookPart!.GetPartById(summarySheet.Id!);
                var summaryRows = summaryPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

                // Header + one row per unique L2 requirement.
                Assert.Equal(3, summaryRows.Count);
                var headerValues = summaryRows[0].Elements<Cell>().Select(c => c.CellValue?.Text ?? string.Empty).ToList();
                Assert.Equal(new[] { "SR num", "L2 REQ Title", "L2 Run Status", "L2 Owner" }, headerValues);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_SummaryUsesFullTitleAndDedupesBySrWhenL2ReqIdMissing()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                        {
                            { "L2 REQ ID", "" },
                            { "SR #", "SR0054" },
                            { "L2 REQ Title", "Short split title" },
                            { "L2 REQ Full Title", "SR0054 - Full title should be used in summary" },
                            { "L2 Owner", "ESUK" },
                            { "L2 SubSystem", "Power" },
                            { "L2 Run Status", "Pass" },
                            { "Bug ID", "" },
                            { "Bug Title", "" },
                            { "Bug Responsibility", "" },
                            { "L3 REQ ID", "" },
                            { "L3 REQ Title", "" },
                            { "L4 REQ ID", "" },
                            { "L4 REQ Title", "" },
                        },
                        new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                        {
                            { "L2 REQ ID", "" },
                            { "SR #", "SR0054" },
                            { "L2 REQ Title", "Another split title row" },
                            { "L2 REQ Full Title", "SR0054 - Same full title" },
                            { "L2 Owner", "ESUK" },
                            { "L2 SubSystem", "Power" },
                            { "L2 Run Status", "Pass" },
                            { "Bug ID", 1234 },
                            { "Bug Title", "Bug 1234" },
                            { "Bug Responsibility", "ESUK" },
                            { "L3 REQ ID", "" },
                            { "L3 REQ Title", "" },
                            { "L4 REQ ID", "" },
                            { "L4 REQ Title", "" },
                        },
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var workbook = document.WorkbookPart!.Workbook;
                var summarySheet = workbook.Descendants<Sheet>()
                    .First(s => (s.Name?.Value ?? string.Empty).Contains("Summary", StringComparison.OrdinalIgnoreCase));
                var summaryPart = (WorksheetPart)document.WorkbookPart!.GetPartById(summarySheet.Id!);
                var summaryRows = summaryPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToList();

                Assert.Equal(2, summaryRows.Count); // header + single deduped SR row

                var dataCells = summaryRows[1].Elements<Cell>().ToList();
                Assert.Equal("SR0054", dataCells[0].CellValue?.Text);
                Assert.Equal("SR0054 - Full title should be used in summary", dataCells[1].CellValue?.Text);
                Assert.Equal("Pass", dataCells[2].CellValue?.Text);
                Assert.Equal("ESUK", dataCells[3].CellValue?.Text);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_HighlightsDuplicateIngestedRows()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow("6101", "SR6101", "Req 6101", "Fail", bugId: 7001, l3ReqId: "L3-1", l3ReqTitle: "L3 One", l4ReqId: "L4-1", l4ReqTitle: "L4 One"),
                        CreateCoverageRow("6102", "SR6102", "Req 6102", "Fail", bugId: 7001, l3ReqId: "L3-1", l3ReqTitle: "L3 One", l4ReqId: "L4-1", l4ReqTitle: "L4 One"),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var dataRows = sheetData.Elements<Row>().Skip(1).ToList();

                var firstRowCells = dataRows[0].Elements<Cell>().ToList();
                var secondRowCells = dataRows[1].Elements<Cell>().ToList();

                Assert.Equal(20U, firstRowCells[6].StyleIndex!.Value); // regular bug number style
                Assert.Equal(29U, secondRowCells[6].StyleIndex!.Value); // duplicate bug number style
                Assert.Equal(31U, secondRowCells[9].StyleIndex!.Value); // duplicate L3 style
                Assert.Equal(33U, secondRowCells[11].StyleIndex!.Value); // duplicate L4 style
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_HighlightsDuplicateIngestedRows_ByTitleFallbackWhenIdsAreMissing()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                );

                var model = new MewpCoverageReporterModel
                {
                    TestPlanName = "MEWP",
                    Rows = new List<Dictionary<string, object>>
                    {
                        CreateCoverageRow("6201", "SR6201", "Req 6201", "Fail",
                            bugId: null,
                            bugTitle: "Shared bug title",
                            bugResponsibility: "ESUK",
                            l3ReqId: "",
                            l3ReqTitle: "Shared L3 title",
                            l4ReqId: "",
                            l4ReqTitle: "Shared L4 title"),
                        CreateCoverageRow("6202", "SR6202", "Req 6202", "Fail",
                            bugId: null,
                            bugTitle: "Shared bug title",
                            bugResponsibility: "ESUK",
                            l3ReqId: "",
                            l3ReqTitle: "Shared L3 title",
                            l4ReqId: "",
                            l4ReqTitle: "Shared L4 title"),
                    }
                };

                service.Insert(document, "MEWP Coverage", model);

                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                var dataRows = sheetData.Elements<Row>().Skip(1).ToList();
                var firstRowCells = dataRows[0].Elements<Cell>().ToList();
                var secondRowCells = dataRows[1].Elements<Cell>().ToList();

                Assert.Equal(18U, firstRowCells[7].StyleIndex!.Value); // first Bug Title style
                Assert.Equal(27U, secondRowCells[7].StyleIndex!.Value); // duplicate Bug Title style
                Assert.Equal(31U, secondRowCells[10].StyleIndex!.Value); // duplicate L3 title style
                Assert.Equal(33U, secondRowCells[12].StyleIndex!.Value); // duplicate L4 title style
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        private static Dictionary<string, object> CreateCoverageRow(
            string l2ReqId,
            string srNumber,
            string reqName,
            string runStatus,
            int? bugId = null,
            string bugTitle = "",
            string bugResponsibility = "",
            string l3ReqId = "",
            string l3ReqTitle = "",
            string l4ReqId = "",
            string l4ReqTitle = ""
        )
        {
            var normalizedBugId = bugId.HasValue && bugId.Value > 0 ? (object)bugId.Value : string.Empty;
            var normalizedBugTitle = string.IsNullOrWhiteSpace(bugTitle) && bugId.HasValue ? $"Bug {bugId.Value}" : bugTitle;
            var normalizedBugResponsibility = string.IsNullOrWhiteSpace(bugResponsibility) && bugId.HasValue ? "ESUK" : bugResponsibility;
            return new Dictionary<string, object>
            {
                { "L2 REQ ID", l2ReqId },
                { "SR #", srNumber },
                { "L2 REQ Title", reqName },
                { "L2 REQ Full Title", reqName },
                { "L2 Owner", "ESUK" },
                { "L2 SubSystem", "Power" },
                { "L2 Run Status", runStatus },
                { "Bug ID", normalizedBugId },
                { "Bug Title", normalizedBugTitle },
                { "Bug Responsibility", normalizedBugResponsibility },
                { "L3 REQ ID", l3ReqId },
                { "L3 REQ Title", l3ReqTitle },
                { "L4 REQ ID", l4ReqId },
                { "L4 REQ Title", l4ReqTitle },
            };
        }
    }
}
