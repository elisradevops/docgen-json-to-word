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
                using (var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                ))
                {
                    var model = new MewpCoverageReporterModel
                    {
                        TestPlanName = "MEWP",
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR1001" },
                                { "L2 REQ Title", "Req 1001" },
                                { "L2 SubSystem", "ESUK" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 9001 },
                                { "Bug Title", "Bug 9001" },
                                { "Bug Responsibility", "ESUK" },
                                { "L3 REQ ID", "L3-10" },
                                { "L3 REQ Title", "Linked L3" },
                                { "L4 REQ ID", "" },
                                { "L4 REQ Title", "" },
                            },
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR1002" },
                                { "L2 REQ Title", "Uncovered" },
                                { "L2 SubSystem", "IL" },
                                { "L2 Run Status", "Not Run" },
                                { "Bug ID", "" },
                                { "Bug Title", "" },
                                { "Bug Responsibility", "" },
                                { "L3 REQ ID", "" },
                                { "L3 REQ Title", "" },
                                { "L4 REQ ID", "" },
                                { "L4 REQ Title", "" },
                            },
                        }
                    };

                    service.Insert(document, "MEWP Coverage", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var rows = sheetData.Elements<Row>().ToList();

                    var firstDataRowCells = rows[1].Elements<Cell>().ToList();
                    Assert.Equal(CellValues.Number, firstDataRowCells[4].DataType?.Value);

                    var secondDataRowCells = rows[2].Elements<Cell>().ToList();
                    Assert.Equal(CellValues.String, secondDataRowCells[4].DataType?.Value);
                }
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
                using (var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                ))
                {
                    var model = new MewpCoverageReporterModel
                    {
                        TestPlanName = "MEWP",
                        MergeDuplicateL2Cells = true,
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR5303" },
                                { "L2 REQ Title", "Mock Requirement SR5303" },
                                { "L2 SubSystem", "Power" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 10003 },
                                { "Bug Title", "Mock Bug 10003" },
                                { "Bug Responsibility", "Elisra" },
                                { "L3 REQ ID", "9003" },
                                { "L3 REQ Title", "L3 Link 9003" },
                                { "L4 REQ ID", "" },
                                { "L4 REQ Title", "" },
                            },
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR5303" },
                                { "L2 REQ Title", "Mock Requirement SR5303" },
                                { "L2 SubSystem", "Power" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 20003 },
                                { "Bug Title", "Mock Bug 20003" },
                                { "Bug Responsibility", "ESUK" },
                                { "L3 REQ ID", "" },
                                { "L3 REQ Title", "" },
                                { "L4 REQ ID", "9103" },
                                { "L4 REQ Title", "L4 Link 9103" },
                            },
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
                }
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
                using (var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                ))
                {
                    var model = new MewpCoverageReporterModel
                    {
                        TestPlanName = "MEWP",
                        MergeDuplicateL2Cells = true,
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR5301" },
                                { "L2 REQ Title", "Req 1" },
                                { "L2 SubSystem", "Power" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 10001 },
                                { "Bug Title", "Bug 1" },
                                { "Bug Responsibility", "ESUK" },
                                { "L3 REQ ID", "9001" },
                                { "L3 REQ Title", "L3 9001" },
                                { "L4 REQ ID", "" },
                                { "L4 REQ Title", "" },
                            },
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR5301" },
                                { "L2 REQ Title", "Req 1" },
                                { "L2 SubSystem", "Power" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 20001 },
                                { "Bug Title", "Bug 2" },
                                { "Bug Responsibility", "Elisra" },
                                { "L3 REQ ID", "" },
                                { "L3 REQ Title", "" },
                                { "L4 REQ ID", "9101" },
                                { "L4 REQ Title", "L4 9101" },
                            },
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR5302" },
                                { "L2 REQ Title", "Req 2" },
                                { "L2 SubSystem", "Mission" },
                                { "L2 Run Status", "Pass" },
                                { "Bug ID", "" },
                                { "Bug Title", "" },
                                { "Bug Responsibility", "" },
                                { "L3 REQ ID", "9002" },
                                { "L3 REQ Title", "L3 9002" },
                                { "L4 REQ ID", "" },
                                { "L4 REQ Title", "" },
                            },
                        }
                    };

                    service.Insert(document, "MEWP Coverage", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var rows = sheetData.Elements<Row>().ToList();

                    // Bug ID column (E) is not merged, so its style index reflects zebra decision.
                    var bugIdStyleFirstGroupRow1 = rows[1].Elements<Cell>().ElementAt(4).StyleIndex!.Value;
                    var bugIdStyleFirstGroupRow2 = rows[2].Elements<Cell>().ElementAt(4).StyleIndex!.Value;
                    var bugIdStyleSecondGroupRow1 = rows[3].Elements<Cell>().ElementAt(4).StyleIndex!.Value;

                    Assert.Equal(bugIdStyleFirstGroupRow1, bugIdStyleFirstGroupRow2);
                    Assert.NotEqual(bugIdStyleFirstGroupRow1, bugIdStyleSecondGroupRow1);
                }
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
                using (var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                ))
                {
                    var model = new MewpCoverageReporterModel
                    {
                        TestPlanName = "MEWP",
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "L2 REQ ID", "SR7001" },
                                { "L2 REQ Title", "Req 7001" },
                                { "L2 SubSystem", "Comms" },
                                { "L2 Run Status", "Fail" },
                                { "Bug ID", 12345 },
                                { "Bug Title", "Bug 12345" },
                                { "Bug Responsibility", "ESUK" },
                                { "L3 REQ ID", "9001" },
                                { "L3 REQ Title", "L3 9001" },
                                { "L4 REQ ID", "9101" },
                                { "L4 REQ Title", "L4 9101" },
                            }
                        }
                    };

                    service.Insert(document, "MEWP Coverage", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var row = sheetData.Elements<Row>().ElementAt(1);
                    var cells = row.Elements<Cell>().ToList();

                    var l2Style = cells[0].StyleIndex!.Value;      // A - L2 REQ ID
                    var bugStyle = cells[4].StyleIndex!.Value;     // E - Bug ID
                    var linkedStyle = cells[7].StyleIndex!.Value;  // H - L3 REQ ID

                    Assert.NotEqual(l2Style, bugStyle);
                    Assert.NotEqual(l2Style, linkedStyle);
                    Assert.NotEqual(bugStyle, linkedStyle);
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }
    }
}
