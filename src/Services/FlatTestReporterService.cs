using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class FlatTestReporterService : IFlatTestReporterService
    {
        private readonly ILogger<FlatTestReporterService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;

        private static readonly string[] ColumnOrder = new[]
        {
            "PlanID",
            "PlanName",
            "Suites.parentSuite.name",
            "Suites.parentSuite.ID",
            "Suites.name",
            "Suites.id",
            "Steps.Steps.outcome",
            "Steps.Steps.stepIdentifier",
            "SubSystem",
            "TestCase.id",
            "testCase.State",
            "ResultsOutcome",
            "testCaseResults",
            "TestCaseResults.RunDateCompleted",
            "TestCaseResults.RunStats.outcome",
            "TestCaseResults.testRunId",
            "TestCaseResults.testPointId",
            "Assigned To Test",
            "tester",
            "Number Rel",
            "Loading Data",
        };

        public FlatTestReporterService(ILogger<FlatTestReporterService> logger, ISpreadsheetService spreadsheetService)
        {
            _logger = logger;
            _spreadsheetService = spreadsheetService;
        }

        public void Insert(SpreadsheetDocument document, string worksheetName, FlatTestReporterModel flatReportModel)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (flatReportModel == null)
                throw new ArgumentNullException(nameof(flatReportModel));

            var safeName = string.IsNullOrWhiteSpace(worksheetName) ? "Flat Report" : worksheetName;

            try
            {
                WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
                if (workbookPart.Workbook == null)
                {
                    workbookPart.Workbook = new Workbook();
                }

                WorksheetPart worksheetPart = _spreadsheetService.GetOrCreateWorksheetPart(workbookPart, safeName);
                worksheetPart.Worksheet = new Worksheet();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet.Append(sheetData);

                uint rowIndex = 1;
                Row headerRow = new Row { RowIndex = rowIndex };
                sheetData.Append(headerRow);
                for (int i = 0; i < ColumnOrder.Length; i++)
                {
                    string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                    Cell headerCell = _spreadsheetService.CreateTextCell(
                        $"{columnLetter}{rowIndex}",
                        ColumnOrder[i],
                        0
                    );
                    headerRow.Append(headerCell);
                }

                var rows = flatReportModel.Rows ?? new List<Dictionary<string, object>>();
                foreach (var row in rows)
                {
                    rowIndex++;
                    Row dataRow = new Row { RowIndex = rowIndex };
                    sheetData.Append(dataRow);
                    for (int i = 0; i < ColumnOrder.Length; i++)
                    {
                        string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                        string cellRef = $"{columnLetter}{rowIndex}";
                        object value;
                        string cellValue = string.Empty;
                        if (row != null && row.TryGetValue(ColumnOrder[i], out value))
                        {
                            cellValue = value?.ToString() ?? string.Empty;
                        }
                        Cell cell = _spreadsheetService.CreateTextCell(cellRef, cellValue, 0);
                        dataRow.Append(cell);
                    }
                }

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error inserting flat test reporter worksheet '{worksheetName}': {ex.Message}");
                throw;
            }
        }
    }
}
