using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class TestReporterService : ITestReporterService
    {
        #region Fields
        private readonly ILogger<TestReporterService> _logger;
        private readonly IColumnService _columnService;
        private readonly ISpreadsheetService _spreadsheetService;
        private readonly IReportDataService _reportDataService;
        private readonly IStylesheetService _stylesheetService;
        
        #endregion
        #region Constructor
        public TestReporterService(ILogger<TestReporterService> logger, IColumnService columnService, ISpreadsheetService spreadsheetService,
             IReportDataService reportDataService, IStylesheetService stylesheetService) {
            _logger = logger;
            _columnService = columnService;
            _spreadsheetService = spreadsheetService;
            _reportDataService = reportDataService;
            _stylesheetService = stylesheetService;
        }
        #endregion
        #region Interface Implementations
        public void Insert(SpreadsheetDocument document, string worksheetName, TestReporterModel testReporterModel, bool groupBySuite)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name cannot be empty", nameof(worksheetName));
            if (testReporterModel == null)
                throw new ArgumentNullException(nameof(testReporterModel));
        
            try
            {
                // Get or add workbook part
                WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
                if (workbookPart.Workbook == null)
                {
                    workbookPart.Workbook = new Workbook();
                }
        
                // Add a WorksheetPart to the WorkbookPart if it doesn't exist
                WorksheetPart worksheetPart = _spreadsheetService.GetOrCreateWorksheetPart(workbookPart, worksheetName);
                // Clear any existing worksheet data
                worksheetPart.Worksheet = new Worksheet();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet.Append(sheetData);
        
                // Set the spreadsheet view options
                SheetViews sheetViews = new SheetViews(
                    new SheetView { WorkbookViewId = 0U, RightToLeft = false }
                );
                worksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);
        
                // Create MergeCells collection but don't append it yet
                MergeCells mergeCells = new MergeCells();

                // Define column definitions and widths
                var columnDefinitions = _columnService.DefineColumns(testReporterModel, groupBySuite);
                Columns columns = _columnService.CreateColumns(columnDefinitions);
                worksheetPart.Worksheet.InsertBefore(columns, sheetData);
                var columnCountForeachGroup = _columnService.GetColumnCountForeachGroup(columnDefinitions);

                var groupItemCounts = new Dictionary<string, int>
                {
                    { "Test Cases", testReporterModel.TestSuites.SelectMany(ts => ts.TestCases).Count() },
                    { "Requirements", testReporterModel.TestSuites.SelectMany(ts => ts.TestCases).SelectMany(tc => tc.AssociatedRequirements ?? new()).Count() },
                    { "Bugs", testReporterModel.TestSuites.SelectMany(ts => ts.TestCases).SelectMany(tc => tc.AssociatedBugs ?? new()).Count() },
                    { "CRs", testReporterModel.TestSuites.SelectMany(ts => ts.TestCases).SelectMany(tc => tc.AssociatedCRs ?? new()).Count() }
                };

                // Create header row with column names
                _spreadsheetService.CreateHeaderRow(sheetData, columnDefinitions, mergeCells, columnCountForeachGroup, groupItemCounts);
        
                // Add data rows
                uint rowIndex = 3; // Start after header row
                _reportDataService.AddDataRows(sheetData, mergeCells, testReporterModel.TestSuites, columnDefinitions, columnCountForeachGroup, ref rowIndex, worksheetPart, groupBySuite);
        
                // Only add MergeCells if there are any merge ranges
                if (mergeCells.Any())
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
                }
        
                // Add Stylesheet if it doesn't exist already
                _stylesheetService.EnsureStylesheet(workbookPart);
        
                // Save the worksheet
                worksheetPart.Worksheet.Save();

                // Save the workbook
                workbookPart.Workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error inserting grouped table into worksheet '{worksheetName}': {ex.Message}");
                throw;
            }
        }
        #endregion
    }
}
