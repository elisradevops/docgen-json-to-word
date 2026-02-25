using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.IO;

namespace JsonToWord
{
    public class ExcelService : IExcelService
    {
        #region Fields
        private readonly ILogger<ExcelService> _logger;

        private readonly ITestReporterService _testReporterService;
        private readonly IFlatTestReporterService _flatTestReporterService;
        private readonly IMewpCoverageReporterService _mewpCoverageReporterService;
        private readonly IInternalValidationReporterService _internalValidationReporterService;
        #endregion

        public ExcelService(
            ILogger<ExcelService> logger,
            ITestReporterService tableService,
            IFlatTestReporterService flatTestReporterService,
            IMewpCoverageReporterService mewpCoverageReporterService,
            IInternalValidationReporterService internalValidationReporterService
        )
        {
            _logger = logger;
            _testReporterService = tableService;
            _flatTestReporterService = flatTestReporterService;
            _mewpCoverageReporterService = mewpCoverageReporterService;
            _internalValidationReporterService = internalValidationReporterService;
        }

        public string CreateExcelDocument(ExcelModel excelModel)
        {
            _logger.LogInformation("Creating Excel Document");
            string filePath = excelModel.LocalPath;

            try
            {
                using (var spreadSheet = SpreadsheetDocument.Create(excelModel.LocalPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    _logger.LogInformation("Starting on doc path: " + excelModel.LocalPath);
                    foreach (var contentControl in excelModel.ContentControls)
                    {
                        // If there is no data to write, throw an exception
                        if (contentControl.WordObjects.Count == 0)
                        {
                            throw new Exception("No data aquired for current request. Please refine your selection");
                        }
                        foreach (var excelObject in contentControl.WordObjects)
                        {
                        
                            if (excelObject is TestReporterModel testReporter)
                            {
                                _testReporterService.Insert(spreadSheet, testReporter.TestPlanName, testReporter, contentControl.AllowGrouping);
                            }
                            if (excelObject is FlatTestReporterModel flatReporter)
                            {
                                var sheetName = string.IsNullOrWhiteSpace(flatReporter.TestPlanName)
                                    ? contentControl.Title
                                    : flatReporter.TestPlanName;
                                _flatTestReporterService.Insert(spreadSheet, sheetName, flatReporter);
                            }
                            if (excelObject is MewpCoverageReporterModel mewpCoverageReporter)
                            {
                                var sheetName = string.IsNullOrWhiteSpace(mewpCoverageReporter.TestPlanName)
                                    ? contentControl.Title
                                    : mewpCoverageReporter.TestPlanName;
                                _mewpCoverageReporterService.Insert(spreadSheet, sheetName, mewpCoverageReporter);
                            }
                            if (excelObject is InternalValidationReporterModel internalValidationReporter)
                            {
                                var sheetName = string.IsNullOrWhiteSpace(internalValidationReporter.TestPlanName)
                                    ? contentControl.Title
                                    : internalValidationReporter.TestPlanName;
                                _internalValidationReporterService.Insert(spreadSheet, sheetName, internalValidationReporter);
                            }
                        }

                        spreadSheet.WorkbookPart.Workbook.Save();
                    }
                }
                return excelModel.LocalPath;
            }
            catch(Exception ex)
            {
                _logger.LogError("Error creating Excel document: " + ex.Message);
                // Clean up the file if it exists
                if (File.Exists(filePath))
                {
                    try
                    {
                        File.Delete(filePath);
                        _logger.LogInformation($"Deleted incomplete Excel file: {filePath}");
                    }
                    catch (Exception deleteEx)
                    {
                        _logger.LogWarning(deleteEx, $"Failed to delete incomplete Excel file: {filePath}");
                    }
                }

                // Re-throw the exception
                throw;
            }

        }

    }
}
