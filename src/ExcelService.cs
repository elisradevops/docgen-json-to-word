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
        #endregion

        public ExcelService(ILogger<ExcelService> logger, ITestReporterService tableService)
        {
            _logger = logger;
            _testReporterService = tableService;
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
                            throw new Exception("No data to write!");
                        }
                        foreach (var excelObject in contentControl.WordObjects)
                        {
                        
                            if (excelObject is TestReporterModel testReporter)
                            {
                                _testReporterService.Insert(spreadSheet, testReporter.TestPlanName, testReporter);
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
