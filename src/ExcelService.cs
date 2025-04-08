using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;

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
            using (var spreadSheet = SpreadsheetDocument.Create(excelModel.LocalPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                _logger.LogInformation("Starting on doc path: " + excelModel.LocalPath);
                foreach (var contentControl in excelModel.ContentControls)
                {
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

    }
}
