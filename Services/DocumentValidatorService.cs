using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;
using JsonToWord.Services.Interfaces;

namespace JsonToWord.Services
{
    public class DocumentValidatorService:IDocumentValidatorService
    {
        private readonly ILogger<DocumentValidatorService> _logger;
        private readonly OpenXmlValidator _validator;

        public DocumentValidatorService(ILogger<DocumentValidatorService> logger)
        {
            _logger = logger;
            _validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        }

        #region Public Methods

        public bool ValidateDocument(WordprocessingDocument document)
        {
            var errors = _validator.Validate(document).ToList();
            
            foreach(var error in errors)
            {
                _logger.LogError("Document Error: {Error} at {Path}, {Uri}",
                    error.Description,
                    error.Path.XPath, error.Path.PartUri);
            }

            return !errors.Any();
        }

    
        public List<string> ValidateInnerElementOfContentControl(string contentControlTitle, OpenXmlElement element)
        {
            var errorMsgs = new List<string>();
            if (element == null)
            {
                _logger.LogError("element is null");
                return errorMsgs;
            }
            // altChunk is a special element that is not validated by the OpenXmlValidator
            if(element.LocalName == "altChunk")
            {
                return errorMsgs;
            }
            var errors = _validator.Validate(element).ToList();

            if (errors.Any())
            {
                foreach (var error in errors)
                {
                    var message = $"Element {element.LocalName} Error: {error.Description} at {error.Path.XPath}; {error.Path.PartUri}";
                    errorMsgs.Add(message);
                }
            }
            return errorMsgs;
        }
        #endregion

    }
}
