using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class ContentControlService : IContentControlService
    {
        private readonly IDocumentValidatorService _documentValidator;
        private readonly ILogger<ContentControlService> _logger;
        public ContentControlService(ILogger<ContentControlService> logger, IDocumentValidatorService documentValidatorService)
        {
            _documentValidator = documentValidatorService;
            _logger = logger;
        }

        public void ClearContentControl(WordprocessingDocument document, string contentControlTitle, bool force)
        {
            var sdtBlock = document.MainDocumentPart.Document.Body.Descendants<SdtBlock>()
                .FirstOrDefault(e => e.Descendants<SdtAlias>().FirstOrDefault()?.Val == contentControlTitle);

            if (sdtBlock == null)
                throw new Exception("Did not find a content control with the title " + contentControlTitle);

            if (!string.IsNullOrEmpty(sdtBlock.InnerText) && sdtBlock.InnerText == "Click or tap here to enter text." || force)
                RemoveAllStdContentBlock(sdtBlock);
        }

        public SdtBlock FindContentControl(WordprocessingDocument preprocessingDocument, string contentControlTitle)
        {
            var sdtBlock = preprocessingDocument.MainDocumentPart.Document.Body.Descendants<SdtBlock>().FirstOrDefault(e => e.Descendants<SdtAlias>().FirstOrDefault()?.Val == contentControlTitle);

            if (sdtBlock == null)
                throw new Exception("Did not find a content control with the title " + contentControlTitle);

            return sdtBlock;
        }

        public void RemoveContentControl(WordprocessingDocument document, string contentControlTitle)
        {
            var contentControl = FindContentControl(document, contentControlTitle);
            _logger.LogInformation("Removing content control: " + contentControlTitle);
            var errors = new List<string>();
            foreach (var element in contentControl.Elements())
            {
                if (element is SdtContentBlock)
                {
                    foreach (var innerElement in element.Elements())
                    {
                        var errorMsgs = _documentValidator.ValidateInnerElementOfContentControl(contentControlTitle, innerElement);
                        errors.AddRange(errorMsgs);
                        contentControl.Parent.InsertBefore(innerElement.CloneNode(true), contentControl);
                    }
                }
            }
            if (errors.Any())
            {
                var message = string.Join("\n", errors);
                _logger.LogError(message);
                throw new Exception($"{contentControlTitle} Content control is not valid");
            }
            contentControl.Remove();
            _logger.LogInformation("Content control removed: " + contentControlTitle);
        }

        public void RemoveAllStdContentBlock(SdtBlock sdtBlock)
        {
            var childElements = new List<OpenXmlElement>();

            foreach (var childElement in sdtBlock.ChildElements)
            {
                if (childElement is SdtContentBlock)
                {
                    childElements.Add(childElement);
                }
            }

            foreach (var childElement in childElements)
            {
                sdtBlock.RemoveChild(childElement);
            }
        }
    }
}