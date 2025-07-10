using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

        /// <summary>
        /// Determines if a content control is located under a standard heading or a custom heading.
        /// </summary>
        /// <param name="sdtBlock">The content control element to check.</param>
        /// <returns>
        /// True if the content control is under a standard heading (e.g., Heading1, Heading2) or no heading.
        /// False if the content control is under a custom heading (e.g., Appendix).
        /// </returns>
        public bool IsUnderStandardHeading(SdtBlock sdtBlock)
        {
            try
            {
                // Get the document and styles part
                var document = sdtBlock.Ancestors<Document>().FirstOrDefault();
                if (document == null) return true;

                var mainPart = document.MainDocumentPart;
                if (mainPart?.StyleDefinitionsPart?.Styles == null) return true;

                // Find the parent element that contains both the content control and headings
                // This is typically the Body element
                var body = sdtBlock.Ancestors<Body>().FirstOrDefault();
                if (body == null) return true;

                // Get all elements in the document body
                var allElements = body.Descendants().ToList();
                
                // Find the index of our content control
                int contentControlIndex = allElements.IndexOf(sdtBlock);
                if (contentControlIndex < 0) return true; // Not found, assume it's ok

                // Look backward from the content control to find the nearest paragraph
                for (int i = contentControlIndex - 1; i >= 0; i--)
                {
                    if (allElements[i] is Paragraph p && p.ParagraphProperties?.ParagraphStyleId != null)
                    {
                        var styleId = p.ParagraphProperties.ParagraphStyleId.Val;
                        var style = mainPart.StyleDefinitionsPart.Styles
                            .Elements<Style>()
                            .FirstOrDefault(s => s.StyleId == styleId);

                        if (style != null)
                        {
                            // Check if this is a heading style
                            bool isHeadingStyle = styleId.ToString().StartsWith("Heading") || 
                                (style.BasedOn != null && style.BasedOn.Val.ToString().StartsWith("Heading"));
                            
                            if (isHeadingStyle)
                            {
                                // Check if this is a custom heading style
                                bool isCustomHeading = style.CustomStyle?.Value == true;
                                
                                // Return false only if it's a custom heading
                                if (isCustomHeading)
                                {
                                    return false;
                                }
                                
                                // If we found a standard heading, we can stop looking
                                return true;
                            }
                        }
                    }
                }

                // If we didn't find any heading, return true (default case)
                return true;
            }
            catch (Exception)
            {
                // In case of any errors, default to true to allow the document processing to continue
                return true;
            }
        }
    }
}