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
        
        // Dictionary to store content control heading status
        // Key: Content control title
        // Value: True if under standard heading, False if under custom heading or following another content control
        private Dictionary<string, bool> _contentControlHeadingStatus = new Dictionary<string, bool>();
        
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
        /// Maps the heading status for a content control
        /// </summary>
        /// <param name="contentControlTitle">Content control title</param>
        /// <param name="isUnderStandardHeading">True if under standard heading, false otherwise</param>
        public void MapContentControlHeading(string contentControlTitle, bool isUnderStandardHeading)
        {
            _contentControlHeadingStatus[contentControlTitle] = isUnderStandardHeading;
        }

        /// <summary>
        /// Gets the heading status for a content control from the map
        /// </summary>
        /// <param name="contentControlTitle">Content control title</param>
        /// <returns>True if under standard heading, false otherwise. Returns true if not found in map.</returns>
        public bool GetContentControlHeadingStatus(string contentControlTitle)
        {
            if (string.IsNullOrEmpty(contentControlTitle))
                return true;
                
            if (_contentControlHeadingStatus.TryGetValue(contentControlTitle, out bool status))
                return status;
                
            return true; // Default to true if not found
        }
        
        /// <summary>
        /// Clears the content control heading status map
        /// </summary>
        public void ClearContentControlHeadingMap()
        {
            _contentControlHeadingStatus.Clear();
            _logger.LogInformation("Content control heading map cleared");
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
                // Get the content control title
                var title = sdtBlock.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    _logger.LogWarning("Content control without title, assuming it's under a standard heading");
                    return true;
                }
                
                // Check if we have the heading status in our map
                if (_contentControlHeadingStatus.TryGetValue(title, out bool status))
                {
                    _logger.LogInformation($"Using mapped heading status for {title}: {status}");
                    return status;
                }

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

                // Check if there's another content control before this one
                for (int i = contentControlIndex - 1; i >= 0; i--)
                {
                    if (allElements[i] is SdtBlock previousSdtBlock)
                    {
                        // Found a previous content control, get its heading status
                        var previousTitle = previousSdtBlock.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.ToString();
                        if (!string.IsNullOrEmpty(previousTitle) && _contentControlHeadingStatus.TryGetValue(previousTitle, out bool previousStatus))
                        {
                            // Inherit the heading status of the previous content control
                            _logger.LogInformation($"Content control {title} follows {previousTitle}, inheriting status: {previousStatus}");
                            return previousStatus;
                        }
                        
                        // If previous content control status isn't in the map, check its heading directly
                        _logger.LogInformation($"Content control {title} follows {previousTitle}, determining heading status from scratch");
                        return DetermineHeadingStatusFromDocument(document, previousSdtBlock);
                    }
                    
                    // Look backward from the content control to find the nearest paragraph with heading style
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
                                    _logger.LogInformation($"Content control {title} is under a custom heading, returning false");
                                    return false;
                                }
                                
                                // Found a standard heading, return true
                                _logger.LogInformation($"Content control {title} is under a standard heading, returning true");
                                return true;
                            }
                        }
                    }
                }

                // If we didn't find any heading or previous content control, return true (default case)
                _logger.LogInformation($"Content control {title} has no heading or previous content control, returning true");
                return true;
            }
            catch (Exception ex)
            {
                // In case of any errors, default to true to allow the document processing to continue
                _logger.LogError(ex, "Error checking if content control is under standard heading, defaulting to true");
                return true;
            }
        }
        
        // Helper method to determine the heading status directly from document structure
        private bool DetermineHeadingStatusFromDocument(Document document, SdtBlock sdtBlock)
        {
            try
            {
                var mainPart = document.MainDocumentPart;
                if (mainPart?.StyleDefinitionsPart?.Styles == null) return true;
                
                var body = sdtBlock.Ancestors<Body>().FirstOrDefault();
                if (body == null) return true;
                
                var allElements = body.Descendants().ToList();
                int contentControlIndex = allElements.IndexOf(sdtBlock);
                if (contentControlIndex < 0) return true;
                
                for (int i = contentControlIndex - 1; i >= 0; i--)
                {
                    // Skip other content controls, we're looking for the nearest heading
                    if (allElements[i] is SdtBlock) continue;
                    
                    if (allElements[i] is Paragraph p && p.ParagraphProperties?.ParagraphStyleId != null)
                    {
                        var styleId = p.ParagraphProperties.ParagraphStyleId.Val;
                        var style = mainPart.StyleDefinitionsPart.Styles
                            .Elements<Style>()
                            .FirstOrDefault(s => s.StyleId == styleId);

                        if (style != null)
                        {
                            bool isHeadingStyle = styleId.ToString().StartsWith("Heading") || 
                                (style.BasedOn != null && style.BasedOn.Val.ToString().StartsWith("Heading"));
                            
                            if (isHeadingStyle)
                            {
                                bool isCustomHeading = style.CustomStyle?.Value == true;
                                return !isCustomHeading; // false if custom heading, true otherwise
                            }
                        }
                    }
                }
                
                return true; // Default if no heading found
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error determining heading status from document, defaulting to true");
                return true;
            }
        }
    }
}