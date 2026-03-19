using DocumentFormat.OpenXml.Packaging;

namespace JsonToWord.Services.Interfaces
{
    public interface ISectionPlaceholderService
    {
        /// <summary>
        /// Walks the document body, tracks heading counters, and resolves
        /// {{section:X.Y}} placeholders in table cells by prepending the
        /// parent heading number. Also supports anchored placeholders:
        /// {{section:anchor-name:X.Y}} resolved from
        /// {{section-anchor:anchor-name}} marker position.
        /// </summary>
        void ResolveSectionPlaceholders(WordprocessingDocument document);
    }
}
