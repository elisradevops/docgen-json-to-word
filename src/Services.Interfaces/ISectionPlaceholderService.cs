using DocumentFormat.OpenXml.Packaging;

namespace JsonToWord.Services.Interfaces
{
    public interface ISectionPlaceholderService
    {
        /// <summary>
        /// Walks the document body, tracks heading counters, and resolves
        /// {{section:X.Y}} placeholders in table cells by prepending the
        /// parent heading number.
        /// </summary>
        void ResolveSectionPlaceholders(WordprocessingDocument document);
    }
}
