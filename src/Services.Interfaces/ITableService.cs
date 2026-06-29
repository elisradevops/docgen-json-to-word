using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using System.Collections.Generic;

namespace JsonToWord.Services.Interfaces
{
    public interface ITableService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordTable wordTable, FormattingSettings formattingSettings);
        void SetSectionBookmarks(HashSet<string> bookmarks);
    }
}
