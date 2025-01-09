using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordList : IWordObject
    {
        public WordObjectType Type { get; set; }
        public List<WordListItem> ListItems { get; set; }
        public bool IsOrdered { get; set; }
    }
}
