using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordListItem
    {
        public List<WordRun> Runs { get; set; }
        public int Level { get; set; }
    }
}
