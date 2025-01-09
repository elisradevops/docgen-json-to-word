namespace JsonToWord.Models
{
    public class WordRun
    {
        public string Type { get; set; }
        public string Value { get; set; }
        public string? Src { get; set; }
        public StyleOptions TextStyling { get; set; }


        public WordRun()
        {
           TextStyling = new StyleOptions();
        }
    }
}