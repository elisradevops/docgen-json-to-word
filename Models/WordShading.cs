namespace JsonToWord.Models
{
    public class WordShading
    {
        public string Color { get; set; }
        public string Fill { get; set; }

        public WordShading()
        {
            Color = "auto";
        }
    }
}