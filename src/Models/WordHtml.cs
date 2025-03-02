namespace JsonToWord.Models
{
    public class WordHtml : IWordObject
    {
        public string Html { get; set; }
        public WordObjectType Type { get; set; }
        public string Font {  get; set; }   
        public uint FontSize { get; set; }
    }
}