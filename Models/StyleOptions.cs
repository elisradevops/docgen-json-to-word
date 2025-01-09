namespace JsonToWord.Models
{
    public class StyleOptions
    {
        public string Font { get; set; }
        public bool InsertLineBreak { get; set; }
        public bool InsertSpace { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public int Size { get; set; }
        public string Uri { get; set; }
        public string FontColor { get; set; }
        public StyleOptions() {
            Font = "Arial";
            Size = 12;
        }
    }
}
