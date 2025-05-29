using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class TestReporterContentControl
    {
        public bool ForceClean { get; set; }
        public string Title { get; set; }
        public List<ITestReporterObject> WordObjects { get; set; }
        public bool AllowGrouping { get; set; } = false;
    }
}