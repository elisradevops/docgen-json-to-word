using JsonToWord.Models.S3;
using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class ExcelZipPackageModel
    {
        public UploadProperties UploadProperties { get; set; }
        public List<DownloadableObjectModel> Files { get; set; }
    }
}
