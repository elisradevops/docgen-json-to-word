using System.Collections.Generic;
using Newtonsoft.Json;

namespace JsonToWord.Models.TestReporterModels
{
    public class AssociatedItemModel
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string WorkItemType { get; set; }
        public string Url { get; set; }
        [JsonExtensionData]
        public Dictionary<string, object> CustomFields { get; set; }
    }
}