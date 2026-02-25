using System;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonToWord.Converters
{
    public class TestReporterConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            serializer.Serialize(writer, value, typeof(ITestReporterObject));
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            var property = "type";
            var jo = JObject.Load(reader);

            if (string.Equals(jo[property].Value<string>(), TestReporterObjectType.TestReporter.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<TestReporterModel>(serializer);
            if (string.Equals(jo[property].Value<string>(), TestReporterObjectType.FlatTestReporter.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<FlatTestReporterModel>(serializer);
            if (string.Equals(jo[property].Value<string>(), TestReporterObjectType.MewpCoverageReporter.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<MewpCoverageReporterModel>(serializer);
            if (string.Equals(jo[property].Value<string>(), TestReporterObjectType.InternalValidationReporter.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<InternalValidationReporterModel>(serializer);

            return serializer.Deserialize(reader, typeof(ITestReporterObject));
        }

        public override bool CanConvert(Type objectType)
        {
            var result = (objectType == typeof(ITestReporterObject));
            return result;
        }
    }
}
