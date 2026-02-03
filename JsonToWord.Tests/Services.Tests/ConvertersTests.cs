using System;
using JsonToWord.Converters;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;
using Newtonsoft.Json;

namespace JsonToWord.Services.Tests
{
    public class ConvertersTests
    {
        [Theory]
        [InlineData("File", typeof(WordAttachment))]
        [InlineData("Html", typeof(WordHtml))]
        [InlineData("Paragraph", typeof(WordParagraph))]
        [InlineData("Picture", typeof(WordAttachment))]
        [InlineData("Table", typeof(WordTable))]
        public void WordObjectConverter_ReadJson_ReturnsExpectedType(string type, Type expectedType)
        {
            var json = $"{{\"type\":\"{type}\",\"html\":\"<p>h</p>\",\"path\":\"file.docx\",\"rows\":[]}}";
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new WordObjectConverter());

            var result = JsonConvert.DeserializeObject<IWordObject>(json, settings);

            Assert.IsType(expectedType, result);
        }

        [Fact]
        public void WordObjectConverter_CanConvert_IWordObject()
        {
            var converter = new WordObjectConverter();

            Assert.True(converter.CanConvert(typeof(IWordObject)));
        }

        [Fact]
        public void WordObjectConverter_WriteJson_SerializesWordObject()
        {
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new WordObjectConverter());

            var json = JsonConvert.SerializeObject(
                new WordAttachment { Type = WordObjectType.File, Path = "file.docx", Name = "file.docx" },
                typeof(IWordObject),
                settings);

            Assert.Contains("\"Type\"", json);
        }

        [Fact]
        public void TestReporterConverter_ReadJson_ReturnsTestReporterModel()
        {
            var json = "{\"type\":\"TestReporter\",\"testPlanName\":\"Plan A\"}";
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new TestReporterConverter());

            var result = JsonConvert.DeserializeObject<ITestReporterObject>(json, settings);

            Assert.IsType<TestReporterModel>(result);
        }

        [Fact]
        public void TestReporterConverter_ReadJson_ReturnsFlatTestReporterModel()
        {
            var json = "{\"type\":\"FlatTestReporter\",\"testPlanName\":\"Plan A\",\"rows\":[]}";
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new TestReporterConverter());

            var result = JsonConvert.DeserializeObject<ITestReporterObject>(json, settings);

            Assert.IsType<FlatTestReporterModel>(result);
        }

        [Fact]
        public void TestReporterConverter_CanConvert_ITestReporterObject()
        {
            var converter = new TestReporterConverter();

            Assert.True(converter.CanConvert(typeof(ITestReporterObject)));
        }

        [Fact]
        public void TestReporterConverter_WriteJson_SerializesTestReporterObject()
        {
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new TestReporterConverter());

            var json = JsonConvert.SerializeObject(
                new TestReporterModel { TestPlanName = "Plan A" },
                typeof(ITestReporterObject),
                settings);

            Assert.Contains("TestPlanName", json);
        }

        [Fact]
        public void TestReporterConverter_ReadJson_ReturnsNullForUnknownType()
        {
            var json = "{\"type\":\"Other\"}";
            var settings = new JsonSerializerSettings();
            settings.Converters.Add(new TestReporterConverter());

            Assert.Throws<JsonReaderException>(() => JsonConvert.DeserializeObject<ITestReporterObject>(json, settings));
        }
    }
}
