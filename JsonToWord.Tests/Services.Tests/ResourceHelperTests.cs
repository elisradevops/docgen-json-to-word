using System.Resources;
using JsonToWord;

namespace JsonToWord.Services.Tests
{
    public class ResourceHelperTests
    {
        [Fact]
        public void GetStream_ReturnsEmbeddedResource()
        {
            using var stream = ResourceHelper.GetStream("Resources.template.docx");

            Assert.NotNull(stream);
            Assert.True(stream.CanRead);
        }

        [Fact]
        public void GetString_ReturnsResourceContent()
        {
            var content = ResourceHelper.GetString(typeof(WordService).Assembly, "Resources.template.docx");

            Assert.NotNull(content);
        }

        [Fact]
        public void GetStream_ThrowsWhenResourceMissing()
        {
            Assert.Throws<MissingManifestResourceException>(() => ResourceHelper.GetStream("missing.resource"));
        }
    }
}
