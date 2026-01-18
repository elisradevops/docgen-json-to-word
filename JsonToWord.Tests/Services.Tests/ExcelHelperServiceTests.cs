using System.Text.Json;
using JsonToWord.Services.ExcelServices;

namespace JsonToWord.Services.Tests
{
    public class ExcelHelperServiceTests
    {
        [Fact]
        public void GetValueString_HandlesJsonElementString()
        {
            var service = new ExcelHelperService();
            using var doc = JsonDocument.Parse("\"hello\"");

            var value = service.GetValueString(doc.RootElement);

            Assert.Equal("hello", value);
        }

        [Fact]
        public void GetValueString_HandlesJsonElementNonString()
        {
            var service = new ExcelHelperService();
            using var doc = JsonDocument.Parse("123");

            var value = service.GetValueString(doc.RootElement);

            Assert.Equal("123", value);
        }

        [Fact]
        public void GetValueString_ReturnsEmptyForNull()
        {
            var service = new ExcelHelperService();
            var value = service.GetValueString(null);

            Assert.Equal(string.Empty, value);
        }
    }
}
