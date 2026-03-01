using JsonToWord.Controllers;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Globalization;

namespace JsonToWord.Controllers.Tests
{
    public class HealthControllerTests
    {
        [Fact]
        public void GetHealth_ReturnsExpectedPayload()
        {
            var controller = new HealthController();

            var result = controller.GetHealth();

            var ok = Assert.IsType<OkObjectResult>(result);
            var json = JsonConvert.SerializeObject(ok.Value);

            Assert.Contains("\"service\":\"json-to-word\"", json);
            Assert.Contains("\"status\":\"up\"", json);
            Assert.Contains("\"connectionStatus\":\"connected\"", json);
            Assert.Contains("\"version\":\"", json);
            Assert.Contains("\"timestamp\":\"", json);
        }

        [Fact]
        public void GetHealth_ReturnsIsoUtcTimestampAndNonEmptyVersion()
        {
            var controller = new HealthController();

            var result = controller.GetHealth();

            var ok = Assert.IsType<OkObjectResult>(result);
            var json = JsonConvert.SerializeObject(ok.Value);
            var payload = JObject.Parse(json);

            var version = payload["version"]?.Value<string>();
            Assert.False(string.IsNullOrWhiteSpace(version));

            var timestampText = payload["timestamp"]?.Value<string>();
            Assert.False(string.IsNullOrWhiteSpace(timestampText));
            Assert.True(
                DateTime.TryParse(
                    timestampText,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out var parsedTimestamp
                )
            );
            Assert.True(parsedTimestamp > DateTime.MinValue);
        }
    }
}
