using JsonToWord.Controllers;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

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
    }
}
