using JsonToWord.Logging;

namespace JsonToWord.Logging.Tests
{
    // The persistence policy: warn/error always; info/debug only for a run that asked for extra capture.
    public class DiagnosticsHttpAppenderTests
    {
        [Theory]
        [InlineData("warn", null)]
        [InlineData("error", null)]
        [InlineData("warn", "verbose")]
        [InlineData("error", "retain-on-failure")]
        public void WarnAndError_AreAlwaysPersisted(string level, string? capture)
        {
            Assert.True(DiagnosticsHttpAppender.ShouldPersist(level, capture));
        }

        [Theory]
        [InlineData("info")]
        [InlineData("debug")]
        public void InfoAndDebug_AreNotPersistedInNormalMode(string level)
        {
            Assert.False(DiagnosticsHttpAppender.ShouldPersist(level, null));
            Assert.False(DiagnosticsHttpAppender.ShouldPersist(level, ""));
        }

        [Theory]
        [InlineData("info", "verbose")]
        [InlineData("debug", "verbose")]
        [InlineData("info", "retain-on-failure")]
        [InlineData("debug", "retain-on-failure")]
        public void InfoAndDebug_ArePersistedWhenTheRunAskedForCapture(string level, string capture)
        {
            Assert.True(DiagnosticsHttpAppender.ShouldPersist(level, capture));
        }

        [Fact]
        public void AnUnknownCaptureValueDoesNotOpenTheGate()
        {
            Assert.False(DiagnosticsHttpAppender.ShouldPersist("info", "everything"));
            Assert.False(DiagnosticsHttpAppender.ShouldPersist("trace", "verbose"));
        }
    }
}
