using JsonToWord.Middleware;
using log4net;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace JsonToWord.Middleware.Tests
{
    // IDisposable gives per-test setup/teardown in xUnit.
    // We clear the three known LogicalThreadContext keys before and after each test
    // to prevent static state from one test bleeding into another.
    public class RunContextMiddlewareTests : IDisposable
    {
        public RunContextMiddlewareTests() => ClearProperties();
        public void Dispose() => ClearProperties();

        private static void ClearProperties()
        {
            LogicalThreadContext.Properties.Remove("runId");
            LogicalThreadContext.Properties.Remove("docType");
            LogicalThreadContext.Properties.Remove("project");
            LogicalThreadContext.Properties.Remove("capture");
        }

        [Fact]
        public async Task ValidRunId_SetsLogicalThreadContextProperty()
        {
            string? capturedRunId = null;

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedRunId = LogicalThreadContext.Properties["runId"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "abc-123";

            await middleware.InvokeAsync(context);

            Assert.Equal("abc-123", capturedRunId);
        }

        [Fact]
        public async Task ValidRunId_FlowsAcrossAsyncContinuation()
        {
            // Verifies that LogicalThreadContext (AsyncLocal-backed in log4net 2.0.x on .NET 6)
            // correctly propagates across a thread-pool context switch.
            string? capturedRunId = null;

            var middleware = new RunContextMiddleware(async ctx =>
            {
                await Task.Yield(); // forces continuation on a different thread-pool thread
                capturedRunId = LogicalThreadContext.Properties["runId"]?.ToString();
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "async-flow-test";

            await middleware.InvokeAsync(context);

            Assert.Equal("async-flow-test", capturedRunId);
        }

        [Fact]
        public async Task ValidDocTypeAndProject_SetProperties()
        {
            string? capturedDocType = null;
            string? capturedProject = null;

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedDocType = LogicalThreadContext.Properties["docType"]?.ToString();
                capturedProject = LogicalThreadContext.Properties["project"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-doc-type"] = "SVD";
            context.Request.Headers["x-docgen-project"] = "MyProject";

            await middleware.InvokeAsync(context);

            Assert.Equal("SVD", capturedDocType);
            Assert.Equal("MyProject", capturedProject);
        }

        [Fact]
        public async Task MissingHeader_DoesNotSetProperty()
        {
            string capturedRunId = "sentinel";

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedRunId = LogicalThreadContext.Properties["runId"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext(); // no x-docgen-run-id header

            await middleware.InvokeAsync(context);

            Assert.Null(capturedRunId);
        }

        [Fact]
        public async Task InvalidHeader_TooLong_DoesNotSetProperty()
        {
            string capturedRunId = "sentinel";
            var tooLong = new string('a', 65); // 65 chars — exceeds 64-char limit

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedRunId = LogicalThreadContext.Properties["runId"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = tooLong;

            await middleware.InvokeAsync(context);

            Assert.Null(capturedRunId);
        }

        [Fact]
        public async Task InvalidHeader_BadChars_DoesNotSetProperty()
        {
            string capturedRunId = "sentinel";

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedRunId = LogicalThreadContext.Properties["runId"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "bad value!"; // space and ! are invalid

            await middleware.InvokeAsync(context);

            Assert.Null(capturedRunId);
        }

        [Fact]
        public async Task PropertiesAreCleared_AfterRequestCompletes()
        {
            var middleware = new RunContextMiddleware(ctx => Task.CompletedTask);

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "run-cleanup-test";
            context.Request.Headers["x-docgen-doc-type"] = "SVD";
            context.Request.Headers["x-docgen-project"] = "Proj1";

            await middleware.InvokeAsync(context);

            Assert.Null(LogicalThreadContext.Properties["runId"]?.ToString());
            Assert.Null(LogicalThreadContext.Properties["docType"]?.ToString());
            Assert.Null(LogicalThreadContext.Properties["project"]?.ToString());
        }

        [Fact]
        public async Task MissingHeader_NotRemovedFromContext_LeavingUnrelatedValuesIntact()
        {
            // Verifies that only properties that were actually set are removed in finally —
            // a property set externally for "project" must not be wiped when this request
            // did not supply x-docgen-project.
            LogicalThreadContext.Properties["project"] = "external-value";

            string capturedProject = null;

            var middleware = new RunContextMiddleware(ctx =>
            {
                capturedProject = LogicalThreadContext.Properties["project"]?.ToString();
                return Task.CompletedTask;
            });

            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "only-run-id";
            // no x-docgen-project header

            await middleware.InvokeAsync(context);

            // The middleware should not have removed "project" since it did not set it
            Assert.Equal("external-value", capturedProject);
            Assert.Equal("external-value", LogicalThreadContext.Properties["project"]?.ToString());
        }
            [Theory]
        [InlineData("Cube ADCS", "Cube ADCS")]
        [InlineData("  MEWP  ", "MEWP")]
        [InlineData("TestProject-CMMI", "TestProject-CMMI")]
        public void SanitizeProject_KeepsNamesWithSpacesAndPunctuation(string input, string expected)
        {
            Assert.Equal(expected, RunContextMiddleware.SanitizeProject(input));
        }

        [Fact]
        public void SanitizeProject_DecodesPercentEncodedNonAsciiNames()
        {
            var encoded = Uri.EscapeDataString("פרויקט MEWP");
            Assert.Equal("פרויקט MEWP", RunContextMiddleware.SanitizeProject(encoded));
        }

        [Fact]
        public void SanitizeProject_StripsControlCharactersAndBoundsLength()
        {
            Assert.Equal("MEWPFAKE: line", RunContextMiddleware.SanitizeProject("MEWP\r\nFAKE: line"));
            Assert.Equal(128, RunContextMiddleware.SanitizeProject(new string('x', 500))!.Length);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("\r\n")]
        public void SanitizeProject_ReturnsNullWhenNothingUsableIsLeft(string input)
        {
            Assert.Null(RunContextMiddleware.SanitizeProject(input));
        }

        [Fact]
        public void SanitizeProject_UsesAMalformedEscapeAsSent()
        {
            Assert.Equal("100%", RunContextMiddleware.SanitizeProject("100%"));
        }

        [Fact]
        public async Task ProjectWithSpaces_IsStampedOnTheLogContext()
        {
            string? captured = null;
            var middleware = new RunContextMiddleware(ctx =>
            {
                captured = LogicalThreadContext.Properties["project"]?.ToString();
                return Task.CompletedTask;
            });
            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-run-id"] = "abc-123";
            context.Request.Headers["x-docgen-project"] = "Cube ADCS";

            await middleware.InvokeAsync(context);

            Assert.Equal("Cube ADCS", captured);
        }

        [Theory]
        [InlineData("verbose", "verbose")]
        [InlineData("retain-on-failure", "retain-on-failure")]
        [InlineData("  verbose ", "verbose")]
        [InlineData("normal", null)]
        [InlineData("VERBOSE", null)]
        [InlineData("DROP TABLE runs", null)]
        [InlineData("", null)]
        [InlineData(null, null)]
        public void ResolveCaptureMode_IsAWhitelist(string input, string? expected)
        {
            Assert.Equal(expected, RunContextMiddleware.ResolveCaptureMode(input));
        }

        [Fact]
        public async Task CaptureMode_IsStampedForTheRequestAndCleanedUpAfterwards()
        {
            string? during = null;
            var middleware = new RunContextMiddleware(ctx =>
            {
                during = LogicalThreadContext.Properties["capture"]?.ToString();
                return Task.CompletedTask;
            });
            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-capture-mode"] = "verbose";

            await middleware.InvokeAsync(context);

            Assert.Equal("verbose", during);
            Assert.Null(LogicalThreadContext.Properties["capture"]);
        }

        [Fact]
        public async Task UnknownCaptureMode_IsNotStamped()
        {
            string? during = "unset";
            var middleware = new RunContextMiddleware(ctx =>
            {
                during = LogicalThreadContext.Properties["capture"]?.ToString();
                return Task.CompletedTask;
            });
            var context = new DefaultHttpContext();
            context.Request.Headers["x-docgen-capture-mode"] = "everything";

            await middleware.InvokeAsync(context);

            Assert.Null(during);
        }
    }
}
