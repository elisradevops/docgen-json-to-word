using log4net;
using log4net.Appender;
using log4net.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace JsonToWord.Logging
{
    /// <summary>
    /// Buffers log events and flushes them in batches to api-gate's
    /// POST /diagnostics/logs endpoint so json-to-word events appear in the
    /// LogsExplorer dashboard, the same way dg-content-control does via its
    /// Node HttpLogSink. Guards every outbound call so that a failed flush
    /// never propagates into the application call path.
    /// Activated only when DIAGNOSTICS_INGEST_URL and DIAGNOSTICS_INGEST_TOKEN
    /// are both set — otherwise DoAppend is a no-op.
    /// </summary>
    public sealed class DiagnosticsHttpAppender : AppenderSkeleton
    {
        private static readonly HttpClient _http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        private static readonly string _version = Assembly
            .GetExecutingAssembly()
            .GetName()
            .Version?.ToString() ?? "unknown";

        private const int FlushIntervalMs = 2000;
        private const int FlushBatchSize = 200;
        private const int BufferMax = 5000;

        private readonly List<object> _buffer = new List<object>(FlushBatchSize);
        private readonly object _lock = new object();
        private readonly Timer _timer;

        public DiagnosticsHttpAppender()
        {
            _timer = new Timer(_ => Flush(), null, FlushIntervalMs, FlushIntervalMs);
        }

        /// <summary>
        /// Same policy as the other DocGen services: warn and error are always persisted; info and debug
        /// only for a run that asked for extra capture (verbose, or retain-on-failure). Without this every
        /// document run stored ~25 internal "Removing content control…" lines. The console appender is
        /// unaffected: stdout still shows everything.
        /// </summary>
        public static bool ShouldPersist(string level, string captureMode)
        {
            if (level == "warn" || level == "error") return true;
            return (level == "info" || level == "debug") && (captureMode == "verbose" || captureMode == "retain-on-failure");
        }

        protected override void Append(LoggingEvent loggingEvent)
        {
            var ingestUrl = Environment.GetEnvironmentVariable("DIAGNOSTICS_INGEST_URL");
            var ingestToken = Environment.GetEnvironmentVariable("DIAGNOSTICS_INGEST_TOKEN");
            if (string.IsNullOrWhiteSpace(ingestUrl) || string.IsNullOrWhiteSpace(ingestToken))
                return;

            var level = loggingEvent.Level.Name.ToLowerInvariant() switch
            {
                "fatal" => "error",
                "error" => "error",
                "warn" => "warn",
                "info" => "info",
                "debug" => "debug",
                _ => null
            };
            if (level == null) return;

            var capture = LogicalThreadContext.Properties["capture"]?.ToString();
            if (!ShouldPersist(level, capture)) return;

            var runId = LogicalThreadContext.Properties["runId"]?.ToString();
            var docType = LogicalThreadContext.Properties["docType"]?.ToString();
            var project = LogicalThreadContext.Properties["project"]?.ToString();

            object errObj = null;
            if (loggingEvent.ExceptionObject != null)
            {
                var ex = loggingEvent.ExceptionObject;
                errObj = new
                {
                    message = ex.Message?.Length > 2000 ? ex.Message.Substring(0, 2000) : ex.Message,
                    stack = ex.StackTrace?.Length > 4000 ? ex.StackTrace.Substring(0, 4000) : ex.StackTrace,
                };
            }

            var msg = loggingEvent.RenderedMessage ?? string.Empty;
            var evt = new
            {
                ts = loggingEvent.TimeStamp.ToUniversalTime().ToString("o"),
                level,
                service = "json-to-word",
                version = _version,
                runId,
                docType,
                project,
                // Everything this service logs for a request happens while rendering the document.
                step = "render-document",
                // A provisional line of a retain-on-failure run: kept only if the run fails.
                retainPending = (level == "info" || level == "debug") && capture == "retain-on-failure" ? (bool?)true : null,
                message = msg.Length > 2000 ? msg.Substring(0, 2000) : msg,
                err = errObj,
            };

            lock (_lock)
            {
                if (_buffer.Count >= BufferMax)
                    _buffer.RemoveAt(0);
                _buffer.Add(evt);
                if (_buffer.Count >= FlushBatchSize)
                    ThreadPool.QueueUserWorkItem(_ => Flush());
            }
        }

        private void Flush()
        {
            var ingestUrl = Environment.GetEnvironmentVariable("DIAGNOSTICS_INGEST_URL");
            var ingestToken = Environment.GetEnvironmentVariable("DIAGNOSTICS_INGEST_TOKEN");
            if (string.IsNullOrWhiteSpace(ingestUrl) || string.IsNullOrWhiteSpace(ingestToken))
                return;

            List<object> batch;
            lock (_lock)
            {
                if (_buffer.Count == 0) return;
                batch = new List<object>(_buffer);
                _buffer.Clear();
            }

            try
            {
                PostWithOneRetry(ingestUrl, ingestToken, batch);
            }
            catch (Exception ex)
            {
                // Never re-log via log4net — that would recurse back into this appender.
                // Drop the batch silently; losing dashboard data must never break generation.
                Console.Error.WriteLine($"DiagnosticsHttpAppender: flush failed: {ex.Message}");
            }
        }

        private static void PostWithOneRetry(string ingestUrl, string ingestToken, List<object> batch)
        {
            try
            {
                Post(ingestUrl, ingestToken, batch);
            }
            catch
            {
                Post(ingestUrl, ingestToken, batch);
            }
        }

        private static void Post(string ingestUrl, string ingestToken, List<object> batch)
        {
            var payload = JsonConvert.SerializeObject(new { events = batch });
            // Task.Run gives the async call a fresh threadpool context with no sync context,
            // avoiding GetAwaiter().GetResult() deadlocks on timer threads.
            Task.Run(async () =>
            {
                using var request = new HttpRequestMessage(HttpMethod.Post, $"{ingestUrl}/diagnostics/logs")
                {
                    Content = new StringContent(payload, Encoding.UTF8, "application/json")
                };
                request.Headers.Add("x-docgen-ingest-token", ingestToken);
                var response = await _http.SendAsync(request).ConfigureAwait(false);
                response.EnsureSuccessStatusCode();
            }).GetAwaiter().GetResult();
        }

        protected override void OnClose()
        {
            _timer.Dispose();
            Flush();
            base.OnClose();
        }
    }
}
