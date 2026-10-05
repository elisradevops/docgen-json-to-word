using log4net;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace JsonToWord.Middleware
{
    /// <summary>
    /// Reads x-docgen-run-id, x-docgen-doc-type, x-docgen-project and x-docgen-capture-mode from the
    /// incoming request and stamps them into log4net's LogicalThreadContext for the duration of the
    /// request, so every log line emitted during that request carries the run correlation fields
    /// (and the diagnostics appender knows whether to persist its info/debug lines).
    /// The run id and doc type must match a strict pattern; the project (a human-readable name that
    /// may contain spaces or non-ASCII characters) is decoded, stripped of control characters and
    /// bounded; the capture mode is a whitelist. Anything else is ignored (log injection prevention).
    /// </summary>
    public class RunContextMiddleware
    {
        // Same validation as api-gate Phase 3: [A-Za-z0-9_-]{1,64}
        private static readonly Regex _validHeader = new Regex(@"^[A-Za-z0-9_-]{1,64}$", RegexOptions.Compiled);
        private readonly RequestDelegate _next;

        public RunContextMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            // Track which properties were actually set so we only remove those in finally —
            // avoids clearing a value that was never written for this request.
            var propertiesSet = new List<string>(4);
            TrySetProperty("runId", context.Request.Headers["x-docgen-run-id"].ToString(), propertiesSet);
            TrySetProperty("docType", context.Request.Headers["x-docgen-doc-type"].ToString(), propertiesSet);
            TrySetValue("project", SanitizeProject(context.Request.Headers["x-docgen-project"].ToString()), propertiesSet);
            TrySetValue("capture", ResolveCaptureMode(context.Request.Headers["x-docgen-capture-mode"].ToString()), propertiesSet);

            try
            {
                await _next(context);
            }
            finally
            {
                foreach (var name in propertiesSet)
                    LogicalThreadContext.Properties.Remove(name);
            }
        }

        private const int ProjectMaxLength = 128;

        /// <summary>
        /// A project name as it should appear in a log line: percent-decoded when it was encoded (api-gate
        /// encodes a name that cannot travel as an HTTP header), control characters removed, trimmed and
        /// bounded. Null when nothing usable is left. A name with spaces or non-ASCII characters used to be
        /// dropped entirely by the strict run-id pattern.
        /// </summary>
        public static string SanitizeProject(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            var text = value;
            try
            {
                text = Uri.UnescapeDataString(value);
            }
            catch (UriFormatException)
            {
                // not percent-encoded (or malformed): use as sent
            }
            var cleaned = new string(text.Where(c => !char.IsControl(c)).ToArray()).Trim();
            if (cleaned.Length > ProjectMaxLength) cleaned = cleaned.Substring(0, ProjectMaxLength);
            return cleaned.Length == 0 ? null : cleaned;
        }

        /// <summary>The capture mode api-gate forwarded, if it is one we know; null (normal) otherwise.</summary>
        public static string ResolveCaptureMode(string value)
        {
            var mode = (value ?? string.Empty).Trim();
            return mode == "verbose" || mode == "retain-on-failure" ? mode : null;
        }

        private static void TrySetValue(string name, string value, List<string> propertiesSet)
        {
            if (!string.IsNullOrEmpty(value))
            {
                LogicalThreadContext.Properties[name] = value;
                propertiesSet.Add(name);
            }
        }

        private static void TrySetProperty(string name, string value, List<string> propertiesSet)
        {
            if (!string.IsNullOrWhiteSpace(value) && _validHeader.IsMatch(value))
            {
                LogicalThreadContext.Properties[name] = value;
                propertiesSet.Add(name);
            }
        }
    }
}
