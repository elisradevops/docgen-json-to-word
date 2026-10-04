using log4net;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace JsonToWord.Middleware
{
    /// <summary>
    /// Reads x-docgen-run-id, x-docgen-doc-type, and x-docgen-project from the incoming request
    /// and stamps them into log4net's LogicalThreadContext for the duration of the request, so
    /// every log line emitted during that request carries the run correlation fields.
    /// Values that do not match the allowed pattern are ignored (log injection prevention).
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
            var propertiesSet = new List<string>(3);
            TrySetProperty("runId", context.Request.Headers["x-docgen-run-id"].ToString(), propertiesSet);
            TrySetProperty("docType", context.Request.Headers["x-docgen-doc-type"].ToString(), propertiesSet);
            TrySetProperty("project", context.Request.Headers["x-docgen-project"].ToString(), propertiesSet);

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
