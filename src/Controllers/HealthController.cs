using Microsoft.AspNetCore.Mvc;
using System;
using System.Diagnostics;
using System.Reflection;

namespace JsonToWord.Controllers
{
    [Route("health")]
    [ApiController]
    public class HealthController : ControllerBase
    {
        [HttpGet]
        public IActionResult GetHealth()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var fileVersion = FileVersionInfo.GetVersionInfo(assembly.Location)?.FileVersion;
            var infoVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            var assemblyVersion = assembly.GetName().Version?.ToString();
            var resolvedVersion = infoVersion ?? fileVersion ?? assemblyVersion ?? "unknown";

            return Ok(new
            {
                service = "json-to-word",
                status = "up",
                connectionStatus = "connected",
                version = resolvedVersion,
                timestamp = DateTime.UtcNow.ToString("O"),
            });
        }
    }
}
