using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using Microsoft.Extensions.Logging;

namespace SharePointSync.Functions;

public sealed class HealthCheckFunction
{
    private readonly HealthCheckService _healthCheckService;
    private readonly ILogger<HealthCheckFunction> _logger;

    public HealthCheckFunction(
        HealthCheckService healthCheckService, 
        ILogger<HealthCheckFunction> logger)
    {
        _healthCheckService = healthCheckService;
        _logger = logger;
    }

    [Function("HealthCheck")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "health")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("Health check requested");

        var healthReport = await _healthCheckService.CheckHealthAsync(cancellationToken);
        
        var response = req.CreateResponse(
            healthReport.Status == HealthStatus.Healthy ? HttpStatusCode.OK : HttpStatusCode.ServiceUnavailable);
        
        response.Headers.Add("Content-Type", "application/json");
        
        var result = new
        {
            status = healthReport.Status.ToString(),
            checks = healthReport.Entries.Select(e => new
            {
                name = e.Key,
                status = e.Value.Status.ToString(),
                description = e.Value.Description,
                duration = e.Value.Duration.TotalMilliseconds
            }),
            totalDuration = healthReport.TotalDuration.TotalMilliseconds
        };
        
        await response.WriteAsJsonAsync(result, cancellationToken);
        return response;
    }
}
