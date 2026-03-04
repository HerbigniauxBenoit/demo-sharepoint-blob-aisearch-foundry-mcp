using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Services;

namespace SharePointSync.Functions;

public sealed class HealthCheckFunction
{
    private readonly HealthCheckService _healthCheckService;
    private readonly ILogger<HealthCheckFunction> _logger;
    private readonly IdentityService _identityService;
    private readonly IConfiguration _configuration;

    public HealthCheckFunction(
        HealthCheckService healthCheckService, 
        ILogger<HealthCheckFunction> logger,
        IdentityService identityService,
        IConfiguration configuration)
    {
        _healthCheckService = healthCheckService;
        _logger = logger;
        _identityService = identityService;
        _configuration = configuration;
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

    [Function("IdentityCheck")]
    public async Task<HttpResponseData> CheckIdentityAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "health/identity")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("Identity check requested");

        try
        {
            // Log identity details to console/logs
            await _identityService.LogIdentityDetailsAsync(cancellationToken);

            // Validate SharePoint access
            var sharePointSiteUrl = _configuration["SHAREPOINT_SITE_URL"];
            string? sharePointValidation = null;
            bool? sharePointAccess = null;

            if (!string.IsNullOrEmpty(sharePointSiteUrl))
            {
                var (success, message) = await _identityService.ValidateSharePointAccessAsync(sharePointSiteUrl, cancellationToken);
                sharePointAccess = success;
                sharePointValidation = message;
            }

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json");

            var result = new
            {
                status = "healthy",
                identity = new
                {
                    type = !string.IsNullOrEmpty(_configuration["AZURE_CLIENT_ID"]) ? "UserAssigned" : "SystemAssigned",
                    clientId = _configuration["AZURE_CLIENT_ID"],
                    tenantId = _configuration["AZURE_TENANT_ID"]
                },
                sharePoint = new
                {
                    siteUrl = sharePointSiteUrl,
                    accessGranted = sharePointAccess,
                    validationMessage = sharePointValidation
                },
                message = "Check the logs for detailed identity information including roles and permissions"
            };

            await response.WriteAsJsonAsync(result, cancellationToken);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Identity check failed");

            var response = req.CreateResponse(HttpStatusCode.InternalServerError);
            response.Headers.Add("Content-Type", "application/json");

            await response.WriteAsJsonAsync(new
            {
                status = "error",
                error = ex.Message,
                stackTrace = ex.StackTrace
            }, cancellationToken);

            return response;
        }
    }
}
