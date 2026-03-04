using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
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
    private readonly TokenCredential _tokenCredential;
    private readonly IHttpClientFactory _httpClientFactory;

    public HealthCheckFunction(
        HealthCheckService healthCheckService, 
        ILogger<HealthCheckFunction> logger,
        IdentityService identityService,
        IConfiguration configuration,
        TokenCredential tokenCredential,
        IHttpClientFactory httpClientFactory)
    {
        _healthCheckService = healthCheckService;
        _logger = logger;
        _identityService = identityService;
        _configuration = configuration;
        _tokenCredential = tokenCredential;
        _httpClientFactory = httpClientFactory;
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
        _logger.LogInformation("========== IDENTITY CHECK START ==========");

        try
        {
            var sharePointSiteUrl = _configuration["SHAREPOINT_SITE_URL"];
            string? sharePointValidation = null;
            bool? sharePointAccess = null;

            _logger.LogInformation("Identity type configured: {IdentityType}",
                !string.IsNullOrEmpty(_configuration["AZURE_CLIENT_ID"]) ? "UserAssigned" : "SystemAssigned");
            _logger.LogInformation("Configured AZURE_CLIENT_ID: {ClientId}", _configuration["AZURE_CLIENT_ID"] ?? "(not set)");
            _logger.LogInformation("Configured AZURE_TENANT_ID: {TenantId}", _configuration["AZURE_TENANT_ID"] ?? "(not set)");
            _logger.LogInformation("Configured SHAREPOINT_SITE_URL: {SiteUrl}", sharePointSiteUrl ?? "(not set)");

            _logger.LogInformation("Step 1/3 - Listing sites accessible by current token...");
            var (sitesListSuccess, sitesListMessage, accessibleSites) = await TryListAccessibleSitesAsync(cancellationToken);
            _logger.LogInformation("Step 1/3 result - Success: {Success}, Message: {Message}, Returned sites: {Count}",
                sitesListSuccess,
                sitesListMessage,
                accessibleSites.Count);

            _logger.LogInformation("Step 2/3 - Logging managed identity and token details...");
            await _identityService.LogIdentityDetailsAsync(cancellationToken);
            _logger.LogInformation("Step 2/3 completed.");

            _logger.LogInformation("Step 3/3 - Validating access to configured SharePoint site...");
            if (!string.IsNullOrEmpty(sharePointSiteUrl))
            {
                await _identityService.LogCandidateSitesForTargetAsync(sharePointSiteUrl, cancellationToken);

                var (success, message) = await _identityService.ValidateSharePointAccessAsync(sharePointSiteUrl, cancellationToken);
                sharePointAccess = success;
                sharePointValidation = message;
                _logger.LogInformation("Step 3/3 result - AccessGranted: {AccessGranted}, Message: {Message}", success, message);
            }
            else
            {
                _logger.LogWarning("Step 3/3 skipped - SHAREPOINT_SITE_URL is not configured.");
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
                sitesDiscovery = new
                {
                    success = sitesListSuccess,
                    message = sitesListMessage,
                    items = accessibleSites
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
            _logger.LogInformation("========== IDENTITY CHECK END (SUCCESS) ==========");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "========== IDENTITY CHECK END (ERROR) ==========");

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

    [Function("GraphCheck")]
    public async Task<HttpResponseData> CheckGraphAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "health/graph")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("========== GRAPH CHECK START ==========");

        try
        {
            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var token = await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);

            using var client = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/sites/root");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            using var response = await client.SendAsync(request, cancellationToken);
            var payload = await response.Content.ReadAsStringAsync(cancellationToken);

            var result = new
            {
                statusCode = (int)response.StatusCode,
                reasonPhrase = response.ReasonPhrase,
                success = response.IsSuccessStatusCode,
                endpoint = "GET /v1.0/sites/root",
                responseBody = payload
            };

            var httpResponse = req.CreateResponse(response.IsSuccessStatusCode ? HttpStatusCode.OK : HttpStatusCode.BadGateway);
            await httpResponse.WriteAsJsonAsync(result, cancellationToken);

            _logger.LogInformation("========== GRAPH CHECK END (Status={Status}) ==========", response.StatusCode);
            return httpResponse;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "========== GRAPH CHECK END (ERROR) ==========");

            var response = req.CreateResponse(HttpStatusCode.InternalServerError);
            await response.WriteAsJsonAsync(new
            {
                status = "error",
                error = ex.Message
            }, cancellationToken);

            return response;
        }
    }

    private async Task<(bool Success, string Message, IReadOnlyList<object> Sites)> TryListAccessibleSitesAsync(CancellationToken cancellationToken)
    {
        try
        {
            _logger.LogInformation("Trying Graph sites listing without role pre-check...");

            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var token = await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);

            using var client = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/sites?search=*");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            using var response = await client.SendAsync(request, cancellationToken);
            var payload = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                var message = $"Unable to list sites. Status: {response.StatusCode}";
                _logger.LogWarning("{Message}. Body: {Body}", message, payload);
                return (false, message, Array.Empty<object>());
            }

            using var doc = JsonDocument.Parse(payload);
            if (!doc.RootElement.TryGetProperty("value", out var sitesElement) || sitesElement.ValueKind != JsonValueKind.Array)
            {
                const string message = "No site list returned by Graph.";
                _logger.LogWarning(message);
                return (true, message, Array.Empty<object>());
            }

            var sites = sitesElement.EnumerateArray()
                .Select(site => new
                {
                    id = site.TryGetProperty("id", out var idEl) ? idEl.GetString() : null,
                    displayName = site.TryGetProperty("displayName", out var dnEl) ? dnEl.GetString() : null,
                    webUrl = site.TryGetProperty("webUrl", out var wuEl) ? wuEl.GetString() : null
                })
                .Where(s => !string.IsNullOrWhiteSpace(s.webUrl))
                .Take(30)
                .Cast<object>()
                .ToList();

            _logger.LogInformation("Sites listing completed successfully. Count: {Count}", sites.Count);
            return (true, $"Listed {sites.Count} site(s).", sites);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not list accessible sites for identity check.");
            return (false, $"Exception while listing sites: {ex.Message}", Array.Empty<object>());
        }
    }
}
