using System.Net;
using System.IdentityModel.Tokens.Jwt;
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
        _logger.LogInformation("Identity check requested");

        try
        {
            var sharePointSiteUrl = _configuration["SHAREPOINT_SITE_URL"];
            string? sharePointValidation = null;
            bool? sharePointAccess = null;

            var (sitesListSuccess, sitesListMessage, accessibleSites) = await TryListAccessibleSitesAsync(cancellationToken);

            // 1) First step: list candidate sites in logs
            if (!string.IsNullOrEmpty(sharePointSiteUrl))
            {
                await _identityService.LogCandidateSitesForTargetAsync(sharePointSiteUrl, cancellationToken);
            }

            // 2) Then log identity details
            await _identityService.LogIdentityDetailsAsync(cancellationToken);

            // 3) Then validate access to configured target site
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

    private async Task<(bool Success, string Message, IReadOnlyList<object> Sites)> TryListAccessibleSitesAsync(CancellationToken cancellationToken)
    {
        try
        {
            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var token = await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);
            var jwtToken = new JwtSecurityTokenHandler().ReadJwtToken(token.Token);

            var roles = jwtToken.Claims
                .Where(c => c.Type == "roles")
                .Select(c => c.Value)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var canListSites = roles.Contains("Sites.Read.All", StringComparer.OrdinalIgnoreCase)
                               || roles.Contains("Sites.ReadWrite.All", StringComparer.OrdinalIgnoreCase);

            if (!canListSites)
            {
                const string message = "Listing all sites requires Sites.Read.All or Sites.ReadWrite.All. With Sites.Selected only, global listing is not available.";
                return (false, message, Array.Empty<object>());
            }

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

            return (true, $"Listed {sites.Count} site(s).", sites);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not list accessible sites for identity check.");
            return (false, $"Exception while listing sites: {ex.Message}", Array.Empty<object>());
        }
    }
}
