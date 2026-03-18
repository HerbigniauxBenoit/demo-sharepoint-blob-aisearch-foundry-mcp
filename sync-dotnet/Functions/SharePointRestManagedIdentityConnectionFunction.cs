using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace SharePointSync.Functions;

public sealed class SharePointRestManagedIdentityConnectionFunction
{
    private readonly ILogger<SharePointRestManagedIdentityConnectionFunction> _logger;
    private readonly IHttpClientFactory _httpClientFactory;

    public SharePointRestManagedIdentityConnectionFunction(
        ILogger<SharePointRestManagedIdentityConnectionFunction> logger,
        IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    [Function("SharePointRestManagedIdentityConnection")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sharepoint/msi-rest-connection")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        try
        {
            _logger.LogInformation("SharePointRestManagedIdentityConnection started: {Method} /api/sharepoint/msi-rest-connection", req.Method);
            var settings = ReadAndValidateSettings();

            LogStepStart(2, "Token MSI pour SharePoint REST");
            var sharePointScope = $"https://{settings.SharePointTenant}/.default";
            var credential = new ManagedIdentityCredential(settings.ManagedIdentityClientId);
            var accessToken = await credential.GetTokenAsync(new TokenRequestContext([sharePointScope]), cancellationToken);

            _logger.LogInformation("✓ MSI token acquired for scope {Scope}", sharePointScope);
            LogTokenScopesAndRoles(accessToken.Token);
            LogStepEnd(2, "Token MSI pour SharePoint REST");

            LogStepStart(3, "Appel direct SharePoint REST");
            var webUrl = $"https://{settings.SharePointTenant}/{settings.SharePointSitePath.Trim('/')}/_api/web?$select=Id,Title,Url";
            _logger.LogInformation("SharePoint REST URL: {WebUrl}", webUrl);

            using var client = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Get, webUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

            using var response = await client.SendAsync(request, cancellationToken);
            var payload = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("✗ SharePoint REST error: status={StatusCode}", (int)response.StatusCode);
                _logger.LogError("SharePoint REST payload: {Payload}", payload);
                return await CreateErrorResponseAsync(req, (int)response.StatusCode, payload, cancellationToken);
            }

            string? siteId = null;
            string? siteTitle = null;
            string? siteUrl = null;

            using (var doc = JsonDocument.Parse(payload))
            {
                siteId = doc.RootElement.TryGetProperty("Id", out var idElement) ? idElement.GetString() : null;
                siteTitle = doc.RootElement.TryGetProperty("Title", out var titleElement) ? titleElement.GetString() : null;
                siteUrl = doc.RootElement.TryGetProperty("Url", out var urlElement) ? urlElement.GetString() : null;
            }

            _logger.LogInformation("✓ SharePoint REST site resolved: {Title} ({Id})", siteTitle ?? "(null)", siteId ?? "(null)");
            LogStepEnd(3, "Appel direct SharePoint REST");

            var ok = req.CreateResponse(HttpStatusCode.OK);
            await ok.WriteAsJsonAsync(new
            {
                success = true,
                authentication = "managed-identity-direct-sharepoint-rest",
                endpoint = webUrl,
                site = new
                {
                    id = siteId,
                    title = siteTitle,
                    url = siteUrl
                }
            }, cancellationToken);

            return ok;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "✗ SharePointRestManagedIdentityConnection failed: {Message}", ex.Message);
            if (ex.InnerException is not null)
            {
                _logger.LogError("Inner error: {InnerError}", ex.InnerException.Message);
            }

            var error = req.CreateResponse(HttpStatusCode.InternalServerError);
            await error.WriteAsJsonAsync(new
            {
                success = false,
                error = ex.Message
            }, cancellationToken);

            return error;
        }
    }

    private Settings ReadAndValidateSettings()
    {
        LogStepStart(1, "Lecture et validation des variables d'environnement");

        var managedIdentityClientId = Environment.GetEnvironmentVariable("MANAGED_IDENTITY_CLIENT_ID");
        var sharePointTenant = Environment.GetEnvironmentVariable("SHAREPOINT_TENANT");
        var sharePointSitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH");
        var appRegistrationClientId = Environment.GetEnvironmentVariable("APP_REGISTRATION_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");

        _logger.LogInformation(
            "Variables: MANAGED_IDENTITY_CLIENT_ID={MsiStatus}, SHAREPOINT_TENANT={SharePointTenant}, SHAREPOINT_SITE_PATH={SitePath}, APP_REGISTRATION_CLIENT_ID={AppRegistrationStatus}, AZURE_TENANT_ID={TenantStatus}",
            DescribePresence(managedIdentityClientId),
            sharePointTenant ?? "(missing)",
            sharePointSitePath ?? "(missing)",
            DescribePresence(appRegistrationClientId),
            DescribePresence(tenantId));

        var missing = new List<string>();
        if (string.IsNullOrWhiteSpace(managedIdentityClientId)) missing.Add("MANAGED_IDENTITY_CLIENT_ID");
        if (string.IsNullOrWhiteSpace(sharePointTenant)) missing.Add("SHAREPOINT_TENANT");
        if (string.IsNullOrWhiteSpace(sharePointSitePath)) missing.Add("SHAREPOINT_SITE_PATH");

        if (missing.Count > 0)
        {
            var message = $"Missing required environment variables: {string.Join(", ", missing)}";
            _logger.LogError("✗ {Message}", message);
            LogStepEnd(1, "Lecture et validation des variables d'environnement");
            throw new InvalidOperationException(message);
        }

        _logger.LogInformation("✓ Variables validated");
        if (!string.IsNullOrWhiteSpace(appRegistrationClientId))
        {
            _logger.LogInformation("APP_REGISTRATION_CLIENT_ID is not used by this endpoint.");
        }

        LogStepEnd(1, "Lecture et validation des variables d'environnement");
        return new Settings(managedIdentityClientId!, sharePointTenant!, sharePointSitePath!.Trim('/'));
    }

    private async Task<HttpResponseData> CreateErrorResponseAsync(
        HttpRequestData req,
        int statusCode,
        string payload,
        CancellationToken cancellationToken)
    {
        var response = req.CreateResponse((HttpStatusCode)Math.Clamp(statusCode, 100, 599));
        await response.WriteAsJsonAsync(new
        {
            success = false,
            statusCode,
            error = payload
        }, cancellationToken);

        return response;
    }

    private void LogTokenScopesAndRoles(string jwt)
    {
        var handler = new JwtSecurityTokenHandler();
        if (!handler.CanReadToken(jwt))
        {
            _logger.LogWarning("Unable to decode MSI access token.");
            return;
        }

        var token = handler.ReadJwtToken(jwt);
        var scopes = token.Claims
            .Where(c => string.Equals(c.Type, "scp", StringComparison.OrdinalIgnoreCase))
            .Select(c => c.Value)
            .ToList();

        var roles = token.Claims
            .Where(c => string.Equals(c.Type, "roles", StringComparison.OrdinalIgnoreCase))
            .Select(c => c.Value)
            .ToList();

        _logger.LogInformation("MSI token scp: {Scopes}", scopes.Count > 0 ? string.Join(", ", scopes) : "(missing)");
        _logger.LogInformation("MSI token roles: {Roles}", roles.Count > 0 ? string.Join(", ", roles) : "(missing)");
    }

    private void LogStepStart(int stepNumber, string stepName)
    {
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
        _logger.LogInformation("STEP {StepNumber} : {StepName}", stepNumber, stepName);
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
    }

    private void LogStepEnd(int stepNumber, string stepName)
    {
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
        _logger.LogInformation("END STEP {StepNumber} : {StepName}", stepNumber, stepName);
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
    }

    private static string DescribePresence(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "missing" : "set";
    }

    private sealed record Settings(string ManagedIdentityClientId, string SharePointTenant, string SharePointSitePath);
}