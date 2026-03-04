using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Text.Json;

namespace SharePointSync.Functions.Services;

public class IdentityService
{
    private readonly ILogger<IdentityService> _logger;
    private readonly IConfiguration _configuration;
    private readonly TokenCredential _tokenCredential;
    private readonly IHttpClientFactory _httpClientFactory;

    public IdentityService(
        ILogger<IdentityService> logger,
        IConfiguration configuration,
        TokenCredential tokenCredential,
        IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _configuration = configuration;
        _tokenCredential = tokenCredential;
        _httpClientFactory = httpClientFactory;
    }

    public async Task<(bool Success, string Message)> ValidateSharePointAccessAsync(string siteUrl, CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogInformation("========== VALIDATING SHAREPOINT ACCESS ==========");
            _logger.LogInformation("Target Site URL: {SiteUrl}", siteUrl);

            var tokenRequestContext = new TokenRequestContext(["https://graph.microsoft.com/.default"]);
            var token = await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);

            var siteUri = new Uri(siteUrl);
            var relativePath = siteUri.AbsolutePath.TrimStart('/');
            var siteLookup = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/{relativePath}";

            _logger.LogInformation("Testing access to: {SiteLookup}", siteLookup);

            using var client = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Get, siteLookup);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            var response = await client.SendAsync(request, cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync(cancellationToken);
                using var siteDoc = JsonDocument.Parse(content);
                var siteId = siteDoc.RootElement.GetProperty("id").GetString();
                var siteName = siteDoc.RootElement.GetProperty("displayName").GetString();

                _logger.LogInformation("SUCCESS: Access granted to site '{SiteName}' (ID: {SiteId})", siteName, siteId);
                _logger.LogInformation("==================================================");
                return (true, $"Access granted to site '{siteName}'");
            }

            var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
            _logger.LogError("FAILED: Access denied to site. Status: {StatusCode}", response.StatusCode);
            _logger.LogError("Error details: {ErrorContent}", errorContent);
            _logger.LogInformation("==================================================");

            return response.StatusCode switch
            {
                System.Net.HttpStatusCode.Forbidden => (false, "Access denied (403). The identity does not have permissions to access this site. Check Sites.Selected permissions."),
                System.Net.HttpStatusCode.NotFound => (false, "Site not found (404). Verify the SharePoint site URL is correct."),
                System.Net.HttpStatusCode.Unauthorized => (false, "Unauthorized (401). The token may be invalid or the identity needs proper Graph API permissions."),
                _ => (false, $"Failed with status {response.StatusCode}: {errorContent}")
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "EXCEPTION while validating SharePoint access");
            return (false, $"Exception: {ex.Message}");
        }
    }

    public async Task LogIdentityDetailsAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogInformation("========== MANAGED IDENTITY INFORMATION ==========");

            var configuredClientId = _configuration["AZURE_CLIENT_ID"];
            if (!string.IsNullOrEmpty(configuredClientId))
            {
                _logger.LogInformation("Identity Type: User Assigned");
                _logger.LogInformation("Configured Client ID: {ClientId}", configuredClientId);
            }
            else
            {
                _logger.LogInformation("Identity Type: System Assigned");
            }

            var tenantId = _configuration["AZURE_TENANT_ID"];
            if (!string.IsNullOrEmpty(tenantId))
            {
                _logger.LogInformation("Tenant ID: {TenantId}", tenantId);
            }

            var tokenRequestContext = new TokenRequestContext(["https://graph.microsoft.com/.default"]);
            var token = await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadJwtToken(token.Token);

            var appId = jwtToken.Claims.FirstOrDefault(c => c.Type == "appid")?.Value;
            var oid = jwtToken.Claims.FirstOrDefault(c => c.Type == "oid")?.Value;
            var miResourceId = jwtToken.Claims.FirstOrDefault(c => c.Type == "xms_mirid")?.Value;
            var userAssignedName = TryExtractManagedIdentityName(miResourceId);

            _logger.LogInformation("Token App ID: {AppId}", appId ?? "n/a");
            _logger.LogInformation("Principal Object ID (OID): {Oid}", oid ?? "n/a");
            _logger.LogInformation("Managed Identity Resource ID: {ResourceId}", miResourceId ?? "n/a");
            _logger.LogInformation("User Assigned Identity Name: {IdentityName}", userAssignedName ?? "n/a");

            var roles = jwtToken.Claims.Where(c => c.Type == "roles").Select(c => c.Value).Distinct().ToList();
            if (roles.Count > 0)
            {
                _logger.LogInformation("Global Graph App Roles ({Count}): {Roles}", roles.Count, string.Join(", ", roles));
                _logger.LogInformation("Sites.Selected present: {HasSitesSelected}", roles.Contains("Sites.Selected", StringComparer.OrdinalIgnoreCase));
            }
            else
            {
                _logger.LogWarning("No Graph application roles found in token.");
            }

            var siteUrl = _configuration["SHAREPOINT_SITE_URL"];
            if (!string.IsNullOrWhiteSpace(siteUrl))
            {
                await LogTargetSitePermissionsAsync(token.Token, siteUrl, appId, cancellationToken);
            }
            else
            {
                _logger.LogWarning("SHAREPOINT_SITE_URL not configured. Site-level rights not checked.");
            }

            _logger.LogInformation("==================================================");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CRITICAL ERROR: Unable to log identity details");
            throw;
        }
    }

    private async Task LogTargetSitePermissionsAsync(string accessToken, string siteUrl, string? appId, CancellationToken cancellationToken)
    {
        var siteUri = new Uri(siteUrl);
        var relativePath = siteUri.AbsolutePath.TrimStart('/');
        var siteLookup = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/{relativePath}";

        using var client = _httpClientFactory.CreateClient();

        using var siteRequest = new HttpRequestMessage(HttpMethod.Get, siteLookup);
        siteRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        using var siteResponse = await client.SendAsync(siteRequest, cancellationToken);

        if (!siteResponse.IsSuccessStatusCode)
        {
            var siteError = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
            _logger.LogWarning("Could not resolve target site for permissions check. Status: {Status}. Body: {Body}", siteResponse.StatusCode, siteError);
            return;
        }

        var sitePayload = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
        using var siteDoc = JsonDocument.Parse(sitePayload);
        var siteId = siteDoc.RootElement.GetProperty("id").GetString();
        var siteName = siteDoc.RootElement.GetProperty("displayName").GetString();

        _logger.LogInformation("Target SharePoint Site: {SiteName} ({SiteId})", siteName, siteId);

        if (string.IsNullOrWhiteSpace(siteId))
        {
            _logger.LogWarning("Site ID empty, cannot check site-level permissions.");
            return;
        }

        using var permRequest = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/sites/{siteId}/permissions");
        permRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        using var permResponse = await client.SendAsync(permRequest, cancellationToken);

        if (!permResponse.IsSuccessStatusCode)
        {
            var permError = await permResponse.Content.ReadAsStringAsync(cancellationToken);
            _logger.LogWarning("Could not read site permissions. Status: {Status}. Body: {Body}", permResponse.StatusCode, permError);
            return;
        }

        var permPayload = await permResponse.Content.ReadAsStringAsync(cancellationToken);
        using var permDoc = JsonDocument.Parse(permPayload);

        var allPermissionRoles = new List<string>();
        var appSpecificRoles = new List<string>();

        if (permDoc.RootElement.TryGetProperty("value", out var permissions))
        {
            foreach (var permission in permissions.EnumerateArray())
            {
                if (permission.TryGetProperty("roles", out var rolesElement))
                {
                    var roles = rolesElement.EnumerateArray().Select(r => r.GetString()).Where(r => !string.IsNullOrWhiteSpace(r)).Cast<string>().ToList();
                    allPermissionRoles.AddRange(roles);

                    if (!string.IsNullOrWhiteSpace(appId)
                        && permission.TryGetProperty("grantedToIdentitiesV2", out var identities)
                        && identities.EnumerateArray().Any(i =>
                            i.TryGetProperty("application", out var app)
                            && app.TryGetProperty("id", out var id)
                            && string.Equals(id.GetString(), appId, StringComparison.OrdinalIgnoreCase)))
                    {
                        appSpecificRoles.AddRange(roles);
                    }
                }
            }
        }

        _logger.LogInformation("SharePoint Site Permission Roles (all assignments): {Roles}",
            allPermissionRoles.Count > 0 ? string.Join(", ", allPermissionRoles.Distinct(StringComparer.OrdinalIgnoreCase)) : "none visible");

        _logger.LogInformation("SharePoint Rights for current Managed Identity (AppId={AppId}): {Roles}",
            appId ?? "n/a",
            appSpecificRoles.Count > 0 ? string.Join(", ", appSpecificRoles.Distinct(StringComparer.OrdinalIgnoreCase)) : "none found for this app (or not visible)");
    }

    private static string? TryExtractManagedIdentityName(string? miResourceId)
    {
        if (string.IsNullOrWhiteSpace(miResourceId))
        {
            return null;
        }

        var marker = "/userAssignedIdentities/";
        var idx = miResourceId.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0)
        {
            return null;
        }

        var start = idx + marker.Length;
        if (start >= miResourceId.Length)
        {
            return null;
        }

        return miResourceId[start..].Split('/', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
    }
}
