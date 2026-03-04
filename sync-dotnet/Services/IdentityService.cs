using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
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

            var token = await GetGraphTokenAsync(cancellationToken);
            var siteUri = new Uri(siteUrl);
            var relativePath = siteUri.AbsolutePath.TrimStart('/');
            var siteLookup = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/{relativePath}";

            _logger.LogInformation("Testing access to: {SiteLookup}", siteLookup);

            using var client = _httpClientFactory.CreateClient();
            using var request = new HttpRequestMessage(HttpMethod.Get, siteLookup);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            using var response = await client.SendAsync(request, cancellationToken);

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
                HttpStatusCode.Forbidden => (false, "Access denied (403). The identity does not have permissions to access this site. Check Sites.Selected grants on target site."),
                HttpStatusCode.NotFound => (false, "Site not found (404). Verify the SharePoint site URL is correct."),
                HttpStatusCode.Unauthorized => (false, "Unauthorized (401). The token may be invalid or Graph permissions are missing."),
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
            if (!string.IsNullOrWhiteSpace(configuredClientId))
            {
                _logger.LogInformation("Identity Type: User Assigned");
                _logger.LogInformation("Configured Client ID: {ClientId}", configuredClientId);
            }
            else
            {
                _logger.LogInformation("Identity Type: System Assigned");
            }

            var tenantId = _configuration["AZURE_TENANT_ID"];
            if (!string.IsNullOrWhiteSpace(tenantId))
            {
                _logger.LogInformation("Configured Tenant ID: {TenantId}", tenantId);
            }

            var token = await GetGraphTokenAsync(cancellationToken);
            var jwtToken = new JwtSecurityTokenHandler().ReadJwtToken(token.Token);

            var appId = jwtToken.Claims.FirstOrDefault(c => c.Type == "appid")?.Value;
            var oid = jwtToken.Claims.FirstOrDefault(c => c.Type == "oid")?.Value;
            var audience = jwtToken.Claims.FirstOrDefault(c => c.Type == "aud")?.Value;
            var tokenTenantId = jwtToken.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;
            var miResourceId = jwtToken.Claims.FirstOrDefault(c => c.Type == "xms_mirid")?.Value;
            var userAssignedName = TryExtractManagedIdentityName(miResourceId);

            _logger.LogInformation("Token audience (aud): {Audience}", audience ?? "n/a");
            _logger.LogInformation("Token tenant (tid): {Tid}", tokenTenantId ?? "n/a");
            _logger.LogInformation("Token App ID (appid): {AppId}", appId ?? "n/a");
            _logger.LogInformation("Principal Object ID (oid): {Oid}", oid ?? "n/a");
            _logger.LogInformation("Managed Identity Resource ID (xms_mirid): {ResourceId}", miResourceId ?? "n/a");
            _logger.LogInformation("User Assigned Identity Name: {IdentityName}", userAssignedName ?? "n/a (claim xms_mirid absente)");

            if (!string.IsNullOrWhiteSpace(configuredClientId) && !string.IsNullOrWhiteSpace(appId))
            {
                _logger.LogInformation("ConfiguredClientId == Token appid: {IsMatch}",
                    string.Equals(configuredClientId, appId, StringComparison.OrdinalIgnoreCase));
            }

            var roles = GetRoles(jwtToken);
            if (roles.Count > 0)
            {
                _logger.LogInformation("Global Graph App Roles ({Count}): {Roles}", roles.Count, string.Join(", ", roles));
                _logger.LogInformation("Sites.Selected present: {HasSitesSelected}", roles.Contains("Sites.Selected", StringComparer.OrdinalIgnoreCase));
            }
            else
            {
                _logger.LogWarning("No Graph application roles found in token.");
                _logger.LogWarning("Check Graph app role assignments on the Managed Identity service principal and refresh token cache.");
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

    public async Task LogCandidateSitesForTargetAsync(string siteUrl, CancellationToken cancellationToken = default)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                _logger.LogWarning("No site URL provided for candidate site discovery.");
                return;
            }

            var token = await GetGraphTokenAsync(cancellationToken);
            var roles = GetRoles(new JwtSecurityTokenHandler().ReadJwtToken(token.Token));
            var canListSites = roles.Contains("Sites.Read.All", StringComparer.OrdinalIgnoreCase)
                               || roles.Contains("Sites.ReadWrite.All", StringComparer.OrdinalIgnoreCase);

            if (!canListSites)
            {
                _logger.LogInformation("Skipping global site listing: token has no Sites.Read.All/Sites.ReadWrite.All (Sites.Selected only is expected)." );
                return;
            }

            using var client = _httpClientFactory.CreateClient();
            await LogExistingSitesAsync(client, token.Token, new Uri(siteUrl), cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not run candidate site discovery pre-check.");
        }
    }

    private async Task<AccessToken> GetGraphTokenAsync(CancellationToken cancellationToken)
    {
        var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
        return await _tokenCredential.GetTokenAsync(tokenRequestContext, cancellationToken);
    }

    private static List<string> GetRoles(JwtSecurityToken jwtToken) =>
        jwtToken.Claims
            .Where(c => c.Type == "roles")
            .Select(c => c.Value)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

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
                if (!permission.TryGetProperty("roles", out var rolesElement))
                {
                    continue;
                }

                var roles = rolesElement.EnumerateArray()
                    .Select(r => r.GetString())
                    .Where(r => !string.IsNullOrWhiteSpace(r))
                    .Cast<string>()
                    .ToList();

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

        _logger.LogInformation("SharePoint Site Permission Roles (all assignments): {Roles}",
            allPermissionRoles.Count > 0
                ? string.Join(", ", allPermissionRoles.Distinct(StringComparer.OrdinalIgnoreCase))
                : "none visible");

        _logger.LogInformation("SharePoint Rights for current Managed Identity (AppId={AppId}): {Roles}",
            appId ?? "n/a",
            appSpecificRoles.Count > 0
                ? string.Join(", ", appSpecificRoles.Distinct(StringComparer.OrdinalIgnoreCase))
                : "none found for this app (or not visible)");
    }

    private async Task LogExistingSitesAsync(HttpClient client, string accessToken, Uri targetSiteUri, CancellationToken cancellationToken)
    {
        try
        {
            _logger.LogInformation("Listing candidate sites from Graph for hostname {Host}...", targetSiteUri.Host);

            var searchTerm = targetSiteUri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
            var candidateUrls = new[]
            {
                "https://graph.microsoft.com/v1.0/sites?search=*",
                string.IsNullOrWhiteSpace(searchTerm) ? null : $"https://graph.microsoft.com/v1.0/sites?search={Uri.EscapeDataString(searchTerm)}"
            }.Where(u => !string.IsNullOrWhiteSpace(u)).Cast<string>();

            foreach (var url in candidateUrls)
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                using var res = await client.SendAsync(req, cancellationToken);

                if (!res.IsSuccessStatusCode)
                {
                    var body = await res.Content.ReadAsStringAsync(cancellationToken);
                    _logger.LogWarning("Unable to list sites with '{Url}'. Status: {Status}. Body: {Body}", url, res.StatusCode, body);
                    continue;
                }

                var payload = await res.Content.ReadAsStringAsync(cancellationToken);
                using var doc = JsonDocument.Parse(payload);

                if (!doc.RootElement.TryGetProperty("value", out var sites) || sites.GetArrayLength() == 0)
                {
                    _logger.LogInformation("No sites returned by query: {Url}", url);
                    continue;
                }

                var rows = sites.EnumerateArray()
                    .Select(site => new
                    {
                        Id = site.TryGetProperty("id", out var idEl) ? idEl.GetString() : null,
                        DisplayName = site.TryGetProperty("displayName", out var dnEl) ? dnEl.GetString() : null,
                        WebUrl = site.TryGetProperty("webUrl", out var wuEl) ? wuEl.GetString() : null
                    })
                    .Where(s => !string.IsNullOrWhiteSpace(s.WebUrl)
                                && s.WebUrl!.Contains(targetSiteUri.Host, StringComparison.OrdinalIgnoreCase))
                    .Take(30)
                    .ToList();

                _logger.LogInformation("Candidate sites from query '{Url}' (host {Host}, count={Count}):", url, targetSiteUri.Host, rows.Count);
                foreach (var row in rows)
                {
                    _logger.LogInformation(" - Site: '{DisplayName}' | Url: {WebUrl} | Id: {Id}",
                        row.DisplayName ?? "(no displayName)",
                        row.WebUrl ?? "(no webUrl)",
                        row.Id ?? "(no id)");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not list candidate sites from Graph.");
        }
    }

    private static string? TryExtractManagedIdentityName(string? miResourceId)
    {
        if (string.IsNullOrWhiteSpace(miResourceId))
        {
            return null;
        }

        const string marker = "/userAssignedIdentities/";
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
