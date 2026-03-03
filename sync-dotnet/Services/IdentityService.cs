using Azure.Core;
using Azure.Identity;
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

            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
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
                var siteDoc = JsonDocument.Parse(content);
                var siteId = siteDoc.RootElement.GetProperty("id").GetString();
                var siteName = siteDoc.RootElement.GetProperty("displayName").GetString();

                _logger.LogInformation("✅ SUCCESS: Access granted to site '{SiteName}' (ID: {SiteId})", siteName, siteId);
                _logger.LogInformation("==================================================");
                return (true, $"Access granted to site '{siteName}'");
            }
            else
            {
                var errorContent = await response.Content.ReadAsStringAsync(cancellationToken);
                _logger.LogError("❌ FAILED: Access denied to site. Status: {StatusCode}", response.StatusCode);
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
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "❌ EXCEPTION while validating SharePoint access");
            return (false, $"Exception: {ex.Message}");
        }
    }

    public void LogIdentityDetails()
    {
        try
        {
            _logger.LogInformation("========== MANAGED IDENTITY INFORMATION ==========");

            // Log identity type and client ID
            var clientId = _configuration["AZURE_CLIENT_ID"];
            if (!string.IsNullOrEmpty(clientId))
            {
                _logger.LogInformation("Identity Type: User Assigned");
                _logger.LogInformation("Client ID: {ClientId}", clientId);
            }
            else
            {
                _logger.LogInformation("Identity Type: System Assigned");
            }

            // Log tenant ID if available
            var tenantId = _configuration["AZURE_TENANT_ID"];
            if (!string.IsNullOrEmpty(tenantId))
            {
                _logger.LogInformation("Tenant ID: {TenantId}", tenantId);
            }

            _logger.LogInformation("Attempting to acquire token from Microsoft Graph...");

            // Get access token and parse JWT claims
            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var tokenTask = _tokenCredential.GetTokenAsync(tokenRequestContext, CancellationToken.None);
            var token = tokenTask.GetAwaiter().GetResult();

            _logger.LogInformation("✅ Token acquired successfully!");

            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadJwtToken(token.Token);

            _logger.LogInformation("Token Expiration: {Expiration}", 
                DateTimeOffset.FromUnixTimeSeconds(jwtToken.Claims
                    .FirstOrDefault(c => c.Type == "exp")?.Value is string expStr && long.TryParse(expStr, out var exp)
                    ? exp : 0).DateTime);

            // Log principal ID (OID)
            var oid = jwtToken.Claims.FirstOrDefault(c => c.Type == "oid")?.Value;
            if (!string.IsNullOrEmpty(oid))
            {
                _logger.LogInformation("Principal ID (OID): {PrincipalId}", oid);
            }

            // Log app ID
            var appId = jwtToken.Claims.FirstOrDefault(c => c.Type == "appid")?.Value;
            if (!string.IsNullOrEmpty(appId))
            {
                _logger.LogInformation("App ID: {AppId}", appId);
            }

            // Log identity name
            var uniqueName = jwtToken.Claims.FirstOrDefault(c => c.Type == "unique_name")?.Value;
            if (!string.IsNullOrEmpty(uniqueName))
            {
                _logger.LogInformation("Identity Name: {UniqueName}", uniqueName);
            }

            // Log roles if present
            var roles = jwtToken.Claims.Where(c => c.Type == "roles").ToList();
            if (roles.Any())
            {
                _logger.LogInformation("✅ Assigned Roles ({Count}):", roles.Count);
                foreach (var role in roles)
                {
                    _logger.LogInformation("  - {Role}", role.Value);
                    
                    // Check for specific permissions
                    if (role.Value == "Sites.Selected")
                    {
                        _logger.LogInformation("    ℹ️  Sites.Selected: Grants access to specific sites only (most secure)");
                    }
                    else if (role.Value == "Sites.Read.All")
                    {
                        _logger.LogInformation("    ℹ️  Sites.Read.All: Grants read access to all SharePoint sites");
                    }
                    else if (role.Value == "Sites.ReadWrite.All")
                    {
                        _logger.LogInformation("    ℹ️  Sites.ReadWrite.All: Grants read/write access to all SharePoint sites");
                    }
                }
            }
            else
            {
                _logger.LogWarning("⚠️  WARNING: No application roles found in token!");
                _logger.LogWarning("    Make sure to assign Graph API permissions (Sites.Selected or Sites.Read.All)");
            }

            // Log scopes/scp if present
            var scopes = jwtToken.Claims.FirstOrDefault(c => c.Type == "scp")?.Value;
            if (!string.IsNullOrEmpty(scopes))
            {
                _logger.LogInformation("Delegated Permissions (Scopes): {Scopes}", scopes);
            }

            // Log all claims for complete debugging
            _logger.LogInformation("--- All JWT Claims (for debugging) ---");
            foreach (var claim in jwtToken.Claims)
            {
                _logger.LogDebug("  {ClaimType}: {ClaimValue}", claim.Type, claim.Value);
            }

            _logger.LogInformation("==================================================");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "❌ CRITICAL ERROR: Unable to log identity details. Message: {Message}", ex.Message);
            throw;
        }
    }
}
