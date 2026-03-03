using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IdentityModel.Tokens.Jwt;

namespace SharePointSync.Functions.Services;

public class IdentityService
{
    private readonly ILogger<IdentityService> _logger;
    private readonly IConfiguration _configuration;
    private readonly TokenCredential _tokenCredential;

    public IdentityService(
        ILogger<IdentityService> logger,
        IConfiguration configuration,
        TokenCredential tokenCredential)
    {
        _logger = logger;
        _configuration = configuration;
        _tokenCredential = tokenCredential;
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

            // Get access token and parse JWT claims
            var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
            var token = _tokenCredential.GetTokenAsync(tokenRequestContext, CancellationToken.None).Result;

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

            // Log roles if present
            var roles = jwtToken.Claims.Where(c => c.Type == "roles");
            if (roles.Any())
            {
                _logger.LogInformation("Assigned Roles: {Roles}", string.Join(", ", roles.Select(r => r.Value)));
            }

            // Log scopes/scp if present
            var scopes = jwtToken.Claims.FirstOrDefault(c => c.Type == "scp")?.Value;
            if (!string.IsNullOrEmpty(scopes))
            {
                _logger.LogInformation("Delegated Permissions (Scopes): {Scopes}", scopes);
            }

            _logger.LogInformation("================================================");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error logging identity details");
        }
    }
}
