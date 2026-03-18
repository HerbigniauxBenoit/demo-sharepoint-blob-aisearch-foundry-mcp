using System.Net;
using System.IdentityModel.Tokens.Jwt;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace SharePointSync.Functions;

public sealed class SharePointManagedIdentityConnectionFunction
{
    private const string GraphScope = "https://graph.microsoft.com/.default";
    private readonly ILogger<SharePointManagedIdentityConnectionFunction> _logger;

    public SharePointManagedIdentityConnectionFunction(ILogger<SharePointManagedIdentityConnectionFunction> logger)
    {
        _logger = logger;
    }

    [Function("SharePointManagedIdentityConnection")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sharepoint/msi-direct-connection")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        try
        {
            _logger.LogInformation("SharePointManagedIdentityConnection started: {Method} /api/sharepoint/msi-direct-connection", req.Method);
            var settings = ReadAndValidateSettings();

            LogStepStart(2, "Authentification directe avec la Managed Identity");
            var managedIdentityCredential = new ManagedIdentityCredential(settings.ManagedIdentityClientId);
            _logger.LogInformation("Requested Graph scope: {Scope}", GraphScope);

            var token = await managedIdentityCredential.GetTokenAsync(
                new TokenRequestContext([GraphScope]),
                cancellationToken);

            LogTokenScopesAndRoles(token.Token);

            var graphClient = new GraphServiceClient(managedIdentityCredential, [GraphScope]);
            var siteIdentifier = $"{settings.SharePointTenant}:/{settings.SharePointSitePath}";
            _logger.LogInformation("Graph site identifier: {SiteIdentifier}", siteIdentifier);

            var site = await graphClient.Sites[siteIdentifier].GetAsync(cancellationToken: cancellationToken);

            _logger.LogInformation("✓ SharePoint site resolved with direct MSI: {DisplayName} ({SiteId})", site?.DisplayName ?? "(null)", site?.Id ?? "(null)");
            LogStepEnd(2, "Authentification directe avec la Managed Identity");

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteAsJsonAsync(new
            {
                success = true,
                authentication = "managed-identity-direct",
                site = new
                {
                    id = site?.Id,
                    displayName = site?.DisplayName,
                    identifier = siteIdentifier
                }
            }, cancellationToken);

            return response;
        }
        catch (ODataError ex)
        {
            _logger.LogError("✗ Graph error: status={StatusCode}, code={Code}, message={Message}", ex.ResponseStatusCode, ex.Error?.Code ?? "(null)", ex.Error?.Message ?? ex.Message);
            if (ex.InnerException is not null)
            {
                _logger.LogError("Inner error: {InnerError}", ex.InnerException.Message);
            }

            return await CreateErrorResponseAsync(req, ex.Error?.Message ?? ex.Message, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "✗ SharePointManagedIdentityConnection failed: {Message}", ex.Message);
            if (ex.InnerException is not null)
            {
                _logger.LogError("Inner error: {InnerError}", ex.InnerException.Message);
            }

            return await CreateErrorResponseAsync(req, ex.Message, cancellationToken);
        }
    }

    private Settings ReadAndValidateSettings()
    {
        LogStepStart(1, "Lecture et validation des variables d'environnement");

        var managedIdentityClientId = Environment.GetEnvironmentVariable("MANAGED_IDENTITY_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        var sharePointTenant = Environment.GetEnvironmentVariable("SHAREPOINT_TENANT");
        var sharePointSitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH");
        var appRegistrationClientId = Environment.GetEnvironmentVariable("APP_REGISTRATION_CLIENT_ID");

        _logger.LogInformation(
            "Variables: MANAGED_IDENTITY_CLIENT_ID={MsiStatus}, AZURE_TENANT_ID={TenantStatus}, SHAREPOINT_TENANT={SharePointTenant}, SHAREPOINT_SITE_PATH={SitePath}, APP_REGISTRATION_CLIENT_ID={AppRegistrationStatus}",
            DescribePresence(managedIdentityClientId),
            DescribePresence(tenantId),
            sharePointTenant ?? "(missing)",
            sharePointSitePath ?? "(missing)",
            DescribePresence(appRegistrationClientId));

        var missingVariables = new List<string>();
        if (string.IsNullOrWhiteSpace(managedIdentityClientId))
        {
            missingVariables.Add("MANAGED_IDENTITY_CLIENT_ID");
        }

        if (string.IsNullOrWhiteSpace(sharePointTenant))
        {
            missingVariables.Add("SHAREPOINT_TENANT");
        }

        if (string.IsNullOrWhiteSpace(sharePointSitePath))
        {
            missingVariables.Add("SHAREPOINT_SITE_PATH");
        }

        if (missingVariables.Count > 0)
        {
            var message = $"Missing required environment variables: {string.Join(", ", missingVariables)}";
            _logger.LogError("✗ {Message}", message);
            LogStepEnd(1, "Lecture et validation des variables d'environnement");
            throw new InvalidOperationException(message);
        }

        var settings = new Settings(
            managedIdentityClientId!,
            tenantId,
            sharePointTenant!,
            sharePointSitePath!.Trim('/'));

        _logger.LogInformation("✓ Variables validated");
        if (string.IsNullOrWhiteSpace(settings.TenantId))
        {
            _logger.LogInformation("AZURE_TENANT_ID is not required for direct MSI mode.");
        }

        if (!string.IsNullOrWhiteSpace(appRegistrationClientId))
        {
            _logger.LogInformation("APP_REGISTRATION_CLIENT_ID is present but not used by this endpoint.");
        }

        LogStepEnd(1, "Lecture et validation des variables d'environnement");
        return settings;
    }

    private async Task<HttpResponseData> CreateErrorResponseAsync(
        HttpRequestData req,
        string message,
        CancellationToken cancellationToken)
    {
        var response = req.CreateResponse(HttpStatusCode.InternalServerError);
        await response.WriteAsJsonAsync(new
        {
            success = false,
            error = message
        }, cancellationToken);

        return response;
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

        _logger.LogInformation(
            "MSI token scp: {Scopes}",
            scopes.Count > 0 ? string.Join(", ", scopes) : "(missing)");

        _logger.LogInformation(
            "MSI token roles: {Roles}",
            roles.Count > 0 ? string.Join(", ", roles) : "(missing)");
    }

    private sealed record Settings(
        string ManagedIdentityClientId,
        string? TenantId,
        string SharePointTenant,
        string SharePointSitePath);
}