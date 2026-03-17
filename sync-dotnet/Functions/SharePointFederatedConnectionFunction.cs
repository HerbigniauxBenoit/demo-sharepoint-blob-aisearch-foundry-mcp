using System.IdentityModel.Tokens.Jwt;
using System.Net;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace SharePointSync.Functions;

public sealed class SharePointFederatedConnectionFunction
{
    private const string TokenExchangeScope = "api://AzureADTokenExchange/.default";
    private const string GraphScope = "https://graph.microsoft.com/.default";
    private readonly ILogger<SharePointFederatedConnectionFunction> _logger;

    public SharePointFederatedConnectionFunction(ILogger<SharePointFederatedConnectionFunction> logger)
    {
        _logger = logger;
    }

    [Function("SharePointFederatedConnection")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sharepoint/federated-connection")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        try
        {
            _logger.LogInformation("SharePointFederatedConnection started: {Method} /api/sharepoint/federated-connection", req.Method);
            var settings = ReadAndValidateSettings();

            LogStepStart(2, "Obtention du token MSI");
            var managedIdentityCredential = new ManagedIdentityCredential(settings.ManagedIdentityClientId);
            var msiToken = await managedIdentityCredential.GetTokenAsync(
                new TokenRequestContext([TokenExchangeScope]),
                cancellationToken);

            _logger.LogInformation("✓ MSI token acquired for scope {Scope}", TokenExchangeScope);
            LogJwtClaims("MSI token", msiToken.Token, ["aud", "appid"]);
            LogStepEnd(2, "Obtention du token MSI");

            LogStepStart(3, "Echange du token MSI via ClientAssertionCredential");
            var sharePointScope = $"https://{settings.SharePointTenant}/.default";
            _logger.LogInformation("Requested SharePoint scope: {Scope}", sharePointScope);

            var clientAssertionCredential = new ClientAssertionCredential(
                settings.TenantId,
                settings.AppRegistrationClientId,
                _ => Task.FromResult(msiToken.Token));

            var sharePointToken = await clientAssertionCredential.GetTokenAsync(
                new TokenRequestContext([sharePointScope]),
                cancellationToken);

            _logger.LogInformation("✓ SharePoint token acquired for scope {Scope}", sharePointScope);
            LogJwtClaims("SharePoint token", sharePointToken.Token, ["aud", "roles", "appid"]);
            LogStepEnd(3, "Echange du token MSI via ClientAssertionCredential");

            LogStepStart(4, "Appel SharePoint via Microsoft Graph");
            var siteIdentifier = $"{settings.SharePointTenant}:/{settings.SharePointSitePath}";
            _logger.LogInformation("Graph scope: {Scope}", GraphScope);
            _logger.LogInformation("Graph site identifier: {SiteIdentifier}", siteIdentifier);

            var graphClient = new GraphServiceClient(clientAssertionCredential, [GraphScope]);
            var site = await graphClient.Sites[siteIdentifier].GetAsync(cancellationToken: cancellationToken);

            _logger.LogInformation("✓ SharePoint site resolved: {DisplayName} ({SiteId})", site?.DisplayName ?? "(null)", site?.Id ?? "(null)");
            LogStepEnd(4, "Appel SharePoint via Microsoft Graph");

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteAsJsonAsync(new
            {
                success = true,
                authentication = "federated-credential",
                managedIdentity = "user-assigned",
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
            _logger.LogError(ex, "✗ SharePointFederatedConnection failed: {Message}", ex.Message);
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
        var appRegistrationClientId = Environment.GetEnvironmentVariable("APP_REGISTRATION_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        var sharePointTenant = Environment.GetEnvironmentVariable("SHAREPOINT_TENANT");
        var sharePointSitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH");

        _logger.LogInformation(
            "Variables: MANAGED_IDENTITY_CLIENT_ID={MsiStatus}, APP_REGISTRATION_CLIENT_ID={AppStatus}, AZURE_TENANT_ID={TenantStatus}, SHAREPOINT_TENANT={SharePointTenant}, SHAREPOINT_SITE_PATH={SitePath}",
            DescribePresence(managedIdentityClientId),
            DescribePresence(appRegistrationClientId),
            DescribePresence(tenantId),
            sharePointTenant ?? "(missing)",
            sharePointSitePath ?? "(missing)");

        var missingVariables = new List<string>();
        if (string.IsNullOrWhiteSpace(managedIdentityClientId))
        {
            missingVariables.Add("MANAGED_IDENTITY_CLIENT_ID");
        }

        if (string.IsNullOrWhiteSpace(appRegistrationClientId))
        {
            missingVariables.Add("APP_REGISTRATION_CLIENT_ID");
        }

        if (string.IsNullOrWhiteSpace(tenantId))
        {
            missingVariables.Add("AZURE_TENANT_ID");
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
            appRegistrationClientId!,
            tenantId!,
            sharePointTenant!,
            sharePointSitePath!.Trim('/'));

        _logger.LogInformation("✓ Variables validated");
        LogStepEnd(1, "Lecture et validation des variables d'environnement");
        return settings;
    }

    private void LogJwtClaims(string tokenLabel, string token, IReadOnlyList<string> claimTypes)
    {
        var handler = new JwtSecurityTokenHandler();
        if (!handler.CanReadToken(token))
        {
            _logger.LogWarning("✗ Unable to decode {TokenLabel}: token is not a readable JWT", tokenLabel);
            return;
        }

        var jwtToken = handler.ReadJwtToken(token);
        foreach (var claimType in claimTypes)
        {
            var values = jwtToken.Claims
                .Where(claim => string.Equals(claim.Type, claimType, StringComparison.OrdinalIgnoreCase))
                .Select(claim => claim.Value)
                .ToList();

            _logger.LogInformation(
                "{TokenLabel} claim {ClaimType}: {ClaimValue}",
                tokenLabel,
                claimType,
                values.Count > 0 ? string.Join(", ", values) : "(missing)");
        }
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

    private sealed record Settings(
        string ManagedIdentityClientId,
        string AppRegistrationClientId,
        string TenantId,
        string SharePointTenant,
        string SharePointSitePath);
}