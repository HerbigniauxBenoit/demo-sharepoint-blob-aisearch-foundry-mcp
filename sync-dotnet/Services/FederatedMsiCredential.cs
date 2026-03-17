using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;

namespace SharePointSync.Functions.Services;

public sealed class FederatedMsiCredential : TokenCredential
{
    private const string TokenExchangeScope = "api://AzureADTokenExchange/.default";

    private readonly string _tenantId;
    private readonly string _appRegistrationClientId;
    private readonly string? _msiClientId;
    private readonly ManagedIdentityCredential _managedIdentityCredential;
    private readonly IConfidentialClientApplication _confidentialClientApplication;
    private readonly ILogger? _logger;

    public FederatedMsiCredential(
        string tenantId,
        string appRegistrationClientId,
        string? msiClientId = null,
        ILogger? logger = null)
    {
        if (string.IsNullOrWhiteSpace(tenantId))
        {
            throw new ArgumentException("TenantId is required.", nameof(tenantId));
        }

        if (string.IsNullOrWhiteSpace(appRegistrationClientId))
        {
            throw new ArgumentException("AppRegistrationClientId is required.", nameof(appRegistrationClientId));
        }

        _tenantId = tenantId;
        _appRegistrationClientId = appRegistrationClientId;
        _msiClientId = string.IsNullOrWhiteSpace(msiClientId) ? null : msiClientId;
        _managedIdentityCredential = _msiClientId is null
            ? new ManagedIdentityCredential()
            : new ManagedIdentityCredential(_msiClientId);
        _confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(_appRegistrationClientId)
            .WithTenantId(_tenantId)
            .WithClientAssertion(GetManagedIdentityAssertionAsync)
            .Build();
        _logger = logger;
    }

    public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        return AcquireTokenAsync(requestContext, cancellationToken).GetAwaiter().GetResult();
    }

    public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        return new ValueTask<AccessToken>(AcquireTokenAsync(requestContext, cancellationToken));
    }

    private async Task<AccessToken> AcquireTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        if (requestContext.Scopes is null || requestContext.Scopes.Length == 0)
        {
            throw new AuthenticationFailedException("FederatedMsiCredential requires at least one scope to request a token.");
        }

        try
        {
            var result = await _confidentialClientApplication
                .AcquireTokenForClient(requestContext.Scopes)
                .ExecuteAsync(cancellationToken);

            return new AccessToken(result.AccessToken, result.ExpiresOn);
        }
        catch (MsalServiceException ex)
        {
            throw new AuthenticationFailedException(
                $"Federated token exchange failed in Entra ID for tenant '{_tenantId}' and app registration '{_appRegistrationClientId}'. Details: {ex.Message}",
                ex);
        }
        catch (MsalClientException ex)
        {
            throw new AuthenticationFailedException(
                $"Federated token exchange failed on client side while using MSI assertion. Verify managed identity and federated credential configuration. Details: {ex.Message}",
                ex);
        }
        catch (AuthenticationFailedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new AuthenticationFailedException(
                $"Unexpected federated authentication error. Requested scopes: {string.Join(", ", requestContext.Scopes)}.",
                ex);
        }
    }

    private async Task<string> GetManagedIdentityAssertionAsync(AssertionRequestOptions assertionRequestOptions)
    {
        var msiAssertionToken = await _managedIdentityCredential.GetTokenAsync(
            new TokenRequestContext([TokenExchangeScope]),
            assertionRequestOptions.CancellationToken);

        _logger?.LogInformation(
            "FederatedMsiCredential: MSI assertion acquired for {IdentityMode} identity. Expires at {ExpiresOnUtc:G}",
            _msiClientId is null ? "system-assigned" : "user-assigned",
            msiAssertionToken.ExpiresOn.UtcDateTime);

        return msiAssertionToken.Token;
    }
}
