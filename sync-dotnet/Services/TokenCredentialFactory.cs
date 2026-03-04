using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SharePointSync.Functions.Services;

public sealed class TokenCredentialFactory
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<TokenCredentialFactory> _logger;

    public TokenCredentialFactory(IConfiguration configuration, ILogger<TokenCredentialFactory> logger)
    {
        _configuration = configuration;
        _logger = logger;
    }

    public TokenCredential Create()
    {
        var managedIdentityClientId = _configuration["AZURE_CLIENT_ID"];

        if (string.IsNullOrWhiteSpace(managedIdentityClientId))
        {
            _logger.LogInformation("Initializing DefaultAzureCredential (AZURE_CLIENT_ID not set)");
            return new DefaultAzureCredential();
        }

        _logger.LogInformation("Initializing User Assigned ManagedIdentityCredential with client ID: {ClientId}", managedIdentityClientId);
        return new ManagedIdentityCredential(managedIdentityClientId);
    }
}
