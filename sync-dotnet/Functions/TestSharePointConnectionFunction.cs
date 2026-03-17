using System.Net;
using System.Net.Http.Headers;
using System.IdentityModel.Tokens.Jwt;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;

namespace SharePointSync.Functions.Functions;

public class TestSharePointConnectionFunction
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<TestSharePointConnectionFunction> _logger;

    public TestSharePointConnectionFunction(
        IHttpClientFactory httpClientFactory,
        ILogger<TestSharePointConnectionFunction> logger)
    {
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    [Function("TestSharePointConnection")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "test-sharepoint")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("═══════════════════════════════════════════════════════════════════");
        _logger.LogInformation("START: TestSharePointConnection Function");
        _logger.LogInformation("═══════════════════════════════════════════════════════════════════");
        _logger.LogInformation("Timestamp (UTC): {Timestamp}", DateTime.UtcNow.ToString("G"));

        var userAssignedClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
        TokenCredential tokenCredential;

        if (!string.IsNullOrWhiteSpace(userAssignedClientId))
        {
            _logger.LogInformation("IDENTITY: User Assigned Identity (Client ID: {ClientId})", userAssignedClientId);
            tokenCredential = new ManagedIdentityCredential(userAssignedClientId);
        }
        else
        {
            _logger.LogInformation("IDENTITY: Default Azure Credential (auto-detection)");
            _logger.LogInformation("Will try: Environment Vars -> Managed Identity -> Azure CLI -> Visual Studio");
            tokenCredential = new DefaultAzureCredential();
        }

        try
        {
            _logger.LogInformation("─────────────────────────────────────────────────────────────────");
            _logger.LogInformation("STEP 1: Parsing request body");
            _logger.LogInformation("─────────────────────────────────────────────────────────────────");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            _logger.LogInformation("Request body length: {BodyLength} chars", requestBody.Length);

            if (string.IsNullOrWhiteSpace(requestBody))
            {
                _logger.LogWarning("Request body is empty");
                return new BadRequestObjectResult(new { error = "Request body cannot be empty. Expected JSON with siteUrl property." });
            }

            using (var jsonDoc = JsonDocument.Parse(requestBody))
            {
                if (!jsonDoc.RootElement.TryGetProperty("siteUrl", out var siteUrlElement))
                {
                    _logger.LogError("Missing 'siteUrl' property in request body");
                    return new BadRequestObjectResult(new { error = "Request must contain 'siteUrl' property" });
                }

                var siteUrl = siteUrlElement.GetString();
                if (string.IsNullOrWhiteSpace(siteUrl))
                {
                    _logger.LogError("'siteUrl' is empty");
                    return new BadRequestObjectResult(new { error = "'siteUrl' cannot be empty" });
                }

                _logger.LogInformation("Target SharePoint site URL: {SiteUrl}", siteUrl);

                _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                _logger.LogInformation("STEP 2: Getting access token from Azure Identity");
                _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                AccessToken token;
                string? tokenTenantId = null;
                try
                {
                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                    token = await tokenCredential.GetTokenAsync(
                        new TokenRequestContext(["https://graph.microsoft.com/.default"]),
                        cancellationToken);
                    stopwatch.Stop();

                    _logger.LogInformation("SUCCESS: Token acquired successfully");
                    _logger.LogInformation("Token acquisition time: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
                    _logger.LogInformation("Token expires at (UTC): {ExpiresOn:G}", token.ExpiresOn);

                    try
                    {
                        var jwtHandler = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
                        if (jwtHandler.CanReadToken(token.Token))
                        {
                            var jwtToken = jwtHandler.ReadJwtToken(token.Token);
                            tokenTenantId = jwtToken.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;
                            var tokenAudience = jwtToken.Claims.FirstOrDefault(c => c.Type == "aud")?.Value;
                            var tokenAppId = jwtToken.Claims.FirstOrDefault(c => c.Type == "appid" || c.Type == "azp")?.Value;
                            var tokenScopes = jwtToken.Claims.Where(c => c.Type == "scp").Select(c => c.Value).ToList();
                            var tokenRoles = jwtToken.Claims.Where(c => c.Type == "roles").Select(c => c.Value).ToList();

                            _logger.LogInformation("Token claims:");
                            foreach (var claim in jwtToken.Claims.Take(10))
                            {
                                _logger.LogInformation("- {ClaimType}: {ClaimValue}", claim.Type, claim.Value);
                            }

                            _logger.LogInformation("Token audience (aud): {Audience}", string.IsNullOrWhiteSpace(tokenAudience) ? "N/A" : tokenAudience);
                            _logger.LogInformation("Token app identifier (appid/azp): {AppId}", string.IsNullOrWhiteSpace(tokenAppId) ? "N/A" : tokenAppId);
                            _logger.LogInformation("Token scopes (scp): {Scopes}", tokenScopes.Count == 0 ? "N/A" : string.Join(", ", tokenScopes));
                            _logger.LogInformation("Token roles: {Roles}", tokenRoles.Count == 0 ? "N/A" : string.Join(", ", tokenRoles));

                            var claimsCount = jwtToken.Claims.Count();
                            if (claimsCount > 10)
                            {
                                _logger.LogInformation("... and {MoreClaims} more claims", claimsCount - 10);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogInformation("Token claims parsing skipped: {Message}", ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "FAILED: Failed to acquire token");
                    return new ObjectResult(new { error = "Failed to acquire access token", exception = ex.Message })
                    {
                        StatusCode = StatusCodes.Status500InternalServerError
                    };
                }

                _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                _logger.LogInformation("STEP 3: Resolving SharePoint Site");
                _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                try
                {
                    var siteUri = new Uri(siteUrl);
                    var relativePath = siteUri.AbsolutePath.Trim('/');
                    var siteLookupUrl = string.IsNullOrWhiteSpace(relativePath)
                        ? $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/"
                        : $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/{relativePath}";

                    _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                    _logger.LogInformation("TENANT CHECK: token tenant vs expected tenant");
                    _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                    var expectedTenantId = Environment.GetEnvironmentVariable("EXPECTED_TENANT_ID");
                    if (string.IsNullOrWhiteSpace(expectedTenantId))
                    {
                        expectedTenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
                    }

                    var sharePointTenantHint = siteUri.Host.Split('.')[0];
                    _logger.LogInformation("Token tid claim: {TokenTenantId}", string.IsNullOrWhiteSpace(tokenTenantId) ? "N/A" : tokenTenantId);
                    _logger.LogInformation("Expected tenant id (EXPECTED_TENANT_ID/AZURE_TENANT_ID): {ExpectedTenantId}", string.IsNullOrWhiteSpace(expectedTenantId) ? "N/A" : expectedTenantId);
                    _logger.LogInformation("SharePoint host tenant hint: {SharePointTenantHint}", sharePointTenantHint);

                    if (!string.IsNullOrWhiteSpace(expectedTenantId) && !string.IsNullOrWhiteSpace(tokenTenantId))
                    {
                        if (string.Equals(expectedTenantId, tokenTenantId, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation("SUCCESS: Tenant check passed: token tid matches expected tenant id");
                        }
                        else
                        {
                            _logger.LogWarning("FAILED: Tenant check mismatch: token tid does not match expected tenant id");
                            _logger.LogWarning("This can cause 401/403 on Graph SharePoint calls if identity is from another tenant.");
                        }
                    }
                    else
                    {
                        _logger.LogInformation("Tenant check is best-effort only. Set EXPECTED_TENANT_ID (or AZURE_TENANT_ID) to enable strict comparison.");
                    }

                    _logger.LogInformation("Site Host: {Host}", siteUri.Host);
                    _logger.LogInformation("Site Path: {Path}", relativePath);
                    _logger.LogInformation("Graph API lookup URL: {LookupUrl}", siteLookupUrl);

                    using (var client = _httpClientFactory.CreateClient())
                    using (var siteRequest = new HttpRequestMessage(HttpMethod.Get, siteLookupUrl))
                    {
                        siteRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                        using var siteResponse = await client.SendAsync(siteRequest, cancellationToken);
                        stopwatch.Stop();

                        _logger.LogInformation("Response status code: {StatusCode}", (int)siteResponse.StatusCode);
                        _logger.LogInformation("Response time: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);

                        if (siteResponse.IsSuccessStatusCode)
                        {
                            var siteContent = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
                            using var siteDoc = JsonDocument.Parse(siteContent);
                            var siteId = siteDoc.RootElement.TryGetProperty("id", out var idElement) ? idElement.GetString() : "N/A";
                            var siteName = siteDoc.RootElement.TryGetProperty("displayName", out var nameElement) ? nameElement.GetString() : "N/A";
                            var webUrl = siteDoc.RootElement.TryGetProperty("webUrl", out var urlElement) ? urlElement.GetString() : "N/A";

                            _logger.LogInformation("SUCCESS: Site resolved successfully:");
                            _logger.LogInformation("Site ID: {SiteId}", siteId);
                            _logger.LogInformation("Display Name: {SiteName}", siteName);
                            _logger.LogInformation("Web URL: {WebUrl}", webUrl);

                            _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                            _logger.LogInformation("STEP 4: Listing drives in the SharePoint site");
                            _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                            var drivesUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives";
                            _logger.LogInformation("Drives API endpoint: {DriveUrl}", drivesUrl);

                            using var drivesRequest = new HttpRequestMessage(HttpMethod.Get, drivesUrl);
                            drivesRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                            var driveStopwatch = System.Diagnostics.Stopwatch.StartNew();
                            using var drivesResponse = await client.SendAsync(drivesRequest, cancellationToken);
                            driveStopwatch.Stop();

                            _logger.LogInformation("Response status code: {StatusCode}", (int)drivesResponse.StatusCode);
                            _logger.LogInformation("Response time: {ElapsedMs}ms", driveStopwatch.ElapsedMilliseconds);

                            if (drivesResponse.IsSuccessStatusCode)
                            {
                                var drivesContent = await drivesResponse.Content.ReadAsStringAsync(cancellationToken);
                                using var drivesDoc = JsonDocument.Parse(drivesContent);
                                var drives = drivesDoc.RootElement.GetProperty("value").EnumerateArray().ToList();
                                _logger.LogInformation("SUCCESS: Found {DriveCount} drive(s):", drives.Count);

                                int driveIndex = 1;
                                foreach (var drive in drives)
                                {
                                    var driveId = drive.TryGetProperty("id", out var driveIdElement) ? driveIdElement.GetString() : "N/A";
                                    var driveName = drive.TryGetProperty("name", out var driveNameElement) ? driveNameElement.GetString() : "N/A";
                                    var driveType = drive.TryGetProperty("driveType", out var driveTypeElement) ? driveTypeElement.GetString() : "N/A";
                                    var owner = drive.TryGetProperty("owner", out var ownerElement)
                                        ? (ownerElement.TryGetProperty("user", out var userElement) && userElement.TryGetProperty("displayName", out var displayNameElement) ? displayNameElement.GetString() : "N/A")
                                        : "N/A";

                                    _logger.LogInformation("Drive {Index}:", driveIndex++);
                                    _logger.LogInformation("- ID: {DriveId}", driveId);
                                    _logger.LogInformation("- Name: {DriveName}", driveName);
                                    _logger.LogInformation("- Type: {DriveType}", driveType);
                                    _logger.LogInformation("- Owner: {Owner}", owner);
                                }
                            }
                            else
                            {
                                var drivesErrorContent = await drivesResponse.Content.ReadAsStringAsync(cancellationToken);
                                _logger.LogError("FAILED: Failed to list drives: {StatusCode} - {Content}", (int)drivesResponse.StatusCode, drivesErrorContent);
                            }

                            _logger.LogInformation("═══════════════════════════════════════════════════════════════════");
                            _logger.LogInformation("SUCCESS: SharePoint connection test completed");
                            _logger.LogInformation("═══════════════════════════════════════════════════════════════════");

                            await RunFederatedCredentialStepAsync(siteUrl, cancellationToken);

                            return new OkObjectResult(new
                            {
                                success = true,
                                siteId,
                                siteName,
                                webUrl,
                                message = "SharePoint connection test successful"
                            });
                        }

                        var errorContent = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
                        _logger.LogError("FAILED: Site resolution failed: {StatusCode}", (int)siteResponse.StatusCode);
                        _logger.LogError("Error response: {ErrorContent}", errorContent);

                        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                        _logger.LogInformation("STEP 3B: Trying SharePoint Online REST fallback");
                        _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                        var spoFallback = await TrySharePointOnlineFallbackAsync(
                            client,
                            tokenCredential,
                            siteUrl,
                            cancellationToken);

                        if (spoFallback.Success)
                        {
                            _logger.LogInformation("═══════════════════════════════════════════════════════════════════");
                            _logger.LogInformation("SUCCESS: SharePoint connection test completed via SPO fallback");
                            _logger.LogInformation("═══════════════════════════════════════════════════════════════════");

                            await RunFederatedCredentialStepAsync(siteUrl, cancellationToken);

                            return new OkObjectResult(new
                            {
                                success = true,
                                connectionMode = "spo-fallback",
                                siteId = spoFallback.SiteId,
                                siteName = spoFallback.SiteName,
                                webUrl = spoFallback.WebUrl,
                                message = "SharePoint connection test successful via SPO fallback"
                            });
                        }

                        string errorMessage = siteResponse.StatusCode switch
                        {
                            HttpStatusCode.Forbidden => "Access denied (403). The identity does not have permissions to access this site. Ensure the user assigned identity has 'Sites.Selected' permissions granted on the target SharePoint site.",
                            HttpStatusCode.NotFound => "Site not found (404). Verify the SharePoint site URL is correct and accessible.",
                            HttpStatusCode.Unauthorized => "Unauthorized (401). The token may be invalid or Graph API permissions are missing.",
                            _ => $"Failed with status {(int)siteResponse.StatusCode}"
                        };

                        _logger.LogError("═══════════════════════════════════════════════════════════════════");
                        _logger.LogError("FAILED: {ErrorMessage}", errorMessage);
                        _logger.LogError("═══════════════════════════════════════════════════════════════════");

                        await RunFederatedCredentialStepAsync(siteUrl, cancellationToken);

                        return new ObjectResult(new
                        {
                            success = false,
                            statusCode = (int)siteResponse.StatusCode,
                            error = errorMessage,
                            errorDetails = errorContent,
                            spoFallback = new
                            {
                                success = false,
                                statusCode = spoFallback.StatusCode,
                                error = spoFallback.Error,
                                errorDetails = spoFallback.ErrorDetails
                            }
                        })
                        {
                            StatusCode = (int)siteResponse.StatusCode
                        };
                    }
                }
                catch (UriFormatException ex)
                {
                    _logger.LogError(ex, "FAILED: Invalid SharePoint URL format");
                    _logger.LogError("═══════════════════════════════════════════════════════════════════");

                    return new BadRequestObjectResult(new
                    {
                        error = "Invalid SharePoint URL format",
                        exception = ex.Message
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "FAILED: Exception during SharePoint site resolution");
                    _logger.LogError("═══════════════════════════════════════════════════════════════════");

                    return new ObjectResult(new
                    {
                        error = "Exception during SharePoint testing",
                        exception = ex.Message,
                        stackTrace = ex.StackTrace
                    })
                    {
                        StatusCode = StatusCodes.Status500InternalServerError
                    };
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "FAILED: CRITICAL EXCEPTION in TestSharePointConnection");
            _logger.LogError("═══════════════════════════════════════════════════════════════════");

            return new ObjectResult(new
            {
                error = "Unexpected error",
                exception = ex.Message
            })
            {
                StatusCode = StatusCodes.Status500InternalServerError
            };
        }
    }

    private async Task RunFederatedCredentialStepAsync(
        string siteUrl,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
        _logger.LogInformation("STEP 5: Federated Credential auth (MSI -> App Registration -> SharePoint/Graph)");
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");

        var managedIdentityClientId = Environment.GetEnvironmentVariable("MANAGED_IDENTITY_CLIENT_ID");
        var appRegistrationClientId = Environment.GetEnvironmentVariable("APP_REGISTRATION_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        var sharePointTenant = Environment.GetEnvironmentVariable("SHAREPOINT_TENANT");
        var sharePointSitePath = Environment.GetEnvironmentVariable("SHAREPOINT_SITE_PATH");

        _logger.LogInformation("STEP 5.1: Input parameters and environment configuration");
        _logger.LogInformation("MANAGED_IDENTITY_CLIENT_ID: {ManagedIdentityClientId}", string.IsNullOrWhiteSpace(managedIdentityClientId) ? "<missing>" : managedIdentityClientId);
        _logger.LogInformation("APP_REGISTRATION_CLIENT_ID: {AppRegistrationClientId}", string.IsNullOrWhiteSpace(appRegistrationClientId) ? "<missing>" : appRegistrationClientId);
        _logger.LogInformation("AZURE_TENANT_ID: {TenantId}", string.IsNullOrWhiteSpace(tenantId) ? "<missing>" : tenantId);
        _logger.LogInformation("SHAREPOINT_TENANT: {SharePointTenant}", string.IsNullOrWhiteSpace(sharePointTenant) ? "<missing>" : sharePointTenant);
        _logger.LogInformation("SHAREPOINT_SITE_PATH: {SharePointSitePath}", string.IsNullOrWhiteSpace(sharePointSitePath) ? "<missing>" : sharePointSitePath);

        if ((string.IsNullOrWhiteSpace(sharePointTenant) || string.IsNullOrWhiteSpace(sharePointSitePath))
            && Uri.TryCreate(siteUrl, UriKind.Absolute, out var siteUri))
        {
            sharePointTenant ??= siteUri.Host;
            sharePointSitePath ??= siteUri.AbsolutePath.Trim('/');
            _logger.LogInformation("Derived fallback values from request siteUrl: tenant={Tenant}, path={Path}", sharePointTenant, sharePointSitePath);
        }

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
            _logger.LogError("✗ STEP 5 aborted. Missing required environment variable(s): {MissingVariables}", string.Join(", ", missingVariables));
            _logger.LogInformation("─────────────────────────────────────────────────────────────────");
            return;
        }

        var tokenExchangeScope = "api://AzureADTokenExchange/.default";
        var sharePointScope = $"https://{sharePointTenant}/.default";
        var graphScope = "https://graph.microsoft.com/.default";
        var normalizedSitePath = sharePointSitePath!.Trim('/');
        var graphSiteIdentifier = $"{sharePointTenant}:/{normalizedSitePath}";

        _logger.LogInformation("STEP 5.2: MSI token request for token exchange scope");
        _logger.LogInformation("Requested scope: {TokenExchangeScope}", tokenExchangeScope);

        var managedIdentityCredential = new ManagedIdentityCredential(managedIdentityClientId);
        AccessToken msiToken;

        try
        {
            var tokenStopwatch = System.Diagnostics.Stopwatch.StartNew();
            msiToken = await managedIdentityCredential.GetTokenAsync(
                new TokenRequestContext([tokenExchangeScope]),
                cancellationToken);
            tokenStopwatch.Stop();

            _logger.LogInformation("✓ MSI token acquired successfully");
            _logger.LogInformation("MSI token acquisition time: {ElapsedMs}ms", tokenStopwatch.ElapsedMilliseconds);
            _logger.LogInformation("MSI token expires at (UTC): {ExpiresOn:G}", msiToken.ExpiresOn);
            LogTokenClaims("MSI token", msiToken.Token, ["aud", "iss", "sub", "appid", "azp", "tid"]);
        }
        catch (Exception ex)
        {
            LogDetailedError("✗ Failed to acquire MSI token for token exchange", ex);
            _logger.LogInformation("─────────────────────────────────────────────────────────────────");
            return;
        }

        var assertionCallbackInvocations = 0;
        var cachedAssertionToken = msiToken;

        async Task<string> GetManagedIdentityAssertionAsync(CancellationToken assertionCancellationToken)
        {
            assertionCallbackInvocations++;
            _logger.LogInformation("STEP 5.3: Client assertion callback invocation #{Invocation}", assertionCallbackInvocations);

            if (cachedAssertionToken.ExpiresOn <= DateTimeOffset.UtcNow.AddMinutes(5))
            {
                _logger.LogInformation("Cached MSI assertion token expires soon; requesting a fresh MSI token");
                cachedAssertionToken = await managedIdentityCredential.GetTokenAsync(
                    new TokenRequestContext([tokenExchangeScope]),
                    assertionCancellationToken);
                _logger.LogInformation("✓ Refreshed MSI assertion token. New expiry (UTC): {ExpiresOn:G}", cachedAssertionToken.ExpiresOn);
                LogTokenClaims("Refreshed MSI token", cachedAssertionToken.Token, ["aud", "iss", "sub", "appid", "azp", "tid"]);
            }
            else
            {
                _logger.LogInformation("Reusing cached MSI assertion token. Expiry (UTC): {ExpiresOn:G}", cachedAssertionToken.ExpiresOn);
            }

            return cachedAssertionToken.Token;
        }

        var clientAssertionCredential = new ClientAssertionCredential(
            tenantId,
            appRegistrationClientId,
            GetManagedIdentityAssertionAsync);

        _logger.LogInformation("STEP 5.4: ClientAssertionCredential ready for App Registration federation");
        _logger.LogInformation("Tenant ID: {TenantId}", tenantId);
        _logger.LogInformation("App Registration Client ID: {AppRegistrationClientId}", appRegistrationClientId);
        _logger.LogInformation("SharePoint scope: {SharePointScope}", sharePointScope);
        _logger.LogInformation("Graph scope for validation: {GraphScope}", graphScope);

        AccessToken sharePointToken;
        try
        {
            var sharePointStopwatch = System.Diagnostics.Stopwatch.StartNew();
            sharePointToken = await clientAssertionCredential.GetTokenAsync(
                new TokenRequestContext([sharePointScope]),
                cancellationToken);
            sharePointStopwatch.Stop();

            _logger.LogInformation("✓ Final SharePoint token acquired successfully via federated credential");
            _logger.LogInformation("SharePoint token acquisition time: {ElapsedMs}ms", sharePointStopwatch.ElapsedMilliseconds);
            _logger.LogInformation("SharePoint token expires at (UTC): {ExpiresOn:G}", sharePointToken.ExpiresOn);
            LogTokenClaims("Final SharePoint token", sharePointToken.Token, ["aud", "roles", "appid", "azp", "tid"]);
        }
        catch (Exception ex)
        {
            LogDetailedError("✗ Failed to exchange MSI assertion for SharePoint token", ex);
            _logger.LogInformation("─────────────────────────────────────────────────────────────────");
            return;
        }

        try
        {
            _logger.LogInformation("STEP 5.5: Graph SDK validation call using the same federated credential");
            _logger.LogInformation("Graph validation target: /sites/{SiteIdentifier}", graphSiteIdentifier);
            _logger.LogInformation("Graph validation note: Graph requires a Graph audience token, so the SDK will request {GraphScope} using the same federated credential.", graphScope);

            var graphTokenStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var graphToken = await clientAssertionCredential.GetTokenAsync(
                new TokenRequestContext([graphScope]),
                cancellationToken);
            graphTokenStopwatch.Stop();

            _logger.LogInformation("✓ Graph token acquired successfully via federated credential");
            _logger.LogInformation("Graph token acquisition time: {ElapsedMs}ms", graphTokenStopwatch.ElapsedMilliseconds);
            _logger.LogInformation("Graph token expires at (UTC): {ExpiresOn:G}", graphToken.ExpiresOn);
            LogTokenClaims("Graph validation token", graphToken.Token, ["aud", "roles", "appid", "azp", "tid"]);

            var graphClient = new GraphServiceClient(clientAssertionCredential, [graphScope]);
            var requestStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var site = await graphClient.Sites[graphSiteIdentifier].GetAsync(cancellationToken: cancellationToken);
            requestStopwatch.Stop();

            var graphResponse = JsonSerializer.Serialize(site, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            _logger.LogInformation("✓ Graph validation call succeeded");
            _logger.LogInformation("Graph response time: {ElapsedMs}ms", requestStopwatch.ElapsedMilliseconds);
            _logger.LogInformation("Graph response payload: {GraphResponse}", graphResponse);
        }
        catch (ApiException ex)
        {
            _logger.LogError(ex, "✗ Graph validation call failed");
            _logger.LogError("Graph status code: {StatusCode}", ex.ResponseStatusCode);
            _logger.LogError("Graph error message: {Message}", ex.Message);
            if (ex.InnerException is not null)
            {
                _logger.LogError("Graph inner error: {InnerError}", ex.InnerException.Message);
            }
        }
        catch (Exception ex)
        {
            LogDetailedError("✗ Unexpected error during Graph validation call", ex);
        }

        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
    }

    private async Task<SpoFallbackResult> TrySharePointOnlineFallbackAsync(
        HttpClient client,
        TokenCredential tokenCredential,
        string siteUrl,
        CancellationToken cancellationToken)
    {
        try
        {
            var siteUri = new Uri(siteUrl);
            var spoScope = $"https://{siteUri.Host}/.default";
            _logger.LogInformation("SPO scope requested: {Scope}", spoScope);

            var tokenStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var spoToken = await tokenCredential.GetTokenAsync(new TokenRequestContext([spoScope]), cancellationToken);
            tokenStopwatch.Stop();

            _logger.LogInformation("SUCCESS: SPO token acquired successfully");
            _logger.LogInformation("SPO token acquisition time: {ElapsedMs}ms", tokenStopwatch.ElapsedMilliseconds);
            _logger.LogInformation("SPO token expires at (UTC): {ExpiresOn:G}", spoToken.ExpiresOn);

            var spoSiteUrl = $"{siteUrl.TrimEnd('/')}/_api/web?$select=Id,Title,Url";
            _logger.LogInformation("SPO REST endpoint: {SpoSiteUrl}", spoSiteUrl);

            using var spoRequest = new HttpRequestMessage(HttpMethod.Get, spoSiteUrl);
            spoRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", spoToken.Token);
            spoRequest.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

            var spoStopwatch = System.Diagnostics.Stopwatch.StartNew();
            using var spoResponse = await client.SendAsync(spoRequest, cancellationToken);
            spoStopwatch.Stop();

            _logger.LogInformation("SPO response status code: {StatusCode}", (int)spoResponse.StatusCode);
            _logger.LogInformation("SPO response time: {ElapsedMs}ms", spoStopwatch.ElapsedMilliseconds);

            var spoContent = await spoResponse.Content.ReadAsStringAsync(cancellationToken);
            if (!spoResponse.IsSuccessStatusCode)
            {
                _logger.LogError("FAILED: SPO site resolution failed: {StatusCode}", (int)spoResponse.StatusCode);
                _logger.LogError("SPO error response: {ErrorContent}", spoContent);
                return new SpoFallbackResult(
                    false,
                    (int)spoResponse.StatusCode,
                    null,
                    null,
                    null,
                    $"SPO fallback failed with status {(int)spoResponse.StatusCode}",
                    spoContent);
            }

            using var spoDoc = JsonDocument.Parse(spoContent);
            var siteId = spoDoc.RootElement.TryGetProperty("Id", out var spoIdElement) ? spoIdElement.GetString() : "N/A";
            var siteName = spoDoc.RootElement.TryGetProperty("Title", out var spoNameElement) ? spoNameElement.GetString() : "N/A";
            var webUrl = spoDoc.RootElement.TryGetProperty("Url", out var spoUrlElement) ? spoUrlElement.GetString() : siteUrl;

            _logger.LogInformation("SUCCESS: Site resolved successfully via SPO:");
            _logger.LogInformation("Site ID: {SiteId}", siteId);
            _logger.LogInformation("Display Name: {SiteName}", siteName);
            _logger.LogInformation("Web URL: {WebUrl}", webUrl);

            await LogDocumentLibrariesAsync(client, spoToken.Token, siteUrl, cancellationToken);

            return new SpoFallbackResult(true, 200, siteId, siteName, webUrl, null, null);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "FAILED: Exception during SPO fallback");
            return new SpoFallbackResult(false, 500, null, null, null, ex.Message, ex.ToString());
        }
    }

    private async Task LogDocumentLibrariesAsync(
        HttpClient client,
        string accessToken,
        string siteUrl,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
        _logger.LogInformation("STEP 4: Listing document libraries via SPO");
        _logger.LogInformation("─────────────────────────────────────────────────────────────────");

        var librariesUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseTemplate eq 101 and Hidden eq false";
        _logger.LogInformation("SPO libraries endpoint: {LibrariesUrl}", librariesUrl);

        using var librariesRequest = new HttpRequestMessage(HttpMethod.Get, librariesUrl);
        librariesRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        librariesRequest.Headers.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        using var librariesResponse = await client.SendAsync(librariesRequest, cancellationToken);
        stopwatch.Stop();

        _logger.LogInformation("SPO libraries response status code: {StatusCode}", (int)librariesResponse.StatusCode);
        _logger.LogInformation("SPO libraries response time: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);

        if (!librariesResponse.IsSuccessStatusCode)
        {
            var librariesErrorContent = await librariesResponse.Content.ReadAsStringAsync(cancellationToken);
            _logger.LogError("FAILED: Failed to list document libraries via SPO: {StatusCode} - {Content}", (int)librariesResponse.StatusCode, librariesErrorContent);
            return;
        }

        var librariesContent = await librariesResponse.Content.ReadAsStringAsync(cancellationToken);
        using var librariesDoc = JsonDocument.Parse(librariesContent);
        var libraries = librariesDoc.RootElement.GetProperty("value").EnumerateArray().ToList();
        _logger.LogInformation("SUCCESS: Found {LibraryCount} document library/libraries via SPO:", libraries.Count);

        int libraryIndex = 1;
        foreach (var library in libraries)
        {
            var libraryId = library.TryGetProperty("Id", out var libraryIdElement) ? libraryIdElement.GetString() : "N/A";
            var libraryName = library.TryGetProperty("Title", out var libraryNameElement) ? libraryNameElement.GetString() : "N/A";
            var libraryRoot = library.TryGetProperty("RootFolder", out var rootFolderElement)
                && rootFolderElement.TryGetProperty("ServerRelativeUrl", out var rootUrlElement)
                    ? rootUrlElement.GetString()
                    : "N/A";

            _logger.LogInformation("Library {Index}:", libraryIndex++);
            _logger.LogInformation("- ID: {LibraryId}", libraryId);
            _logger.LogInformation("- Name: {LibraryName}", libraryName);
            _logger.LogInformation("- RootFolder: {LibraryRoot}", libraryRoot);
        }
    }

    private void LogTokenClaims(string tokenLabel, string tokenValue, IReadOnlyCollection<string> primaryClaimTypes)
    {
        try
        {
            var tokenHandler = new JwtSecurityTokenHandler();
            if (!tokenHandler.CanReadToken(tokenValue))
            {
                _logger.LogWarning("Unable to decode {TokenLabel}: token format is not a readable JWT", tokenLabel);
                return;
            }

            var jwtToken = tokenHandler.ReadJwtToken(tokenValue);
            var claimsByType = jwtToken.Claims
                .GroupBy(claim => claim.Type)
                .ToDictionary(group => group.Key, group => group.Select(claim => claim.Value).ToList(), StringComparer.OrdinalIgnoreCase);

            _logger.LogInformation("{TokenLabel} claim dump start", tokenLabel);
            foreach (var claimType in primaryClaimTypes)
            {
                if (claimsByType.TryGetValue(claimType, out var values) && values.Count > 0)
                {
                    _logger.LogInformation("{TokenLabel} claim {ClaimType}: {ClaimValue}", tokenLabel, claimType, string.Join(", ", values));
                }
                else
                {
                    _logger.LogInformation("{TokenLabel} claim {ClaimType}: <missing>", tokenLabel, claimType);
                }
            }

            foreach (var claimGroup in claimsByType.OrderBy(group => group.Key))
            {
                _logger.LogInformation("{TokenLabel} raw claim {ClaimType}: {ClaimValue}", tokenLabel, claimGroup.Key, string.Join(", ", claimGroup.Value));
            }

            _logger.LogInformation("{TokenLabel} claim count: {ClaimCount}", tokenLabel, jwtToken.Claims.Count());
            _logger.LogInformation("{TokenLabel} claim dump end", tokenLabel);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to decode {TokenLabel} claims", tokenLabel);
        }
    }

    private void LogDetailedError(string message, Exception ex)
    {
        _logger.LogError(ex, message);
        _logger.LogError("Error type: {ErrorType}", ex.GetType().FullName);
        _logger.LogError("Error message: {ErrorMessage}", ex.Message);
        if (ex.InnerException is not null)
        {
            _logger.LogError("Inner error type: {InnerErrorType}", ex.InnerException.GetType().FullName);
            _logger.LogError("Inner error message: {InnerErrorMessage}", ex.InnerException.Message);
        }
    }

    private sealed record SpoFallbackResult(
        bool Success,
        int StatusCode,
        string? SiteId,
        string? SiteName,
        string? WebUrl,
        string? Error,
        string? ErrorDetails);
}
