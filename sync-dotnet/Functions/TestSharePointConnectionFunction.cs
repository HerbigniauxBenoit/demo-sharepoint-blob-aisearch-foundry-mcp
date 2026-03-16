using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Azure.Identity;

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
        
        // Log which identity is being used
        var userAssignedClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
        if (!string.IsNullOrWhiteSpace(userAssignedClientId))
        {
            _logger.LogInformation("🔐 IDENTITY: User Assigned Identity (Client ID: {ClientId})", userAssignedClientId);
        }
        else
        {
            _logger.LogInformation("🔐 IDENTITY: Default Azure Credential (auto-detection)");
            _logger.LogInformation("   Will try: Environment Vars → Managed Identity → Azure CLI → Visual Studio");
        }
        
        // 🔐 CREATE CONNECTION WITH USER ASSIGNED IDENTITY - RIGHT HERE IN THE FUNCTION
        TokenCredential tokenCredential;
        if (!string.IsNullOrWhiteSpace(userAssignedClientId))
        {
            // ✓ USE USER ASSIGNED IDENTITY WITH SPECIFIED CLIENT ID
            _logger.LogInformation("🔐 IDENTITY: User Assigned Identity (Client ID: {ClientId})", userAssignedClientId);
            tokenCredential = new ManagedIdentityCredential(userAssignedClientId);
        }
        else
        {
            // ✓ USE DEFAULT AZURE CREDENTIAL (for local development)
            _logger.LogInformation("🔐 IDENTITY: Default Azure Credential (auto-detection)");
            _logger.LogInformation("   Will try: Environment Vars → Managed Identity → Azure CLI → Visual Studio");
            tokenCredential = new DefaultAzureCredential();
        }

        try
        {
            // 1. Parse request body
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

                // 2. Get token
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

                    _logger.LogInformation("✓ Token acquired successfully");
                    _logger.LogInformation("  Token acquisition time: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
                    _logger.LogInformation("  Token expires at (UTC): {ExpiresOn:G}", token.ExpiresOn);

                    try
                    {
                        var jwtHandler = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
                        if (jwtHandler.CanReadToken(token.Token))
                        {
                            var jwtToken = jwtHandler.ReadJwtToken(token.Token);
                            tokenTenantId = jwtToken.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;
                            _logger.LogInformation("  Token claims:");
                            foreach (var claim in jwtToken.Claims.Take(10)) // Log first 10 claims
                            {
                                _logger.LogInformation("    - {ClaimType}: {ClaimValue}", claim.Type, claim.Value);
                            }
                            var claimsCount = jwtToken.Claims.Count();
                            if (claimsCount > 10)
                            {
                                _logger.LogInformation("    ... and {MoreClaims} more claims", claimsCount - 10);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogInformation("  (Token claims parsing skipped: {Message})", ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "✗ Failed to acquire token");
                    return new ObjectResult(new { error = "Failed to acquire access token", exception = ex.Message })
                    {
                        StatusCode = StatusCodes.Status500InternalServerError
                    };
                }

                // 3. Resolve SharePoint Site
                _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                _logger.LogInformation("STEP 3: Resolving SharePoint Site");
                _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                try
                {
                    var siteUri = new Uri(siteUrl);
                    var relativePath = siteUri.AbsolutePath.TrimStart('/');
                    var siteLookupUrl = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:/{Uri.EscapeDataString(relativePath)}";

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
                            _logger.LogInformation("✓ Tenant check passed: token tid matches expected tenant id");
                        }
                        else
                        {
                            _logger.LogWarning("✗ Tenant check mismatch: token tid does not match expected tenant id");
                            _logger.LogWarning("  This can cause 401/403 on Graph SharePoint calls if identity is from another tenant.");
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
                    {
                        using (var siteRequest = new HttpRequestMessage(HttpMethod.Get, siteLookupUrl))
                        {
                            siteRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                            using (var siteResponse = await client.SendAsync(siteRequest, cancellationToken))
                            {
                                stopwatch.Stop();

                                _logger.LogInformation("Response status code: {StatusCode}", (int)siteResponse.StatusCode);
                                _logger.LogInformation("Response time: {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);

                                if (siteResponse.IsSuccessStatusCode)
                                {
                                    var siteContent = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
                                    using (var siteDoc = JsonDocument.Parse(siteContent))
                                    {
                                        var siteId = siteDoc.RootElement.TryGetProperty("id", out var idElement) ? idElement.GetString() : "N/A";
                                        var siteName = siteDoc.RootElement.TryGetProperty("displayName", out var nameElement) ? nameElement.GetString() : "N/A";
                                        var webUrl = siteDoc.RootElement.TryGetProperty("webUrl", out var urlElement) ? urlElement.GetString() : "N/A";

                                        _logger.LogInformation("✓ Site resolved successfully:");
                                        _logger.LogInformation("  Site ID: {SiteId}", siteId);
                                        _logger.LogInformation("  Display Name: {SiteName}", siteName);
                                        _logger.LogInformation("  Web URL: {WebUrl}", webUrl);

                                        // 4. List drives in the site
                                        _logger.LogInformation("─────────────────────────────────────────────────────────────────");
                                        _logger.LogInformation("STEP 4: Listing drives in the SharePoint site");
                                        _logger.LogInformation("─────────────────────────────────────────────────────────────────");

                                        var drivesUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives";
                                        _logger.LogInformation("Drives API endpoint: {DriveUrl}", drivesUrl);

                                        using (var drivesRequest = new HttpRequestMessage(HttpMethod.Get, drivesUrl))
                                        {
                                            drivesRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                                            var driveStopwatch = System.Diagnostics.Stopwatch.StartNew();
                                            using (var drivesResponse = await client.SendAsync(drivesRequest, cancellationToken))
                                            {
                                                driveStopwatch.Stop();

                                                _logger.LogInformation("Response status code: {StatusCode}", (int)drivesResponse.StatusCode);
                                                _logger.LogInformation("Response time: {ElapsedMs}ms", driveStopwatch.ElapsedMilliseconds);

                                                if (drivesResponse.IsSuccessStatusCode)
                                                {
                                                    var drivesContent = await drivesResponse.Content.ReadAsStringAsync(cancellationToken);
                                                    using (var drivesDoc = JsonDocument.Parse(drivesContent))
                                                    {
                                                        var drives = drivesDoc.RootElement.GetProperty("value").EnumerateArray().ToList();
                                                        _logger.LogInformation("✓ Found {DriveCount} drive(s):", drives.Count);

                                                        int driveIndex = 1;
                                                        foreach (var drive in drives)
                                                        {
                                                            var driveId = drive.TryGetProperty("id", out var driveIdElement) ? driveIdElement.GetString() : "N/A";
                                                            var driveName = drive.TryGetProperty("name", out var driveNameElement) ? driveNameElement.GetString() : "N/A";
                                                            var driveType = drive.TryGetProperty("driveType", out var driveTypeElement) ? driveTypeElement.GetString() : "N/A";
                                                            var owner = drive.TryGetProperty("owner", out var ownerElement) 
                                                                ? (ownerElement.TryGetProperty("user", out var userElement) && userElement.TryGetProperty("displayName", out var displayNameElement) ? displayNameElement.GetString() : "N/A")
                                                                : "N/A";

                                                            _logger.LogInformation("  Drive {Index}:", driveIndex++);
                                                            _logger.LogInformation("    - ID: {DriveId}", driveId);
                                                            _logger.LogInformation("    - Name: {DriveName}", driveName);
                                                            _logger.LogInformation("    - Type: {DriveType}", driveType);
                                                            _logger.LogInformation("    - Owner: {Owner}", owner);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    var errorContent = await drivesResponse.Content.ReadAsStringAsync(cancellationToken);
                                                    _logger.LogError("✗ Failed to list drives: {StatusCode} - {Content}", (int)drivesResponse.StatusCode, errorContent);
                                                }
                                            }
                                        }

                                        // 5. Summary
                                        _logger.LogInformation("═══════════════════════════════════════════════════════════════════");
                                        _logger.LogInformation("✓ SUCCESS: SharePoint connection test completed");
                                        _logger.LogInformation("═══════════════════════════════════════════════════════════════════");

                                        return new OkObjectResult(new
                                        {
                                            success = true,
                                            siteId = siteId,
                                            siteName = siteName,
                                            webUrl = webUrl,
                                            message = "SharePoint connection test successful"
                                        });
                                    }
                                }
                                else
                                {
                                    var errorContent = await siteResponse.Content.ReadAsStringAsync(cancellationToken);
                                    _logger.LogError("✗ Site resolution failed: {StatusCode}", (int)siteResponse.StatusCode);
                                    _logger.LogError("Error response: {ErrorContent}", errorContent);

                                    string errorMessage = siteResponse.StatusCode switch
                                    {
                                        HttpStatusCode.Forbidden => "Access denied (403). The identity does not have permissions to access this site. Ensure the user assigned identity has 'Sites.Selected' permissions granted on the target SharePoint site.",
                                        HttpStatusCode.NotFound => "Site not found (404). Verify the SharePoint site URL is correct and accessible.",
                                        HttpStatusCode.Unauthorized => "Unauthorized (401). The token may be invalid or Graph API permissions are missing.",
                                        _ => $"Failed with status {(int)siteResponse.StatusCode}"
                                    };

                                    _logger.LogError("═══════════════════════════════════════════════════════════════════");
                                    _logger.LogError("✗ FAILED: {ErrorMessage}", errorMessage);
                                    _logger.LogError("═══════════════════════════════════════════════════════════════════");

                                    return new ObjectResult(new
                                    {
                                        success = false,
                                        statusCode = (int)siteResponse.StatusCode,
                                        error = errorMessage,
                                        errorDetails = errorContent
                                    })
                                    {
                                        StatusCode = (int)siteResponse.StatusCode
                                    };
                                }
                            }
                        }
                    }
                }
                catch (UriFormatException ex)
                {
                    _logger.LogError(ex, "✗ Invalid SharePoint URL format");
                    _logger.LogError("═══════════════════════════════════════════════════════════════════");
                    
                    return new BadRequestObjectResult(new
                    {
                        error = "Invalid SharePoint URL format",
                        exception = ex.Message
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "✗ Exception during SharePoint site resolution");
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
            _logger.LogError(ex, "✗ CRITICAL EXCEPTION in TestSharePointConnection");
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
}
