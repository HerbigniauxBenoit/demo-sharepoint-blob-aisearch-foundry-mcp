using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace SharePointSync.Functions;

public sealed class VerifySitesSelectedFunction
{
    private readonly ILogger<VerifySitesSelectedFunction> _logger;
    private readonly IHttpClientFactory _httpClientFactory;

    public VerifySitesSelectedFunction(
        ILogger<VerifySitesSelectedFunction> logger,
        IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    [Function("VerifySitesSelected")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "graph/verify-sites-selected")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation("Verifying Sites.Selected scope on Managed Identity...");

        try
        {
            // 1. Get token for Microsoft Graph using DefaultAzureCredential
            const string graphScope = "https://graph.microsoft.com/.default";
            var credential = new DefaultAzureCredential();
            var tokenResponse = await credential.GetTokenAsync(
                new TokenRequestContext([graphScope]),
                cancellationToken);
            var rawToken = tokenResponse.Token;

            _logger.LogInformation("✓ Token acquired for scope {Scope}", graphScope);

            // 2. Decode the JWT to inspect roles/scopes
            var handler = new JwtSecurityTokenHandler();
            if (!handler.CanReadToken(rawToken))
            {
                _logger.LogError("✗ Unable to decode the acquired token.");
                var badResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await badResponse.WriteAsJsonAsync(new { error = "Unable to decode token" }, cancellationToken);
                return badResponse;
            }

            var jwt = handler.ReadJwtToken(rawToken);

            // 3. Extract claims mirroring JS payload fields
            var appId = jwt.Claims.FirstOrDefault(c => c.Type == "appid")?.Value
                     ?? jwt.Claims.FirstOrDefault(c => c.Type == "azp")?.Value;
            var objectId = jwt.Claims.FirstOrDefault(c => c.Type == "oid")?.Value;
            var tenantId = jwt.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;
            var audience = jwt.Claims.FirstOrDefault(c => c.Type == "aud")?.Value;
            var issuer = jwt.Claims.FirstOrDefault(c => c.Type == "iss")?.Value;
            var exp = jwt.Claims.FirstOrDefault(c => c.Type == "exp")?.Value;

            DateTimeOffset? expiresAt = long.TryParse(exp, out var expUnix)
                ? DateTimeOffset.FromUnixTimeSeconds(expUnix)
                : null;

            var roles = jwt.Claims
                .Where(c => c.Type == "roles")
                .Select(c => c.Value)
                .ToList();

            var hasSitesSelected = roles.Contains("Sites.Selected");

            _logger.LogInformation(
                "Token roles: {Roles}",
                roles.Count > 0 ? string.Join(", ", roles) : "(none)");
            _logger.LogInformation(
                "Sites.Selected present: {HasSitesSelected}", hasSitesSelected);

            // 4. Test actual Graph API call — list sites
            object graphTestResult;
            try
            {
                using var client = _httpClientFactory.CreateClient();
                using var graphRequest = new HttpRequestMessage(
                    HttpMethod.Get,
                    "https://graph.microsoft.com/v1.0/sites?search=*&$top=5&$select=id,displayName,webUrl");
                graphRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", rawToken);

                using var graphResponse = await client.SendAsync(graphRequest, cancellationToken);
                var body = await graphResponse.Content.ReadAsStringAsync(cancellationToken);

                _logger.LogInformation(
                    "Graph /sites response: HTTP {StatusCode}", (int)graphResponse.StatusCode);

                List<object> sites = [];
                string? graphError = null;
                int sitesReturned = 0;

                using var doc = JsonDocument.Parse(body);
                if (doc.RootElement.TryGetProperty("value", out var valueElement)
                    && valueElement.ValueKind == JsonValueKind.Array)
                {
                    sitesReturned = valueElement.GetArrayLength();
                    foreach (var site in valueElement.EnumerateArray())
                    {
                        sites.Add(new
                        {
                            displayName = site.TryGetProperty("displayName", out var dn) ? dn.GetString() : null,
                            webUrl = site.TryGetProperty("webUrl", out var wu) ? wu.GetString() : null
                        });
                    }
                }

                if (doc.RootElement.TryGetProperty("error", out var errorElement)
                    && errorElement.TryGetProperty("message", out var msgElement))
                {
                    graphError = msgElement.GetString();
                }

                graphTestResult = new
                {
                    status = (int)graphResponse.StatusCode,
                    sitesReturned,
                    sites,
                    error = graphError
                };
            }
            catch (Exception ex)
            {
                graphTestResult = new { error = ex.Message };
            }

            var result = new
            {
                managedIdentity = new
                {
                    appId,
                    objectId,
                    tenantId
                },
                token = new
                {
                    audience,
                    issuer,
                    expiresAt = expiresAt?.ToString("o"),
                    allRoles = roles
                },
                verification = new
                {
                    hasSitesSelected,
                    sitesSelectedFound = hasSitesSelected
                        ? "Sites.Selected IS present in token roles"
                        : "Sites.Selected NOT found — check app role assignment"
                },
                graphApiTest = graphTestResult
            };

            var ok = req.CreateResponse(HttpStatusCode.OK);
            await ok.WriteAsJsonAsync(result, cancellationToken);
            return ok;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "✗ VerifySitesSelected failed: {Message}", ex.Message);

            var error = req.CreateResponse(HttpStatusCode.InternalServerError);
            await error.WriteAsJsonAsync(new
            {
                success = false,
                error = ex.Message
            }, cancellationToken);
            return error;
        }
    }
}
