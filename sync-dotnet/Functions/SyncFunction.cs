using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Services;
using System.Net;
using System.Text.Json;
using System.Web;

namespace SharePointSync.Functions;

public sealed class SyncFunction
{
    private readonly SharePointSyncOrchestrator _orchestrator;
    private readonly ILogger<SyncFunction> _logger;
    private static readonly SemaphoreSlim SyncLock = new(1, 1);

    public SyncFunction(SharePointSyncOrchestrator orchestrator, ILogger<SyncFunction> logger)
    {
        _orchestrator = orchestrator;
        _logger = logger;
    }

    [Function("SharePointBlobSync")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = "sharepoint-sync")] HttpRequestData req,
        CancellationToken cancellationToken)
    {
        SyncOptions options;
        Dictionary<string, string> appliedOverrides;
        try
        {
            var body = await ParseRequestBodyAsync(req, cancellationToken);
            var overrides = BuildOverrides(req, body);
            appliedOverrides = overrides;
            options = ApplyOverrides(SyncOptions.FromEnvironment(), overrides);
            options.Validate();
        }
        catch (Exception ex) when (ex is InvalidOperationException || ex is ArgumentException || ex is JsonException)
        {
            return await CreateJsonResponseAsync(
                req,
                HttpStatusCode.BadRequest,
                new
                {
                    status = "error",
                    message = ex.Message,
                    allowed_values = new[] { "true", "false", "1", "0", "yes", "no" }
                },
                cancellationToken);
        }

        _logger.LogInformation(
            "HTTP trigger received for SharePoint sync site={Site}, drive={Drive}, folder={Folder}, container={Container}, dryRun={DryRun}, syncPermissions={SyncPermissions}, forceFullSync={ForceFullSync}, overrides={Overrides}",
            options.SharePointSiteUrl,
            options.SharePointDriveName,
            options.SharePointFolderPath,
            options.ContainerName,
            options.DryRun,
            options.SyncPermissions,
            options.ForceFullSync,
            string.Join(",", appliedOverrides.Select(kv => $"{kv.Key}={kv.Value}")));

        await SyncLock.WaitAsync(cancellationToken);
        try
        {
            await _orchestrator.RunAsync(options, _logger, cancellationToken);
            return await CreateJsonResponseAsync(
                req,
                HttpStatusCode.OK,
                new
                {
                    status = "ok",
                    exit_code = 0,
                    applied_overrides = appliedOverrides
                },
                cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "SharePoint sync failed");
            return await CreateJsonResponseAsync(
                req,
                HttpStatusCode.InternalServerError,
                new
                {
                    status = "error",
                    exit_code = 1,
                    message = ex.Message,
                    applied_overrides = appliedOverrides
                },
                cancellationToken);
        }
        finally
        {
            SyncLock.Release();
        }
    }

    private static async Task<HttpResponseData> CreateJsonResponseAsync(
        HttpRequestData req,
        HttpStatusCode statusCode,
        object payload,
        CancellationToken cancellationToken)
    {
        var response = req.CreateResponse(statusCode);
        response.Headers.Add("Content-Type", "application/json");
        var json = JsonSerializer.Serialize(payload);
        await response.WriteStringAsync(json, cancellationToken);
        return response;
    }

    private static async Task<Dictionary<string, string>> ParseRequestBodyAsync(HttpRequestData req, CancellationToken cancellationToken)
    {
        if (req.Body is null || !req.Body.CanRead)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        using var reader = new StreamReader(req.Body);
        var bodyText = await reader.ReadToEndAsync(cancellationToken);
        if (string.IsNullOrWhiteSpace(bodyText))
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(bodyText)
            ?? new Dictionary<string, JsonElement>();

        return values.ToDictionary(
            kv => kv.Key,
            kv => kv.Value.ValueKind == JsonValueKind.String ? kv.Value.GetString() ?? string.Empty : kv.Value.ToString(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, string> BuildOverrides(HttpRequestData req, Dictionary<string, string> bodyValues)
    {
        var query = HttpUtility.ParseQueryString(req.Url.Query);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        AddBooleanOverride("force_full_sync", "FORCE_FULL_SYNC");
        AddBooleanOverride("dry_run", "DRY_RUN");
        AddBooleanOverride("delete_orphaned_blobs", "DELETE_ORPHANED_BLOBS");
        AddBooleanOverride("sync_permissions", "SYNC_PERMISSIONS");

        return result;

        void AddBooleanOverride(string requestKey, string environmentKey)
        {
            var value = query[requestKey];
            if (string.IsNullOrWhiteSpace(value) && bodyValues.TryGetValue(requestKey, out var bodyValue))
            {
                value = bodyValue;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            result[environmentKey] = ParseBoolean(value).ToString().ToLowerInvariant();
        }
    }

    private static SyncOptions ApplyOverrides(SyncOptions options, Dictionary<string, string> overrides)
    {
        bool? forceFullSync = TryGetBoolOverride(overrides, "FORCE_FULL_SYNC");
        bool? dryRun = TryGetBoolOverride(overrides, "DRY_RUN");
        bool? deleteOrphaned = TryGetBoolOverride(overrides, "DELETE_ORPHANED_BLOBS");
        bool? syncPermissions = TryGetBoolOverride(overrides, "SYNC_PERMISSIONS");

        return new SyncOptions
        {
            SharePointSiteUrl = options.SharePointSiteUrl,
            SharePointDriveName = options.SharePointDriveName,
            SharePointFolderPath = options.SharePointFolderPath,
            StorageAccountName = options.StorageAccountName,
            ContainerName = options.ContainerName,
            BlobPrefix = options.BlobPrefix,
            DeleteOrphanedBlobs = deleteOrphaned ?? options.DeleteOrphanedBlobs,
            DryRun = dryRun ?? options.DryRun,
            SyncPermissions = syncPermissions ?? options.SyncPermissions,
            ForceFullSync = forceFullSync ?? options.ForceFullSync
        };
    }

    private static bool? TryGetBoolOverride(Dictionary<string, string> overrides, string key)
    {
        if (!overrides.TryGetValue(key, out var value))
        {
            return null;
        }

        return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ParseBoolean(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        return normalized switch
        {
            "1" or "true" or "yes" or "y" or "on" => true,
            "0" or "false" or "no" or "n" or "off" => false,
            _ => throw new ArgumentException($"Invalid boolean value '{value}'")
        };
    }
}
