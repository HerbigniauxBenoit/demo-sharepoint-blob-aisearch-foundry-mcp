using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Services;

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
    public async Task RunTimerAsync(
        [TimerTrigger("%SYNC_SCHEDULE%")]
        TimerInfo timer,
        CancellationToken cancellationToken)
    {
        _logger.LogInformation(
            "Timer trigger fired for SharePoint sync. IsPastDue={IsPastDue}",
            timer.IsPastDue);

        SyncOptions options;
        try
        {
            options = SyncOptions.FromEnvironment();
            options.Validate();
        }
        catch (Exception ex) when (ex is InvalidOperationException || ex is ArgumentException)
        {
            _logger.LogError(ex, "Invalid configuration for scheduled SharePoint sync: {Message}", ex.Message);
            return;
        }

        _logger.LogInformation(
            "Scheduled SharePoint sync starting. site={Site}, drive={Drive}, folder={Folder}, container={Container}, syncPermissions={SyncPermissions}",
            options.SharePointSiteUrl,
            options.SharePointDriveName,
            options.SharePointFolderPath,
            options.ContainerName,
            options.SyncPermissions);

        await SyncLock.WaitAsync(cancellationToken);
        try
        {
            await _orchestrator.RunAsync(options, _logger, cancellationToken);
            _logger.LogInformation("Scheduled SharePoint sync completed successfully.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Scheduled SharePoint sync failed: {Message}", ex.Message);
            throw;
        }
        finally
        {
            SyncLock.Release();
        }
    }
}
