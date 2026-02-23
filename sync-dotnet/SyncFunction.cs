using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Services;

namespace SharePointSync.Functions;

public sealed class SyncFunction
{
    private readonly SharePointSyncOrchestrator _orchestrator;
    private readonly ILogger<SyncFunction> _logger;

    public SyncFunction(SharePointSyncOrchestrator orchestrator, ILogger<SyncFunction> logger)
    {
        _orchestrator = orchestrator;
        _logger = logger;
    }

    [Function("SharePointBlobSync")]
    public async Task RunAsync([TimerTrigger("%SYNC_SCHEDULE%", RunOnStartup = false)] object _, CancellationToken cancellationToken)
    {
        var options = SyncOptions.FromEnvironment();
        options.Validate();

        _logger.LogInformation(
            "Starting SharePoint sync site={Site}, drive={Drive}, folder={Folder}, container={Container}, dryRun={DryRun}, syncPermissions={SyncPermissions}, forceFullSync={ForceFullSync}",
            options.SharePointSiteUrl,
            options.SharePointDriveName,
            options.SharePointFolderPath,
            options.ContainerName,
            options.DryRun,
            options.SyncPermissions,
            options.ForceFullSync);

        await _orchestrator.RunAsync(options, _logger, cancellationToken);
    }
}
