using Azure.Core;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Models;

namespace SharePointSync.Functions.Services;

public sealed class SharePointSyncOrchestrator
{
    private readonly SharePointGraphClient _graphClient;
    private readonly BlobStorageSyncClient _blobClient;
    private readonly TokenCredential _tokenCredential;

    public SharePointSyncOrchestrator(
        SharePointGraphClient graphClient,
        BlobStorageSyncClient blobClient,
        TokenCredential tokenCredential)
    {
        _graphClient = graphClient;
        _blobClient = blobClient;
        _tokenCredential = tokenCredential;
    }

    public async Task<SyncStats> RunAsync(SyncOptions options, ILogger logger, CancellationToken cancellationToken)
    {
        var stats = new SyncStats();

        await _graphClient.InitializeAsync(options, cancellationToken);
        await _blobClient.InitializeAsync(options, _tokenCredential, cancellationToken);

        var (_, driveId) = _graphClient.GetResolvedIds();

        if (!options.ForceFullSync)
        {
            var deltaLink = await _blobClient.LoadDeltaTokenAsync(cancellationToken);
            await RunDeltaAsync(options, stats, deltaLink, logger, cancellationToken);
        }
        else
        {
            await RunFullAsync(options, stats, logger, cancellationToken);
        }

        if (options.SyncPermissions)
        {
            var allFiles = await _graphClient.ListFilesAsync(options.SharePointFolderPath, cancellationToken);
            foreach (var file in allFiles)
            {
                try
                {
                    var permissions = await _graphClient.GetFilePermissionsAsync(file.Id, file.Path, cancellationToken);
                    var blobName = _blobClient.GetBlobName(file.Path);
                    await _blobClient.UpdateBlobMetadataAsync(blobName, permissions.ToMetadata(), options.DryRun, cancellationToken);
                    stats.PermissionsSynced++;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to sync permissions for {Path}", file.Path);
                    stats.PermissionsFailed++;
                }
            }
        }

        logger.LogInformation(
            "Sync complete mode={Mode}, scanned={Scanned}, added={Added}, updated={Updated}, deleted={Deleted}, unchanged={Unchanged}, failed={Failed}, bytes={Bytes}, permissionsSynced={PermSynced}, permissionsFailed={PermFailed}",
            stats.SyncMode,
            stats.FilesScanned,
            stats.FilesAdded,
            stats.FilesUpdated,
            stats.FilesDeleted,
            stats.FilesUnchanged,
            stats.FilesFailed,
            stats.BytesTransferred,
            stats.PermissionsSynced,
            stats.PermissionsFailed);

        return stats;
    }

    private async Task RunDeltaAsync(
        SyncOptions options,
        SyncStats stats,
        string? deltaLink,
        ILogger logger,
        CancellationToken cancellationToken)
    {
        stats.SyncMode = string.IsNullOrWhiteSpace(deltaLink) ? "delta-initial" : "delta-incremental";

        var deltaResult = await _graphClient.GetDeltaAsync(deltaLink, cancellationToken);
        foreach (var change in deltaResult.Changes)
        {
            stats.FilesScanned++;

            if (change.ChangeType == DeltaChangeType.Deleted)
            {
                if (!options.DeleteOrphanedBlobs)
                {
                    continue;
                }

                try
                {
                    await _blobClient.DeleteBlobAsync(_blobClient.GetBlobName(change.ItemPath), options.DryRun, cancellationToken);
                    stats.FilesDeleted++;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to delete blob for deleted SharePoint item {Path}", change.ItemPath);
                    stats.FilesFailed++;
                }

                continue;
            }

            var file = change.File;
            if (file is null)
            {
                continue;
            }

            try
            {
                var content = await _graphClient.DownloadFileAsync(file.Id, cancellationToken);
                await _blobClient.UploadBlobAsync(
                    file.Path,
                    content,
                    file.Id,
                    file.LastModified,
                    file.ContentHash,
                    options.DryRun,
                    cancellationToken);

                stats.FilesAdded++;
                stats.BytesTransferred += content.LongLength;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to process changed file {Path}", file.Path);
                stats.FilesFailed++;
            }
        }

        await _blobClient.SaveDeltaTokenAsync(deltaResult.DeltaToken, options.DryRun, cancellationToken);
    }

    private async Task RunFullAsync(
        SyncOptions options,
        SyncStats stats,
        ILogger logger,
        CancellationToken cancellationToken)
    {
        stats.SyncMode = "full";

        var existingBlobs = await _blobClient.ListBlobsAsync(cancellationToken);
        var seenBlobNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var files = await _graphClient.ListFilesAsync(options.SharePointFolderPath, cancellationToken);

        foreach (var file in files)
        {
            stats.FilesScanned++;
            var blobName = _blobClient.GetBlobName(file.Path);
            seenBlobNames.Add(blobName);

            try
            {
                if (!existingBlobs.TryGetValue(blobName, out var existingBlob))
                {
                    var content = await _graphClient.DownloadFileAsync(file.Id, cancellationToken);
                    await _blobClient.UploadBlobAsync(file.Path, content, file.Id, file.LastModified, file.ContentHash, options.DryRun, cancellationToken);
                    stats.FilesAdded++;
                    stats.BytesTransferred += content.LongLength;
                }
                else if (_blobClient.ShouldUpdate(existingBlob, file.LastModified, file.ContentHash))
                {
                    var content = await _graphClient.DownloadFileAsync(file.Id, cancellationToken);
                    await _blobClient.UploadBlobAsync(file.Path, content, file.Id, file.LastModified, file.ContentHash, options.DryRun, cancellationToken);
                    stats.FilesUpdated++;
                    stats.BytesTransferred += content.LongLength;
                }
                else
                {
                    stats.FilesUnchanged++;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to process file {Path}", file.Path);
                stats.FilesFailed++;
            }
        }

        if (options.DeleteOrphanedBlobs)
        {
            foreach (var blobName in existingBlobs.Keys)
            {
                if (seenBlobNames.Contains(blobName))
                {
                    continue;
                }

                try
                {
                    await _blobClient.DeleteBlobAsync(blobName, options.DryRun, cancellationToken);
                    stats.FilesDeleted++;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to delete orphan blob {Blob}", blobName);
                    stats.FilesFailed++;
                }
            }
        }
    }
}
