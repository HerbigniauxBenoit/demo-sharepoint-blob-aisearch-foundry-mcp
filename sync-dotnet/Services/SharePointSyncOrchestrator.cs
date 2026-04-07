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

        await RunFullAsync(options, stats, logger, cancellationToken);

        if (options.SyncPermissions)
        {
            var allFiles = await _graphClient.ListFilesAsync(options.SharePointFolderPath, cancellationToken);
            foreach (var file in allFiles)
            {
                try
                {
                    var permissions = await _graphClient.GetFilePermissionsAsync(file.Id, file.Path, cancellationToken);
                    var blobName = _blobClient.GetBlobName(file.Path);
                    await _blobClient.UpdateBlobMetadataAsync(blobName, permissions.ToMetadata(), cancellationToken);
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
            "Sync complete mode={Mode}, scanned={Scanned}, skippedTooLarge={SkippedTooLarge}, added={Added}, updated={Updated}, deleted={Deleted}, unchanged={Unchanged}, failed={Failed}, bytes={Bytes}, permissionsSynced={PermSynced}, permissionsFailed={PermFailed}",
            stats.SyncMode,
            stats.FilesScanned,
            stats.FilesSkippedTooLarge,
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

            if (file.Size > options.MaxFileSizeBytes)
            {
                logger.LogWarning(
                    "Skipping file {Path} because size {SizeBytes} exceeds MAX_FILE_SIZE_MB ({MaxFileSizeMb} MB).",
                    file.Path,
                    file.Size,
                    options.MaxFileSizeMb);
                stats.FilesSkippedTooLarge++;
                continue;
            }

            var blobName = _blobClient.GetBlobName(file.Path);
            seenBlobNames.Add(blobName);

            try
            {
                if (!existingBlobs.TryGetValue(blobName, out var existingBlob))
                {
                    var content = await _graphClient.DownloadFileAsync(file.Id, cancellationToken);
                    await _blobClient.UploadBlobAsync(file.Path, content, file.Id, file.LastModified, file.ContentHash, cancellationToken);
                    stats.FilesAdded++;
                    stats.BytesTransferred += content.LongLength;
                }
                else if (_blobClient.ShouldUpdate(existingBlob, file.LastModified, file.ContentHash))
                {
                    var content = await _graphClient.DownloadFileAsync(file.Id, cancellationToken);
                    await _blobClient.UploadBlobAsync(file.Path, content, file.Id, file.LastModified, file.ContentHash, cancellationToken);
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
                    await _blobClient.DeleteBlobAsync(blobName, cancellationToken);
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
