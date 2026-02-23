using System.Text.Json;
using Azure;
using Azure.Core;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using SharePointSync.Functions.Models;

namespace SharePointSync.Functions.Services;

public sealed class BlobStorageSyncClient
{
    public const string DeltaTokenBlobName = ".sync-state/delta-token.json";
    public const string MetadataSpItemId = "sharepoint_item_id";
    public const string MetadataSpLastModified = "sharepoint_last_modified";
    public const string MetadataSpContentHash = "sharepoint_content_hash";

    private BlobContainerClient? _containerClient;
    private string _blobPrefix = string.Empty;

    public async Task InitializeAsync(SyncOptions options, TokenCredential credential, CancellationToken cancellationToken)
    {
        var service = new BlobServiceClient(new Uri(options.BlobAccountUrl), credential);
        _containerClient = service.GetBlobContainerClient(options.ContainerName);
        _blobPrefix = options.BlobPrefix.Trim('/');

        await _containerClient.CreateIfNotExistsAsync(cancellationToken: cancellationToken);
    }

    public string GetBlobName(string sharePointPath)
    {
        var clean = sharePointPath.TrimStart('/');
        return string.IsNullOrWhiteSpace(_blobPrefix) ? clean : $"{_blobPrefix}/{clean}";
    }

    public async Task<IReadOnlyDictionary<string, BlobFile>> ListBlobsAsync(CancellationToken cancellationToken)
    {
        EnsureInitialized();
        var result = new Dictionary<string, BlobFile>(StringComparer.OrdinalIgnoreCase);
        var prefix = string.IsNullOrWhiteSpace(_blobPrefix) ? null : _blobPrefix;

        await foreach (var blob in _containerClient!.GetBlobsAsync(traits: BlobTraits.Metadata, prefix: prefix, cancellationToken: cancellationToken))
        {
            if (blob.Name.EndsWith('/'))
            {
                continue;
            }

            if (blob.Properties.ContentLength == 0 && !blob.Name.Split('/').Last().Contains('.'))
            {
                continue;
            }

            result[blob.Name] = new BlobFile
            {
                Name = blob.Name,
                Size = blob.Properties.ContentLength ?? 0,
                LastModified = blob.Properties.LastModified ?? DateTimeOffset.UtcNow,
                Metadata = blob.Metadata ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            };
        }

        return result;
    }

    public async Task UploadBlobAsync(
        string sharePointPath,
        byte[] content,
        string sharePointItemId,
        DateTimeOffset? sharePointLastModified,
        string? sharePointContentHash,
        bool dryRun,
        CancellationToken cancellationToken)
    {
        EnsureInitialized();
        if (dryRun)
        {
            return;
        }

        var blobName = GetBlobName(sharePointPath);
        var blob = _containerClient!.GetBlobClient(blobName);

        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [MetadataSpItemId] = sharePointItemId,
            [MetadataSpLastModified] = (sharePointLastModified ?? DateTimeOffset.UtcNow).ToString("O")
        };

        if (!string.IsNullOrWhiteSpace(sharePointContentHash))
        {
            metadata[MetadataSpContentHash] = sharePointContentHash;
        }

        await blob.UploadAsync(BinaryData.FromBytes(content), new BlobUploadOptions
        {
            Metadata = metadata
        }, cancellationToken);
    }

    public bool ShouldUpdate(BlobFile existingBlob, DateTimeOffset? sharePointLastModified, string? sharePointContentHash)
    {
        if (existingBlob.Metadata is null)
        {
            return true;
        }

        if (existingBlob.Metadata.TryGetValue(MetadataSpContentHash, out var storedHash) &&
            !string.IsNullOrWhiteSpace(storedHash) &&
            !string.IsNullOrWhiteSpace(sharePointContentHash) &&
            !string.Equals(storedHash, sharePointContentHash, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (sharePointLastModified is null)
        {
            return true;
        }

        if (existingBlob.Metadata.TryGetValue(MetadataSpLastModified, out var storedDate) &&
            DateTimeOffset.TryParse(storedDate, out var parsedStoredDate))
        {
            return sharePointLastModified > parsedStoredDate;
        }

        return true;
    }

    public async Task DeleteBlobAsync(string blobName, bool dryRun, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        if (dryRun)
        {
            return;
        }

        var blob = _containerClient!.GetBlobClient(blobName);
        try
        {
            await blob.DeleteIfExistsAsync(cancellationToken: cancellationToken);
        }
        catch (RequestFailedException ex) when (string.Equals(ex.ErrorCode, "DirectoryIsNotEmpty", StringComparison.OrdinalIgnoreCase))
        {
            await DeleteDirectoryRecursiveAsync(blobName, cancellationToken);
        }
    }

    public async Task UpdateBlobMetadataAsync(string blobName, IDictionary<string, string> additionalMetadata, bool dryRun, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        if (dryRun)
        {
            return;
        }

        var blob = _containerClient!.GetBlobClient(blobName);
        IDictionary<string, string> currentMetadata;

        try
        {
            var props = await blob.GetPropertiesAsync(cancellationToken: cancellationToken);
            currentMetadata = new Dictionary<string, string>(props.Value.Metadata, StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            currentMetadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        foreach (var key in new[] { "metadata_user_ids", "metadata_group_ids", "acl_user_ids_list", "acl_group_ids_list", "metadata_acl_user_ids", "metdata_acl_group_ids" })
        {
            currentMetadata.Remove(key);
        }

        foreach (var (key, value) in additionalMetadata)
        {
            currentMetadata[key] = value;
        }

        await blob.SetMetadataAsync(currentMetadata, cancellationToken: cancellationToken);
    }

    public async Task<string?> LoadDeltaTokenAsync(CancellationToken cancellationToken)
    {
        EnsureInitialized();
        var blob = _containerClient!.GetBlobClient(DeltaTokenBlobName);

        try
        {
            var content = await blob.DownloadContentAsync(cancellationToken);
            using var document = JsonDocument.Parse(content.Value.Content);
            return document.RootElement.TryGetProperty("delta_link", out var deltaLink)
                ? deltaLink.GetString()
                : null;
        }
        catch
        {
            return null;
        }
    }

    public async Task SaveDeltaTokenAsync(string deltaToken, bool dryRun, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        if (dryRun || string.IsNullOrWhiteSpace(deltaToken))
        {
            return;
        }

        var payload = JsonSerializer.Serialize(new
        {
            delta_link = deltaToken,
            saved_at = DateTimeOffset.UtcNow.ToString("O")
        });

        await _containerClient!
            .GetBlobClient(DeltaTokenBlobName)
            .UploadAsync(BinaryData.FromString(payload), overwrite: true, cancellationToken: cancellationToken);
    }

    private async Task DeleteDirectoryRecursiveAsync(string directoryPath, CancellationToken cancellationToken)
    {
        var prefix = directoryPath.TrimEnd('/') + "/";
        await foreach (var blob in _containerClient!.GetBlobsAsync(prefix: prefix, cancellationToken: cancellationToken))
        {
            await _containerClient.GetBlobClient(blob.Name).DeleteIfExistsAsync(cancellationToken: cancellationToken);
        }

        await _containerClient.GetBlobClient(directoryPath).DeleteIfExistsAsync(cancellationToken: cancellationToken);
    }

    private void EnsureInitialized()
    {
        if (_containerClient is null)
        {
            throw new InvalidOperationException("BlobStorageSyncClient.InitializeAsync must be called before usage.");
        }
    }
}
