namespace SharePointSync.Functions.Services;

public sealed class SyncOptions
{
    public required string SharePointSiteUrl { get; init; }
    public required string SharePointDriveName { get; init; }
    public required string SharePointFolderPath { get; init; }
    public required string StorageAccountName { get; init; }
    public required string ContainerName { get; init; }
    public required string BlobPrefix { get; init; }
    public bool DeleteOrphanedBlobs { get; init; }
    public bool DryRun { get; init; }
    public bool SyncPermissions { get; init; }
    public bool ForceFullSync { get; init; }

    public string BlobAccountUrl => $"https://{StorageAccountName}.blob.core.windows.net";

    public static SyncOptions FromEnvironment()
    {
        return new SyncOptions
        {
            SharePointSiteUrl = Get("SHAREPOINT_SITE_URL"),
            SharePointDriveName = Get("SHAREPOINT_DRIVE_NAME", "Documents"),
            SharePointFolderPath = Get("SHAREPOINT_FOLDER_PATH", "/"),
            StorageAccountName = Get("AZURE_STORAGE_ACCOUNT_NAME"),
            ContainerName = Get("AZURE_BLOB_CONTAINER_NAME", "sharepoint-sync"),
            BlobPrefix = Get("AZURE_BLOB_PREFIX"),
            DeleteOrphanedBlobs = GetBool("DELETE_ORPHANED_BLOBS"),
            DryRun = GetBool("DRY_RUN"),
            SyncPermissions = GetBool("SYNC_PERMISSIONS"),
            ForceFullSync = GetBool("FORCE_FULL_SYNC")
        };
    }

    public void Validate()
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(SharePointSiteUrl))
        {
            errors.Add("SHAREPOINT_SITE_URL is required.");
        }

        if (string.IsNullOrWhiteSpace(StorageAccountName))
        {
            errors.Add("AZURE_STORAGE_ACCOUNT_NAME is required.");
        }

        if (string.IsNullOrWhiteSpace(ContainerName))
        {
            errors.Add("AZURE_BLOB_CONTAINER_NAME is required.");
        }

        if (errors.Count > 0)
        {
            throw new InvalidOperationException(string.Join(" ", errors));
        }
    }

    private static string Get(string key, string defaultValue = "")
    {
        var value = Environment.GetEnvironmentVariable(key);
        return string.IsNullOrWhiteSpace(value) ? defaultValue : value;
    }

    private static bool GetBool(string key)
    {
        var value = Environment.GetEnvironmentVariable(key);
        return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
