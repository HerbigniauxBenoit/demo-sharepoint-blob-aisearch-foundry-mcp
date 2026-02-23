using System.Text.Json;

namespace SharePointSync.Functions.Models;

public sealed class SyncStats
{
    public int FilesScanned { get; set; }
    public int FilesAdded { get; set; }
    public int FilesUpdated { get; set; }
    public int FilesDeleted { get; set; }
    public int FilesUnchanged { get; set; }
    public int FilesFailed { get; set; }
    public long BytesTransferred { get; set; }
    public int PermissionsSynced { get; set; }
    public int PermissionsFailed { get; set; }
    public string SyncMode { get; set; } = "full";
}

public sealed class SharePointFile
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required string Path { get; init; }
    public long Size { get; init; }
    public DateTimeOffset? LastModified { get; init; }
    public string? ContentHash { get; init; }
}

public enum DeltaChangeType
{
    CreatedOrModified,
    Deleted
}

public sealed class DeltaChange
{
    public required DeltaChangeType ChangeType { get; init; }
    public SharePointFile? File { get; init; }
    public required string ItemId { get; init; }
    public required string ItemName { get; init; }
    public required string ItemPath { get; init; }
    public bool IsFolder { get; init; }
}

public sealed class DeltaResult
{
    public List<DeltaChange> Changes { get; init; } = [];
    public string DeltaToken { get; init; } = string.Empty;
    public bool IsInitialSync { get; init; }
}

public sealed class BlobFile
{
    public required string Name { get; init; }
    public long Size { get; init; }
    public DateTimeOffset LastModified { get; init; }
    public IDictionary<string, string> Metadata { get; init; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}

public sealed class FilePermissions
{
    public required string FileId { get; init; }
    public required string FilePath { get; init; }
    public required List<SharePointPermission> Permissions { get; init; }

    public IDictionary<string, string> ToMetadata()
    {
        var users = Permissions
            .Where(p => p.IdentityType == "user" && IsGuid(p.IdentityId))
            .Select(p => p.IdentityId!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var groups = Permissions
            .Where(p => p.IdentityType == "group" && IsGuid(p.IdentityId))
            .Select(p => p.IdentityId!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["sharepoint_permissions"] = JsonSerializer.Serialize(Permissions),
            ["permissions_synced_at"] = DateTimeOffset.UtcNow.ToString("O"),
            ["user_ids"] = users.Length == 0 ? "00000000-0000-0000-0000-000000000000" : string.Join("|", users),
            ["group_ids"] = groups.Length == 0 ? "00000000-0000-0000-0000-000000000001" : string.Join("|", groups)
        };
    }

    private static bool IsGuid(string? value) => Guid.TryParse(value, out _);
}

public sealed class SharePointPermission
{
    public required string Id { get; init; }
    public required string IdentityType { get; init; }
    public required string DisplayName { get; init; }
    public string? IdentityId { get; init; }
    public string? Email { get; init; }
    public required string[] Roles { get; init; }
    public bool Inherited { get; init; }
}
