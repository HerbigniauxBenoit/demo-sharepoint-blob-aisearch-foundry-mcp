using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using SharePointSync.Functions.Models;

namespace SharePointSync.Functions.Services;

public sealed class SharePointGraphClient
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly TokenCredential _credential;
    private readonly string[] _scopes = ["https://graph.microsoft.com/.default"];

    private string? _siteId;
    private string? _driveId;

    public SharePointGraphClient(IHttpClientFactory httpClientFactory, TokenCredential credential)
    {
        _httpClientFactory = httpClientFactory;
        _credential = credential;
    }

    public (string SiteId, string DriveId) GetResolvedIds()
    {
        if (string.IsNullOrWhiteSpace(_siteId) || string.IsNullOrWhiteSpace(_driveId))
        {
            throw new InvalidOperationException("Graph IDs are not resolved.");
        }

        return (_siteId, _driveId);
    }

    public async Task InitializeAsync(SyncOptions options, CancellationToken cancellationToken)
    {
        var siteUri = new Uri(options.SharePointSiteUrl);
        var siteLookup = $"https://graph.microsoft.com/v1.0/sites/{siteUri.Host}:{siteUri.AbsolutePath}";
        var siteDoc = await GetJsonAsync(siteLookup, cancellationToken);
        _siteId = siteDoc.RootElement.GetProperty("id").GetString();

        if (string.IsNullOrWhiteSpace(_siteId))
        {
            throw new InvalidOperationException($"Unable to resolve site from {options.SharePointSiteUrl}");
        }

        var drivesDoc = await GetJsonAsync($"https://graph.microsoft.com/v1.0/sites/{_siteId}/drives", cancellationToken);
        foreach (var drive in drivesDoc.RootElement.GetProperty("value").EnumerateArray())
        {
            var driveName = drive.GetProperty("name").GetString();
            if (string.Equals(driveName, options.SharePointDriveName, StringComparison.OrdinalIgnoreCase))
            {
                _driveId = drive.GetProperty("id").GetString();
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(_driveId))
        {
            throw new InvalidOperationException($"Drive '{options.SharePointDriveName}' not found on site '{options.SharePointSiteUrl}'.");
        }
    }

    public async Task<IReadOnlyList<SharePointFile>> ListFilesAsync(string folderPath, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        var files = new List<SharePointFile>();
        var normalized = string.IsNullOrWhiteSpace(folderPath) ? "/" : folderPath;

        if (normalized == "/")
        {
            await ListChildrenRecursiveAsync($"https://graph.microsoft.com/v1.0/drives/{_driveId}/root/children?$top=200", "/", files, cancellationToken);
        }
        else
        {
            var cleanPath = normalized.Trim('/');
            await ListChildrenRecursiveAsync($"https://graph.microsoft.com/v1.0/drives/{_driveId}/root:/{Uri.EscapeDataString(cleanPath)}:/children?$top=200", normalized, files, cancellationToken);
        }

        return files;
    }

    public async Task<byte[]> DownloadFileAsync(string itemId, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        using var request = await CreateRequestAsync(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/drives/{_driveId}/items/{itemId}/content", cancellationToken);
        using var client = _httpClientFactory.CreateClient();
        using var response = await client.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsByteArrayAsync(cancellationToken);
    }

    public async Task<DeltaResult> GetDeltaAsync(string? deltaLink, CancellationToken cancellationToken)
    {
        EnsureInitialized();

        var isInitial = string.IsNullOrWhiteSpace(deltaLink);
        var nextUrl = isInitial
            ? $"https://graph.microsoft.com/v1.0/drives/{_driveId}/root/delta"
            : deltaLink!;

        var changes = new List<DeltaChange>();
        var deltaToken = string.Empty;

        while (!string.IsNullOrWhiteSpace(nextUrl))
        {
            var document = await GetJsonAsync(nextUrl, cancellationToken);

            if (document.RootElement.TryGetProperty("value", out var values))
            {
                foreach (var item in values.EnumerateArray())
                {
                    var change = ParseDeltaItem(item);
                    if (change is not null && !change.IsFolder)
                    {
                        changes.Add(change);
                    }
                }
            }

            nextUrl = null;
            if (document.RootElement.TryGetProperty("@odata.nextLink", out var nextLinkNode))
            {
                nextUrl = nextLinkNode.GetString();
            }
            else if (document.RootElement.TryGetProperty("@odata.deltaLink", out var deltaLinkNode))
            {
                deltaToken = deltaLinkNode.GetString() ?? string.Empty;
            }
        }

        return new DeltaResult
        {
            Changes = changes,
            DeltaToken = deltaToken,
            IsInitialSync = isInitial
        };
    }

    public async Task<FilePermissions> GetFilePermissionsAsync(string fileId, string filePath, CancellationToken cancellationToken)
    {
        EnsureInitialized();
        var document = await GetJsonAsync($"https://graph.microsoft.com/v1.0/drives/{_driveId}/items/{fileId}/permissions", cancellationToken);

        var permissions = new List<SharePointPermission>();
        if (document.RootElement.TryGetProperty("value", out var values))
        {
            foreach (var perm in values.EnumerateArray())
            {
                var parsed = ParsePermission(perm);
                if (parsed is not null)
                {
                    permissions.Add(parsed);
                }
            }
        }

        return new FilePermissions
        {
            FileId = fileId,
            FilePath = filePath,
            Permissions = permissions
        };
    }

    private async Task ListChildrenRecursiveAsync(string initialUrl, string parentPath, List<SharePointFile> sink, CancellationToken cancellationToken)
    {
        var nextUrl = initialUrl;

        while (!string.IsNullOrWhiteSpace(nextUrl))
        {
            var document = await GetJsonAsync(nextUrl, cancellationToken);
            if (document.RootElement.TryGetProperty("value", out var values))
            {
                foreach (var item in values.EnumerateArray())
                {
                    var name = item.GetProperty("name").GetString() ?? string.Empty;
                    var currentPath = parentPath == "/" ? $"/{name}" : $"{parentPath.TrimEnd('/')}/{name}";

                    if (item.TryGetProperty("folder", out _))
                    {
                        var folderId = item.GetProperty("id").GetString();
                        if (!string.IsNullOrWhiteSpace(folderId))
                        {
                            await ListChildrenRecursiveAsync(
                                $"https://graph.microsoft.com/v1.0/drives/{_driveId}/items/{folderId}/children?$top=200",
                                currentPath,
                                sink,
                                cancellationToken);
                        }
                    }
                    else if (item.TryGetProperty("file", out _))
                    {
                        sink.Add(new SharePointFile
                        {
                            Id = item.GetProperty("id").GetString() ?? string.Empty,
                            Name = name,
                            Path = currentPath,
                            Size = item.TryGetProperty("size", out var sizeNode) ? sizeNode.GetInt64() : 0,
                            LastModified = item.TryGetProperty("lastModifiedDateTime", out var modifiedNode) &&
                                           DateTimeOffset.TryParse(modifiedNode.GetString(), out var modified)
                                ? modified
                                : null,
                            ContentHash = item.TryGetProperty("cTag", out var cTagNode)
                                ? cTagNode.GetString()
                                : item.TryGetProperty("eTag", out var eTagNode) ? eTagNode.GetString() : null
                        });
                    }
                }
            }

            nextUrl = null;
            if (document.RootElement.TryGetProperty("@odata.nextLink", out var nextLinkNode))
            {
                nextUrl = nextLinkNode.GetString();
            }
        }
    }

    private static DeltaChange? ParseDeltaItem(JsonElement item)
    {
        var itemId = item.TryGetProperty("id", out var idNode) ? idNode.GetString() ?? string.Empty : string.Empty;
        var itemName = item.TryGetProperty("name", out var nameNode) ? nameNode.GetString() ?? string.Empty : string.Empty;
        var isFolder = item.TryGetProperty("folder", out _);

        var itemPath = string.Empty;
        if (item.TryGetProperty("parentReference", out var parentReference) && parentReference.TryGetProperty("path", out var pathNode))
        {
            var parentPathRaw = pathNode.GetString() ?? string.Empty;
            var parentPath = parentPathRaw.Contains(':') ? parentPathRaw[(parentPathRaw.IndexOf(':') + 1)..] : string.Empty;
            itemPath = string.IsNullOrEmpty(parentPath) ? $"/{itemName}" : $"{parentPath.TrimEnd('/')}/{itemName}";
        }

        if (item.TryGetProperty("deleted", out _))
        {
            return new DeltaChange
            {
                ChangeType = DeltaChangeType.Deleted,
                ItemId = itemId,
                ItemName = itemName,
                ItemPath = itemPath,
                IsFolder = isFolder
            };
        }

        if (isFolder)
        {
            return new DeltaChange
            {
                ChangeType = DeltaChangeType.CreatedOrModified,
                ItemId = itemId,
                ItemName = itemName,
                ItemPath = itemPath,
                IsFolder = true
            };
        }

        if (!item.TryGetProperty("file", out _))
        {
            return null;
        }

        DateTimeOffset? modified = null;
        if (item.TryGetProperty("lastModifiedDateTime", out var modifiedNode) && DateTimeOffset.TryParse(modifiedNode.GetString(), out var parsed))
        {
            modified = parsed;
        }

        return new DeltaChange
        {
            ChangeType = DeltaChangeType.CreatedOrModified,
            ItemId = itemId,
            ItemName = itemName,
            ItemPath = itemPath,
            IsFolder = false,
            File = new SharePointFile
            {
                Id = itemId,
                Name = itemName,
                Path = itemPath,
                Size = item.TryGetProperty("size", out var sizeNode) ? sizeNode.GetInt64() : 0,
                LastModified = modified,
                ContentHash = item.TryGetProperty("cTag", out var cTagNode)
                    ? cTagNode.GetString()
                    : item.TryGetProperty("eTag", out var eTagNode) ? eTagNode.GetString() : null
            }
        };
    }

    private static SharePointPermission? ParsePermission(JsonElement permission)
    {
        var permissionId = permission.TryGetProperty("id", out var idNode) ? idNode.GetString() ?? string.Empty : string.Empty;
        var inherited = permission.TryGetProperty("inheritedFrom", out _);

        string identityType = "unknown";
        string displayName = string.Empty;
        string? email = null;
        string? identityId = null;

        if (permission.TryGetProperty("grantedToV2", out var grantedToV2))
        {
            if (grantedToV2.TryGetProperty("user", out var user))
            {
                identityType = "user";
                displayName = user.TryGetProperty("displayName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                email = user.TryGetProperty("email", out var em) ? em.GetString() : null;
                identityId = user.TryGetProperty("id", out var uid) ? uid.GetString() : null;
            }
            else if (grantedToV2.TryGetProperty("group", out var group))
            {
                identityType = "group";
                displayName = group.TryGetProperty("displayName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                email = group.TryGetProperty("email", out var em) ? em.GetString() : null;
                identityId = group.TryGetProperty("id", out var gid) ? gid.GetString() : null;
            }
            else if (grantedToV2.TryGetProperty("siteGroup", out var siteGroup))
            {
                identityType = "siteGroup";
                displayName = siteGroup.TryGetProperty("displayName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                identityId = siteGroup.TryGetProperty("id", out var sgid) ? sgid.GetRawText() : null;
            }
            else if (grantedToV2.TryGetProperty("siteUser", out var siteUser))
            {
                identityType = "user";
                displayName = siteUser.TryGetProperty("displayName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                email = siteUser.TryGetProperty("email", out var em) ? em.GetString() : null;
                identityId = siteUser.TryGetProperty("id", out var suid) ? suid.GetString() : null;
            }
        }

        var roles = Array.Empty<string>();
        if (permission.TryGetProperty("roles", out var rolesNode) && rolesNode.ValueKind == JsonValueKind.Array)
        {
            roles = rolesNode.EnumerateArray()
                .Select(r => r.GetString())
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Cast<string>()
                .ToArray();
        }

        return new SharePointPermission
        {
            Id = permissionId,
            IdentityType = identityType,
            DisplayName = displayName,
            IdentityId = identityId,
            Email = email,
            Roles = roles,
            Inherited = inherited
        };
    }

    private void EnsureInitialized()
    {
        if (string.IsNullOrWhiteSpace(_siteId) || string.IsNullOrWhiteSpace(_driveId))
        {
            throw new InvalidOperationException("SharePointGraphClient.InitializeAsync must be called before usage.");
        }
    }

    private async Task<HttpRequestMessage> CreateRequestAsync(HttpMethod method, string url, CancellationToken cancellationToken)
    {
        var token = await _credential.GetTokenAsync(new TokenRequestContext(_scopes), cancellationToken);
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return request;
    }

    private async Task<JsonDocument> GetJsonAsync(string url, CancellationToken cancellationToken)
    {
        using var request = await CreateRequestAsync(HttpMethod.Get, url, cancellationToken);
        using var client = _httpClientFactory.CreateClient();
        using var response = await client.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        return await JsonDocument.ParseAsync(stream, default, cancellationToken);
    }
}
