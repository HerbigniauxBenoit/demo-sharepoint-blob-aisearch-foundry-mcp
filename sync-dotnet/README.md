# Sync (.NET 8 Azure Functions)

Sync SharePoint Online files to Azure Blob Storage using an Azure Functions isolated worker.

## Runtime target

- .NET 8
- Azure Functions v4
- TimerTrigger only (no HttpTrigger)
- Hosting target: Azure Functions on Azure Container Apps

## Implemented behavior

- Full sync SharePoint files to Blob
- Configurable max file size limit (`MAX_FILE_SIZE_MB`, oversized files are skipped)
- Optional orphan blob deletion (`DELETE_ORPHANED_BLOBS`)
- Optional SharePoint permissions sync to blob metadata (`SYNC_PERMISSIONS`)
- Managed identity auth for Graph and Blob access

## Required security prerequisite (manual)

The security team must configure Graph access manually for each companion managed identity:

1. Assign Microsoft Graph application permission `Sites.Selected`.
2. Grant admin consent.
3. Grant site-level authorization on the target SharePoint site.

This prerequisite is intentionally not automated by this project.

## Local run

```powershell
cd sync-dotnet
dotnet restore
func start
```

The sync is started only by timer according to `SYNC_SCHEDULE`.

## Environment variables

| Variable | Required | Default |
|----------|----------|---------|
| `SHAREPOINT_SITE_URL` | Yes | - |
| `SHAREPOINT_DRIVE_NAME` | No | `Documents` |
| `SHAREPOINT_FOLDER_PATH` | No | `/` |
| `AZURE_STORAGE_ACCOUNT_NAME` | Yes | - |
| `AZURE_BLOB_CONTAINER_NAME` | No | `sharepoint-sync` |
| `AZURE_BLOB_PREFIX` | No | empty |
| `MAX_FILE_SIZE_MB` | No | `50` |
| `DELETE_ORPHANED_BLOBS` | No | `false` |
| `SYNC_PERMISSIONS` | No | `false` |
| `SYNC_SCHEDULE` | No | `0 */6 * * *` |
| `AZURE_CLIENT_ID` | No | empty |

For Functions host storage with managed identity, deployment config must provide:

- `AzureWebJobsStorage__accountName`
- `AzureWebJobsStorage__credential=managedidentity`
- `AzureWebJobsStorage__clientId`

## Technical debt

- The code contains delta token support primitives (`GetDeltaAsync`, save/load token), but orchestration currently runs full sync each cycle.
- This is kept unchanged intentionally for a low-risk migration path. A later iteration can enable true delta orchestration with tests.
