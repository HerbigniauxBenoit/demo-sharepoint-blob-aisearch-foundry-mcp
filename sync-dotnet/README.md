# Sync (.NET 10 Azure Function)

Migration C# de la partie `sync` Python, en Azure Function (worker isolé).

## Fonctionnalités portées

- Sync delta via Microsoft Graph (`/root/delta`)
- Persistance du delta token dans Blob (`.sync-state/delta-token.json`)
- Fallback full sync (`FORCE_FULL_SYNC=true`)
- Suppression des blobs orphelins (`DELETE_ORPHANED_BLOBS=true`)
- Sync des permissions SharePoint vers metadata blob (`user_ids`, `group_ids`)
- Mode dry-run (`DRY_RUN=true`)

## Pré-requis

- .NET SDK 10
- Azure Functions Core Tools v4
- Variables d'environnement (mêmes noms que la version Python)

## Lancer en local

```powershell
cd sync-dotnet
dotnet restore
func start
```

Le trigger est HTTP (`GET`/`POST`) sur la route `/api/sharepoint-sync`.

Exemples:

```powershell
# GET
curl "http://localhost:7071/api/sharepoint-sync?dry_run=true"

# POST JSON
curl -X POST "http://localhost:7071/api/sharepoint-sync" -H "Content-Type: application/json" -d "{\"force_full_sync\":true,\"dry_run\":false}"
```

## Variables d'environnement

| Variable | Required | Default |
|----------|----------|---------|
| `SHAREPOINT_SITE_URL` | Yes | — |
| `SHAREPOINT_DRIVE_NAME` | No | `Documents` |
| `SHAREPOINT_FOLDER_PATH` | No | `/` |
| `AZURE_STORAGE_ACCOUNT_NAME` | Yes | — |
| `AZURE_BLOB_CONTAINER_NAME` | No | `sharepoint-sync` |
| `AZURE_BLOB_PREFIX` | No | — |
| `DELETE_ORPHANED_BLOBS` | No | `false` |
| `DRY_RUN` | No | `false` |
| `SYNC_PERMISSIONS` | No | `false` |
| `FORCE_FULL_SYNC` | No | `false` |

## Déploiement Azure Function

Tu peux publier ce projet vers une Function App Linux `dotnet-isolated` (Functions v4).

Exemple rapide:

```powershell
cd sync-dotnet
func azure functionapp publish <FUNCTION_APP_NAME>
```

Configurer les app settings:

```powershell
az functionapp config appsettings set --resource-group <RG> --name <FUNCTION_APP_NAME> --settings \
	FUNCTIONS_WORKER_RUNTIME=dotnet-isolated \
	SHAREPOINT_SITE_URL="https://contoso.sharepoint.com/sites/MySite" \
	SHAREPOINT_DRIVE_NAME="Documents" \
	SHAREPOINT_FOLDER_PATH="/" \
	AZURE_STORAGE_ACCOUNT_NAME="<storage-account>" \
	AZURE_BLOB_CONTAINER_NAME="sharepoint-sync" \
	AZURE_BLOB_PREFIX="" \
	DELETE_ORPHANED_BLOBS="false" \
	SYNC_PERMISSIONS="true" \
	DRY_RUN="false" \
	FORCE_FULL_SYNC="false"
```

Appeler la function en Azure:

```powershell
curl "https://<FUNCTION_APP_NAME>.azurewebsites.net/api/sharepoint-sync?dry_run=true&code=<FUNCTION_KEY>"
```

Récupérer la function key:

```powershell
az functionapp function keys list --resource-group <RG> --name <FUNCTION_APP_NAME> --function-name SharePointBlobSync
```
