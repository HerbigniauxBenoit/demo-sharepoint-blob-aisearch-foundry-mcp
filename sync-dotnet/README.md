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

Le trigger est timer-based, avec cron via `SYNC_SCHEDULE` (défaut: `0 0 * * * *`).

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
| `SYNC_SCHEDULE` | No | `0 0 * * * *` |

## Déploiement Azure Function

Tu peux publier ce projet vers une Function App Linux `dotnet-isolated` (Functions v4).

Exemple rapide:

```powershell
cd sync-dotnet
dotnet publish -c Release
```

Puis déployer avec `az functionapp deployment source config-zip` ou CI/CD.
