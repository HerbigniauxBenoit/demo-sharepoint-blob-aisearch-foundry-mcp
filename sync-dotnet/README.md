# SharePoint to Blob Sync - .NET 9 Azure Function

Migration C# de la partie `sync` Python, en Azure Function moderne avec **isolated worker model**.

## 📚 Documentation

- 🚀 **[Quick Start](QUICKSTART.md)** - Démarrage rapide en 30 minutes
- 📖 **[Guide de déploiement](DEPLOYMENT.md)** - Comment déployer sur Azure (manuel)
- 🔐 **[Configuration Managed Identity](MANAGED_IDENTITY_SETUP.md)** - Détails sur les permissions

## 🚀 Quick Start

### Option 1 : Déploiement rapide (30 min)

```powershell
# 1. Configurer les permissions
.\setup-managed-identity.ps1 `
	-FunctionAppName "votre-function-app" `
	-ResourceGroupName "votre-resource-group" `
	-StorageAccountName "votre-storage-account"

# 2. Déployer le code
.\deploy.ps1 `
	-FunctionAppName "votre-function-app" `
	-ResourceGroupName "votre-resource-group"
```

### Option 2 : Utiliser les commandes PowerShell

```powershell
# Charger les commandes
. .\commands.ps1

# Voir l'aide
Show-Help

# Déploiement complet en une commande
Deploy-Complete -FunctionAppName "ma-fonction" -ResourceGroup "mon-rg" -StorageAccount "mon-storage"
```

👉 **Consultez [QUICKSTART.md](QUICKSTART.md) pour le guide complet**

### Développement local

```powershell
# Restaurer les packages
dotnet restore

# Lancer la function
func start

# Ou utiliser les commandes
. .\commands.ps1
Start-LocalFunction
```

## 🚀 Architecture

- **.NET 9** avec les dernières fonctionnalités
- **Azure Functions v4** (isolated worker)
- **Application Insights** pour la télémétrie
- **Health checks** intégrés
- **Structured logging** avec Microsoft.Extensions.Logging
- **Dependency Injection** moderne

## ✨ Fonctionnalités

- ✅ Sync delta via Microsoft Graph (`/root/delta`)
- ✅ Persistance du delta token dans Blob (`.sync-state/delta-token.json`)
- ✅ Fallback full sync (`FORCE_FULL_SYNC=true`)
- ✅ Suppression des blobs orphelins (`DELETE_ORPHANED_BLOBS=true`)
- ✅ Sync des permissions SharePoint vers metadata blob (`user_ids`, `group_ids`)
- ✅ Mode dry-run (`DRY_RUN=true`)
- ✅ Health check endpoint (`/api/health`)

## 📋 Pré-requis

- **.NET SDK 9.0+**
- **Azure Functions Core Tools v4**
- **Azure CLI** (pour l'authentification)

## 🏃 Lancer en local

```powershell
cd sync-dotnet

# Restaurer les packages
dotnet restore

# Lancer la function
func start
```

## 🔌 Endpoints

### Sync Endpoint
- **URL**: `http://localhost:7071/api/sharepoint-sync`
- **Methods**: `GET`, `POST`
- **Auth**: Function key required

### Health Check
- **URL**: `http://localhost:7071/api/health`
- **Method**: `GET`
- **Auth**: Anonymous

## 📖 Exemples d'utilisation

### Dry run avec GET
```powershell
curl "http://localhost:7071/api/sharepoint-sync?dry_run=true"
```

### Force full sync avec POST
```powershell
curl -X POST "http://localhost:7071/api/sharepoint-sync" `
  -H "Content-Type: application/json" `
  -d '{"force_full_sync": true, "dry_run": false}'
```

### Query parameters disponibles
```powershell
curl "http://localhost:7071/api/sharepoint-sync?force_full_sync=true&dry_run=true&sync_permissions=true"
```

### Health check
```powershell
curl "http://localhost:7071/api/health"
```

## ⚙️ Variables d'environnement

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `SHAREPOINT_SITE_URL` | ✅ Yes | — | URL du site SharePoint |
| `SHAREPOINT_DRIVE_NAME` | No | `Documents` | Nom de la bibliothèque |
| `SHAREPOINT_FOLDER_PATH` | No | `/` | Chemin du dossier à syncer |
| `AZURE_STORAGE_ACCOUNT_NAME` | ✅ Yes | — | Nom du compte de stockage |
| `AZURE_BLOB_CONTAINER_NAME` | No | `sharepoint-sync` | Nom du container blob |
| `AZURE_BLOB_PREFIX` | No | — | Préfixe pour les blobs |
| `DELETE_ORPHANED_BLOBS` | No | `false` | Supprimer les blobs orphelins |
| `DRY_RUN` | No | `false` | Mode simulation |
| `SYNC_PERMISSIONS` | No | `false` | Synchroniser les permissions |
| `FORCE_FULL_SYNC` | No | `false` | Forcer un sync complet |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | No | — | Application Insights |

## 🔐 Authentification

L'application utilise **DefaultAzureCredential** qui supporte :
- Azure CLI (`az login`)
- Managed Identity (en Azure)
- Environment variables
- Visual Studio / VS Code

## 🏗️ Structure du projet

```
sync-dotnet/
├── Functions/
│   ├── HealthCheckFunction.cs    # Health check endpoint
│   └── SyncFunction.cs            # Main sync endpoint
├── Services/
│   ├── SharePointGraphClient.cs   # Microsoft Graph client
│   ├── BlobStorageSyncClient.cs   # Azure Blob Storage client
│   ├── SharePointSyncOrchestrator.cs  # Orchestration logic
│   └── SyncOptions.cs             # Configuration options
├── Models/
│   └── SyncModels.cs              # Data models
├── Program.cs                      # Application entry point
├── host.json                       # Function runtime config
└── local.settings.json            # Local development settings
```

## 📦 Déploiement Azure

```powershell
# Build et publish
dotnet publish -c Release -o ./publish

# Deployer avec Azure Functions Core Tools
func azure functionapp publish <your-function-app-name>

# Ou avec Azure CLI
az functionapp deployment source config-zip `
  -g <resource-group> `
  -n <function-app-name> `
  --src ./publish.zip
```

## 🔍 Monitoring

- **Application Insights** : Télémétrie automatique
- **Structured logs** : Tous les logs incluent des propriétés structurées
- **Health checks** : Endpoint `/api/health` pour monitoring
- **Dependency tracking** : Graph API et Blob Storage

## 🛠️ Développement

### Bonnes pratiques appliquées
- ✅ Isolated worker model (recommandé)
- ✅ Dependency injection
- ✅ Async/await partout
- ✅ CancellationToken support
- ✅ Structured logging
- ✅ Health checks
- ✅ Retry policies configurées
- ✅ Nullable reference types enabled

### Tests
```powershell
dotnet test
```
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
