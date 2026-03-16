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

## Collecte d'informations (sans droits admin)

Objectif: récupérer uniquement des informations de diagnostic à transmettre au client quand vous n'avez pas les droits pour corriger la configuration.

### 1. Contexte Azure courant

```powershell
az account show --query "{tenantId:tenantId,subscriptionId:id,subscriptionName:name,user:user.name}" -o jsonc
az account tenant list -o table
```

### 2. Informations Function App (lecture seule)

```powershell
$RG = "<resource-group>"
$FUNC = "<function-app-name>"

az functionapp show -g $RG -n $FUNC --query "{name:name,defaultHostName:defaultHostName,state:state}" -o jsonc
```

### 3. Variables utiles (si la lecture de config est autorisée)

```powershell
az functionapp config appsettings list -g $RG -n $FUNC `
	--query "[?name=='AZURE_CLIENT_ID' || name=='EXPECTED_TENANT_ID' || name=='AZURE_TENANT_ID'].{name:name,value:value}" -o table
```

### 4. Appeler la fonction de test SharePoint

```powershell
$URL  = "https://<function-app-name>.azurewebsites.net/api/test-sharepoint"
$KEY  = "<function-key>"
$SITE = "https://safrangroup.sharepoint.com/sites/DSI-CGenAIPPM"

$body = @{ siteUrl = $SITE } | ConvertTo-Json

Invoke-RestMethod -Method Post `
	-Uri $URL `
	-Headers @{ "x-functions-key" = $KEY } `
	-ContentType "application/json" `
	-Body $body
```

### 5. Capturer l'erreur détaillée (si échec)

```powershell
try {
	Invoke-WebRequest -Method Post `
		-Uri $URL `
		-Headers @{ "x-functions-key" = $KEY } `
		-ContentType "application/json" `
		-Body $body
} catch {
	$_.Exception.Message
	$_.ErrorDetails.Message
}
```

### 6. Informations à transmettre au client

- `Tenant ID` du contexte Azure courant
- `Subscription ID`
- URL SharePoint testée
- Horodatage UTC de l'erreur
- Code HTTP (`401`/`403`)
- Message d'erreur Graph brut
- `request-id` et `client-request-id` Graph
- Claim token `tid` (si présent dans les logs de la fonction)
- Claim token `scp` et `roles` (si présents dans les logs de la fonction)
- Valeur `EXPECTED_TENANT_ID` ou `AZURE_TENANT_ID` (si définie)

### 7. Template de message client (copier/coller)

```text
Bonjour,

Nous avons effectué un test de connectivité SharePoint via Azure Function (Managed Identity user-assigned).

- Site testé: <site-url>
- Résultat: <HTTP 401/403>
- Horodatage UTC: <timestamp>
- Tenant du token (tid): <token-tid>
- Scopes du token (scp): <token-scopes>
- Roles du token: <token-roles>
- Tenant attendu (EXPECTED_TENANT_ID/AZURE_TENANT_ID): <expected-tenant-id>
- request-id Graph: <request-id>
- client-request-id Graph: <client-request-id>
- Message brut Graph: <raw-error>

Pouvez-vous vérifier côté client:
1. La concordance du tenant entre l'identité et le tenant SharePoint cible.
2. La permission Microsoft Graph application `Sites.Selected`.
3. Le grant de cette application/identité sur le site SharePoint cible.

Merci.
```

## Récupérer EXPECTED_TENANT_ID (AZ CLI)

Si vous voulez uniquement lire la valeur depuis la Function App:

```powershell
az functionapp config appsettings list `
	--resource-group <RG> `
	--name <FUNCTION_APP_NAME> `
	--query "[?name=='EXPECTED_TENANT_ID'].value | [0]" -o tsv
```

Fallback utilisé par le code (`AZURE_TENANT_ID`) si `EXPECTED_TENANT_ID` est vide:

```powershell
az functionapp config appsettings list `
	--resource-group <RG> `
	--name <FUNCTION_APP_NAME> `
	--query "[?name=='AZURE_TENANT_ID'].value | [0]" -o tsv
```

Afficher les 2 en une seule commande:

```powershell
az functionapp config appsettings list `
	--resource-group <RG> `
	--name <FUNCTION_APP_NAME> `
	--query "[?name=='EXPECTED_TENANT_ID' || name=='AZURE_TENANT_ID'].{name:name,value:value}" -o table
```

Si vous n'avez pas la permission de lire les app settings, récupérez au moins le tenant courant du contexte Azure:

```powershell
az account show --query tenantId -o tsv
```

# Site select sur le site 
# 1. Récupérer l'Object ID de Microsoft Graph
az ad sp show \
  --id "00000003-0000-0000-c000-000000000000" \
  --query "id" \
  --output tsv

# 2. Récupérer l'ID du role Sites.Selected
az ad sp show \
  --id "00000003-0000-0000-c000-000000000000" \
  --query "appRoles[?value=='Sites.Selected'].id" \
  --output tsv

# 3. Assigner
az rest \
  --method POST \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/<OBJECT_ID_MSI>/appRoleAssignments" \
  --headers "Content-Type=application/json" \
  --body '{
    "principalId": "<OBJECT_ID_MSI>",
    "resourceId": "<OBJECT_ID_GRAPH_ETAPE_1>",
    "appRoleId": "<ID_ROLE_ETAPE_2>"
  }'
