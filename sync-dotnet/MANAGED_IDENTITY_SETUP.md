# Configuration Managed Identity pour Azure Function

## Prérequis

- Azure CLI installé et connecté (`az login`)
- Permissions suffisantes sur l'Azure Subscription
- Permissions Azure AD pour accorder les App Roles

## Configuration automatique

### Utilisation du script PowerShell

```powershell
.\setup-managed-identity.ps1 `
    -FunctionAppName "votre-function-app" `
    -ResourceGroupName "votre-resource-group" `
    -StorageAccountName "votre-storage-account" `
    -SubscriptionId "votre-subscription-id"
```

## Configuration manuelle

### 1. Activer la System-Assigned Managed Identity

```bash
az functionapp identity assign \
    --name <function-app-name> \
    --resource-group <resource-group>
```

### 2. Récupérer le Principal ID

```bash
PRINCIPAL_ID=$(az functionapp identity show \
    --name <function-app-name> \
    --resource-group <resource-group> \
    --query principalId \
    --output tsv)
```

### 3. Accorder les permissions Blob Storage

```bash
az role assignment create \
    --assignee $PRINCIPAL_ID \
    --role "Storage Blob Data Contributor" \
    --scope "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Storage/storageAccounts/<storage-account-name>"
```

### 4. Accorder les permissions Microsoft Graph API

#### Via Azure Portal
1. Allez sur **Azure Active Directory** > **Enterprise Applications**
2. Recherchez le nom de votre Function App
3. Allez dans **Permissions** > **Add Permission**
4. Sélectionnez **Microsoft Graph** > **Application Permissions**
5. Ajoutez :
   - `Sites.Read.All`
   - `Files.Read.All`
6. Cliquez sur **Grant admin consent**

#### Via Azure CLI

```bash
# Récupérer le Service Principal ID de Microsoft Graph
GRAPH_SP_ID=$(az ad sp show --id 00000003-0000-0000-c000-000000000000 --query id --output tsv)

# Accorder Sites.Read.All
az rest --method POST \
    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PRINCIPAL_ID/appRoleAssignments" \
    --headers "Content-Type=application/json" \
    --body "{
      \"principalId\": \"$PRINCIPAL_ID\",
      \"resourceId\": \"$GRAPH_SP_ID\",
      \"appRoleId\": \"332a536c-c7ef-4017-ab91-336970924f0d\"
    }"

# Accorder Files.Read.All
az rest --method POST \
    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PRINCIPAL_ID/appRoleAssignments" \
    --headers "Content-Type=application/json" \
    --body "{
      \"principalId\": \"$PRINCIPAL_ID\",
      \"resourceId\": \"$GRAPH_SP_ID\",
      \"appRoleId\": \"df85f4d6-205c-4ac5-a5ea-6bf408dba283\"
    }"
```

## Variables d'environnement Azure Function

Configurez ces App Settings dans votre Azure Function :

```json
{
  "SHAREPOINT_SITE_URL": "https://votre-tenant.sharepoint.com/sites/votre-site",
  "SHAREPOINT_DRIVE_NAME": "Documents",
  "SHAREPOINT_FOLDER_PATH": "/",
  "AZURE_STORAGE_ACCOUNT_NAME": "votre-storage-account",
  "AZURE_BLOB_CONTAINER_NAME": "sharepoint-sync",
  "AZURE_BLOB_PREFIX": "",
  "DELETE_ORPHANED_BLOBS": "false",
  "DRY_RUN": "false",
  "SYNC_PERMISSIONS": "true",
  "FORCE_FULL_SYNC": "false"
}
```

### Via Azure CLI

```bash
az functionapp config appsettings set \
    --name <function-app-name> \
    --resource-group <resource-group> \
    --settings \
        SHAREPOINT_SITE_URL="https://votre-tenant.sharepoint.com/sites/votre-site" \
        SHAREPOINT_DRIVE_NAME="Documents" \
        SHAREPOINT_FOLDER_PATH="/" \
        AZURE_STORAGE_ACCOUNT_NAME="votre-storage-account" \
        AZURE_BLOB_CONTAINER_NAME="sharepoint-sync" \
        AZURE_BLOB_PREFIX="" \
        DELETE_ORPHANED_BLOBS="false" \
        DRY_RUN="false" \
        SYNC_PERMISSIONS="true" \
        FORCE_FULL_SYNC="false"
```

## Vérification de la configuration

### 1. Vérifier la Managed Identity

```bash
az functionapp identity show \
    --name <function-app-name> \
    --resource-group <resource-group>
```

### 2. Vérifier les rôles Storage

```bash
az role assignment list \
    --assignee $PRINCIPAL_ID \
    --scope "/subscriptions/<subscription-id>/resourceGroups/<resource-group>/providers/Microsoft.Storage/storageAccounts/<storage-account-name>"
```

### 3. Vérifier les permissions Graph API

```bash
az rest --method GET \
    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$PRINCIPAL_ID/appRoleAssignments"
```

## Dépannage

### Erreur 500 en production

1. **Vérifier les logs Application Insights**
   ```bash
   az monitor app-insights query \
       --app <app-insights-name> \
       --analytics-query "traces | where timestamp > ago(1h) | order by timestamp desc"
   ```

2. **Vérifier que la Managed Identity est activée**
   ```bash
   az functionapp identity show --name <function-app-name> --resource-group <resource-group>
   ```

3. **Vérifier les permissions**
   - Storage : Rôle "Storage Blob Data Contributor"
   - Graph API : App Roles "Sites.Read.All" et "Files.Read.All"

4. **Attendre la propagation des permissions**
   - Les permissions Graph API peuvent prendre 5-10 minutes pour se propager

### Erreur d'authentification Graph API

Si vous obtenez "Insufficient privileges to complete the operation" :
- Vérifiez que les App Roles sont bien accordées (pas les Delegated Permissions)
- Vérifiez que le consentement administrateur a été accordé
- Attendez quelques minutes pour la propagation

### Erreur d'accès Blob Storage

Si vous obtenez "AuthorizationPermissionMismatch" :
- Vérifiez que le rôle "Storage Blob Data Contributor" est bien assigné
- Vérifiez que le nom du Storage Account dans les variables d'environnement est correct
- Le Managed Identity doit avoir accès au niveau du Storage Account ou du Container

## Architecture

```
┌─────────────────────┐
│  Azure Function     │
│  (Managed Identity) │
└──────────┬──────────┘
           │
           ├─────────────────────────┐
           │                         │
           ▼                         ▼
    ┌─────────────┐          ┌──────────────┐
    │   Storage   │          │ Microsoft    │
    │   Account   │          │ Graph API    │
    │             │          │              │
    │ - Sites     │          │ - SharePoint │
    │ - Container │          │ - OneDrive   │
    └─────────────┘          └──────────────┘
```

## Permissions requises

### Microsoft Graph API (Application Permissions)
- **Sites.Read.All** : Lire les sites et les listes SharePoint
- **Files.Read.All** : Lire tous les fichiers

### Azure Storage
- **Storage Blob Data Contributor** : Lire, écrire et supprimer des blobs et des conteneurs

## Sécurité

- Utilisez **System-Assigned Managed Identity** pour simplifier la gestion
- Les identités managées n'ont pas de secrets/mots de passe à gérer
- Appliquez le principe du moindre privilège
- Utilisez des rôles RBAC au lieu de clés d'accès Storage
- Activez les logs et la surveillance via Application Insights

## Références

- [Managed Identity pour Azure Functions](https://learn.microsoft.com/en-us/azure/app-service/overview-managed-identity)
- [Azure Storage avec Managed Identity](https://learn.microsoft.com/en-us/azure/storage/common/storage-auth-aad)
- [Microsoft Graph avec Managed Identity](https://learn.microsoft.com/en-us/graph/auth-v2-service)
- [DefaultAzureCredential](https://learn.microsoft.com/en-us/dotnet/api/azure.identity.defaultazurecredential)
