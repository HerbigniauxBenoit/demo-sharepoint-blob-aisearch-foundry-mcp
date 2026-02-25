param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$FunctionAppName,

    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$StorageAccountName,

    [string]$BlobContainerName = "sharepoint-sync",
    [string]$SharePointDriveName = "Documents",
    [string]$SharePointFolderPath = "/",
    [string]$BlobPrefix = "",
    [string]$TimerSchedule = "0 0 2 * * *",
    [string]$DeleteOrphanedBlobs = "false",
    [string]$SyncPermissions = "true",
    [string]$DryRun = "false",
    [string]$ForceFullSync = "false",
    [string]$PythonRuntimeVersion = "3.12"
)

$ErrorActionPreference = "Stop"

Write-Host "[1/6] Vérification Azure CLI..."
az account show | Out-Null

Write-Host "[2/6] Configuration des app settings..."
az functionapp config appsettings set `
  --resource-group $ResourceGroup `
  --name $FunctionAppName `
  --settings `
    FUNCTIONS_WORKER_RUNTIME=python `
    TIMER_SCHEDULE=$TimerSchedule `
    SHAREPOINT_SITE_URL=$SharePointSiteUrl `
    SHAREPOINT_DRIVE_NAME=$SharePointDriveName `
    SHAREPOINT_FOLDER_PATH=$SharePointFolderPath `
    AZURE_STORAGE_ACCOUNT_NAME=$StorageAccountName `
    AZURE_BLOB_CONTAINER_NAME=$BlobContainerName `
    AZURE_BLOB_PREFIX=$BlobPrefix `
    DELETE_ORPHANED_BLOBS=$DeleteOrphanedBlobs `
    SYNC_PERMISSIONS=$SyncPermissions `
    DRY_RUN=$DryRun `
    FORCE_FULL_SYNC=$ForceFullSync | Out-Null

Write-Host "[3/6] Vérification runtime Linux/Python..."
$linuxFxVersion = az functionapp config show `
  --resource-group $ResourceGroup `
  --name $FunctionAppName `
  --query linuxFxVersion -o tsv

if (-not $linuxFxVersion.StartsWith("PYTHON|")) {
    throw "La Function App '$FunctionAppName' n'est pas en runtime Python Linux (linuxFxVersion=$linuxFxVersion)."
}

$targetLinuxFxVersion = "PYTHON|$PythonRuntimeVersion"
if ($linuxFxVersion -ne $targetLinuxFxVersion) {
    Write-Host "[4/6] Mise à jour runtime vers $targetLinuxFxVersion..."
    az functionapp config set `
      --resource-group $ResourceGroup `
      --name $FunctionAppName `
      --linux-fx-version $targetLinuxFxVersion | Out-Null
}
else {
    Write-Host "[4/6] Runtime déjà conforme ($linuxFxVersion)."
}

Write-Host "[5/6] Déploiement sans Docker (remote build)..."
Push-Location (Join-Path $PSScriptRoot "..")
try {
    func azure functionapp publish $FunctionAppName --python --build remote
}
finally {
    Pop-Location
}

Write-Host "[6/6] Déploiement terminé."
Write-Host "Logs: az functionapp log tail --resource-group $ResourceGroup --name $FunctionAppName"
