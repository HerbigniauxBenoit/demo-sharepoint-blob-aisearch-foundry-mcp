param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$FunctionAppName,

    [string]$FunctionName = "sharepoint_sync_http",
    [string]$Route = "sharepoint-sync",
    [string]$Method = "GET",
    [int]$RequestTimeoutSec = 45,
    [switch]$InvokeEndpoint
)

$ErrorActionPreference = "Stop"

function Fail([string]$message) {
    Write-Host "❌ $message" -ForegroundColor Red
    exit 1
}

function Step([string]$label) {
    Write-Host "\n==> $label" -ForegroundColor Cyan
}

Step "[1/5] Vérification Azure CLI et session"
try {
    az account show | Out-Null
}
catch {
    Fail "Azure CLI non connecté. Lance d'abord: az login"
}

Step "[2/5] Vérification runtime Python"
$linuxFxVersion = az functionapp config show `
    --resource-group $ResourceGroup `
    --name $FunctionAppName `
    --query linuxFxVersion -o tsv

if (-not $linuxFxVersion) {
    Fail "Impossible de lire linuxFxVersion pour $FunctionAppName"
}

if (-not $linuxFxVersion.StartsWith("PYTHON|")) {
    Fail "Runtime invalide: $linuxFxVersion (attendu: PYTHON|3.x)"
}

$workerRuntime = az functionapp config appsettings list `
    --resource-group $ResourceGroup `
    --name $FunctionAppName `
    --query "[?name=='FUNCTIONS_WORKER_RUNTIME'].value | [0]" -o tsv

if ($workerRuntime -ne "python") {
    Fail "FUNCTIONS_WORKER_RUNTIME=$workerRuntime (attendu: python)"
}

Write-Host "✅ Runtime OK: $linuxFxVersion / FUNCTIONS_WORKER_RUNTIME=python" -ForegroundColor Green

Step "[3/5] Vérification de la présence de la fonction"
$functionNames = az functionapp function list `
    --resource-group $ResourceGroup `
    --name $FunctionAppName `
    --query "[].name" -o tsv

if (-not $functionNames) {
    Fail "Aucune fonction détectée dans l'app (liste vide)."
}

$hasFunction = $false
$functionNamesArray = @($functionNames -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
foreach ($fn in $functionNamesArray) {
    if ($fn -eq "$FunctionAppName/$FunctionName" -or $fn -eq $FunctionName -or $fn.EndsWith("/$FunctionName")) {
        $hasFunction = $true
        break
    }
}

if (-not $hasFunction) {
    Write-Host "Fonctions trouvées:" -ForegroundColor Yellow
    $functionNamesArray | ForEach-Object { Write-Host "  - $_" }
    Fail "La fonction '$FunctionName' n'a pas été trouvée."
}

Write-Host "✅ Fonction trouvée: $FunctionName" -ForegroundColor Green

Step "[4/5] Vérification du trigger HTTP"
$httpTriggerCount = az functionapp function show `
    --resource-group $ResourceGroup `
    --name $FunctionAppName `
    --function-name $FunctionName `
    --query "config.bindings[?type=='httpTrigger'] | length(@)" -o tsv

if (-not $httpTriggerCount -or [int]$httpTriggerCount -lt 1) {
    Fail "Aucun binding httpTrigger trouvé pour '$FunctionName'."
}

Write-Host "✅ Binding HTTP détecté" -ForegroundColor Green

Step "[5/5] Résumé"
$baseUrl = "https://$FunctionAppName.azurewebsites.net/api/$Route"
Write-Host "App:      $FunctionAppName"
Write-Host "Function: $FunctionName"
Write-Host "Route:    /api/$Route"
Write-Host "URL:      $baseUrl"

if (-not $InvokeEndpoint) {
    Write-Host "\n✔️ Vérification terminée (sans appel HTTP)." -ForegroundColor Green
    Write-Host "Pour tester l'endpoint: relance avec -InvokeEndpoint" -ForegroundColor Yellow
    exit 0
}

Step "Test HTTP réel de l'endpoint"

$key = az functionapp keys list `
    --resource-group $ResourceGroup `
    --name $FunctionAppName `
    --query "functionKeys.default" -o tsv

if (-not $key) {
    $key = az functionapp keys list `
        --resource-group $ResourceGroup `
        --name $FunctionAppName `
        --query "masterKey" -o tsv
}

if (-not $key) {
    Fail "Impossible de récupérer une clé de fonction (functionKeys.default/masterKey)."
}

$uri = "$baseUrl?dry_run=true&code=$key"

try {
    $response = Invoke-RestMethod -Method $Method -Uri $uri -TimeoutSec $RequestTimeoutSec
    Write-Host "✅ Appel HTTP OK" -ForegroundColor Green
    if ($null -ne $response) {
        $response | ConvertTo-Json -Depth 10
    }
}
catch {
    Fail "Appel HTTP en échec: $($_.Exception.Message)"
}
