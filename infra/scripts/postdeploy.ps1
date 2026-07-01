# Post-deployment configuration for the App Service + Connector Namespace sample.
# See postdeploy.sh for the full explanation. In short:
#   1. Create the Office 365 trigger config whose callbackUrl is a plain App
#      Service route secured by ManagedServiceIdentity auth (Easy Auth validates
#      the attached Entra token).
#   2. Drive OAuth consent for the three connector connections.

$ErrorActionPreference = "Stop"

Write-Host "Post-deployment configuration..." -ForegroundColor Yellow

$outputs = azd env get-values --output json | ConvertFrom-Json

$subscriptionId = $outputs.AZURE_SUBSCRIPTION_ID
$resourceGroupName = $outputs.resourceGroupName
$connectorNamespaceName = $outputs.connectorNamespaceName
$connectorNamespaceConnectionName = $outputs.connectorNamespaceConnectionName
$connectorNamespaceTeamsConnectionName = $outputs.connectorNamespaceTeamsConnectionName
$connectorNamespaceOffice365usersConnectionName = $outputs.connectorNamespaceOffice365usersConnectionName
$appServiceName = $outputs.appServiceName
$appServiceDefaultHostname = $outputs.appServiceDefaultHostname
$office365EndpointName = $outputs.office365EndpointName
$entraAppClientId = $outputs.entraAppClientId
$triggerIdentityResourceId = $outputs.triggerIdentityResourceId

# --- Create Connector Namespace trigger config ---
Write-Host "Creating Connector Namespace trigger config..." -ForegroundColor Yellow

$triggerName = "$connectorNamespaceConnectionName-trigger"

# Plain App Service route. No /runtime/webhooks/connector, no code= system key.
$callbackUrl = "https://$appServiceDefaultHostname/api/$office365EndpointName"

$apiUrl = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/connectorGateways/$connectorNamespaceName/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

$body = @{
  properties = @{
    description = "Office 365 Outlook trigger config (secured with MI + App Service built-in authentication)"
    connectionDetails = @{
      connectorName = "office365"
      connectionName = $connectorNamespaceConnectionName
    }
    operationName = "OnNewEmailV3"
    parameters = @(
      @{ name = "folderPath"; value = "Inbox" }
    )
    notificationDetails = @{
      callbackUrl = $callbackUrl
      httpMethod = "Post"
      authentication = @{
        type = "ManagedServiceIdentity"
        audience = $entraAppClientId
        identity = $triggerIdentityResourceId
      }
    }
  }
} | ConvertTo-Json -Depth 10 -Compress

Write-Host "  API URL: $apiUrl" -ForegroundColor Cyan
Write-Host "  Callback URL: $callbackUrl" -ForegroundColor Cyan
Write-Host "  Token audience: $entraAppClientId" -ForegroundColor Cyan

$tmpFile = [System.IO.Path]::GetTempFileName()
$body | Out-File -FilePath $tmpFile -Encoding utf8
az rest --method PUT --url $apiUrl --body "@$tmpFile" --headers "Content-Type=application/json" | Out-Null
Remove-Item $tmpFile

Write-Host "Connector Namespace trigger config created." -ForegroundColor Green

# --- Install the official connector-namespace az CLI extension ---
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
  $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/Azure/Connectors/releases?per_page=1"
  $asset = $rel.assets | Where-Object { $_.browser_download_url -match "connector_namespace.*\.whl" } | Select-Object -First 1
  if ($asset) { $env:CONNECTOR_NAMESPACE_EXT_URL = $asset.browser_download_url }
}
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
  Write-Error "Could not resolve connector-namespace extension URL from Azure/Connectors releases"
  exit 2
}
$ext = az extension show --name connector-namespace 2>$null
if (-not $ext) {
  Write-Host "Installing 'connector-namespace' Azure CLI extension from $($env:CONNECTOR_NAMESPACE_EXT_URL)" -ForegroundColor Cyan
  az extension add --upgrade --yes --source $env:CONNECTOR_NAMESPACE_EXT_URL
}

# --- Authorize the connector connections (OAuth consent) ---
Write-Host ""
Write-Host "Authorizing connector connections via Azure CLI..." -ForegroundColor Yellow

function Authorize-Connection {
  param([string]$ConnectionName, [string]$Description)

  Write-Host "-> Authorizing $Description connection: $ConnectionName" -ForegroundColor Cyan

  $currentStatus = az connector-namespace connection show `
    -g $resourceGroupName --namespace $connectorNamespaceName `
    -n $ConnectionName --query "properties.overallStatus" -o tsv 2>$null
  if ($currentStatus -and $currentStatus.ToLower() -eq "connected") {
    Write-Host "   already Connected; skipping consent flow" -ForegroundColor Green
    return
  }

  $paramsFile = [System.IO.Path]::GetTempFileName()
  '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' | Out-File -FilePath $paramsFile -Encoding utf8
  $consentJson = az connector-namespace connection list-consent-links `
    -g $resourceGroupName --namespace $connectorNamespaceName `
    --connection-name $ConnectionName --parameters "@$paramsFile" -o json 2>$null
  Remove-Item $paramsFile
  $link = ($consentJson | ConvertFrom-Json).value[0].link
  if (-not $link) {
    Write-Host "   list-consent-links returned no link; skipping" -ForegroundColor Red
    return
  }

  Write-Host "   opening browser for OAuth consent..." -ForegroundColor Cyan
  Write-Host "   (if no tab opens, paste this URL manually:" -ForegroundColor Cyan
  Write-Host "      $link)" -ForegroundColor Cyan
  Start-Process $link

  $deadline = (Get-Date).AddMinutes(5)
  $lastStatus = ""
  while ((Get-Date) -lt $deadline) {
    $s = az connector-namespace connection show `
      -g $resourceGroupName --namespace $connectorNamespaceName `
      -n $ConnectionName --query "properties.overallStatus" -o tsv 2>$null
    if ($s -ne $lastStatus) {
      Write-Host "   status: $s" -ForegroundColor Cyan
      $lastStatus = $s
    }
    if ($s -and $s.ToLower() -eq "connected") {
      Write-Host "   OK $ConnectionName authenticated" -ForegroundColor Green
      return
    }
    Start-Sleep -Seconds 3
  }
  Write-Host "   timed out waiting for consent (5 min). Re-run this script when ready." -ForegroundColor Yellow
}

Authorize-Connection $connectorNamespaceConnectionName               "Office 365 Outlook (trigger + sender history + flag)"
Authorize-Connection $connectorNamespaceTeamsConnectionName          "Teams (post triage card)"
Authorize-Connection $connectorNamespaceOffice365usersConnectionName "Office 365 Users (IN-ORG badge + manager enrichment)"

Write-Host ""
Write-Host "All connector connections authorized." -ForegroundColor Green
Write-Host "Tail logs: az webapp log tail -g $resourceGroupName -n $appServiceName" -ForegroundColor Green
Write-Host ""
