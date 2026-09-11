# Post-deployment configuration for the App Service + Connector Namespace sample.
# See postdeploy.sh for the full explanation. This script authorizes the three
# connector connections. The first-class App Service trigger is then created in
# the Connector Namespace portal, which binds it to the web app and route.

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
$entraAppIdentifierUri = $outputs.entraAppIdentifierUri

# --- Install the official connector-namespace az CLI extension ---
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
  $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/Azure/Connectors/releases?per_page=20"
  $asset = $rel | ForEach-Object { $_.assets } | Where-Object { $_.browser_download_url -match "connector_namespace.*\.whl" } | Select-Object -First 1
  if ($asset) { $env:CONNECTOR_NAMESPACE_EXT_URL = $asset.browser_download_url }
}
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
  Write-Error "Could not resolve connector-namespace extension URL from Azure/Connectors releases"
  exit 2
}
Write-Host "Installing/updating 'connector-namespace' Azure CLI extension from $($env:CONNECTOR_NAMESPACE_EXT_URL)" -ForegroundColor Cyan
az extension add --upgrade --yes --source $env:CONNECTOR_NAMESPACE_EXT_URL

# --- Authorize the connector connections (OAuth consent) ---
Write-Host ""
Write-Host "Authorizing connector connections via Azure CLI..." -ForegroundColor Yellow

function Authorize-Connection {
  param([string]$ConnectionName, [string]$Description)

  Write-Host "-> Authorizing $Description connection: $ConnectionName" -ForegroundColor Cyan

  $currentStatus = az connector-namespace connection show `
    -g $resourceGroupName --namespace $connectorNamespaceName `
    -n $ConnectionName --query "properties.overallStatus" -o tsv
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to read connection status for $ConnectionName"
  }
  if ($currentStatus -and $currentStatus.ToLower() -eq "connected") {
    Write-Host "   already Connected; skipping consent flow" -ForegroundColor Green
    return
  }

  $paramsFile = [System.IO.Path]::GetTempFileName()
  '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' | Out-File -FilePath $paramsFile -Encoding utf8
  $consentJson = az connector-namespace connection list-consent-links `
    -g $resourceGroupName --namespace $connectorNamespaceName `
    --connection-name $ConnectionName --parameters "@$paramsFile" -o json
  $consentExitCode = $LASTEXITCODE
  Remove-Item $paramsFile
  if ($consentExitCode -ne 0) {
    throw "Failed to create an OAuth consent link for $ConnectionName"
  }
  $link = ($consentJson | ConvertFrom-Json).value[0].link
  if (-not $link) {
    throw "list-consent-links returned no link for $ConnectionName"
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
  throw "Timed out waiting for $ConnectionName consent after 5 minutes"
}

Authorize-Connection $connectorNamespaceConnectionName               "Office 365 Outlook (trigger + sender history + flag)"
Authorize-Connection $connectorNamespaceTeamsConnectionName          "Teams (post triage card)"
Authorize-Connection $connectorNamespaceOffice365usersConnectionName "Office 365 Users (IN-ORG badge + manager enrichment)"

Write-Host ""
Write-Host "All connector connections authorized." -ForegroundColor Green
Write-Host "Create the App Service trigger in the Connector Namespace portal:" -ForegroundColor Yellow
Write-Host "  https://connectors.azure.com/$subscriptionId/$resourceGroupName/$connectorNamespaceName/triggers" -ForegroundColor Cyan
Write-Host "  Source: Office 365 Outlook / When a new email arrives (V3)" -ForegroundColor Cyan
Write-Host "  Connection: $connectorNamespaceConnectionName" -ForegroundColor Cyan
Write-Host "  Destination: App Service / $appServiceName" -ForegroundColor Cyan
Write-Host "  Audience: $entraAppIdentifierUri" -ForegroundColor Cyan
Write-Host "  The App Service destination defaults to POST /api/webhook." -ForegroundColor Cyan
Write-Host "Tail logs: az webapp log tail -g $resourceGroupName -n $appServiceName" -ForegroundColor Green
Write-Host ""
