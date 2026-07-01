#!/bin/bash
# Post-deployment configuration for the App Service + Connector Namespace sample.
#
# This is the App Service analogue of the Functions sample's postdeploy. Two jobs:
#   1. Create the Office 365 OnNewEmailV3 trigger config. Unlike the Functions
#      version, the callback URL is a plain App Service route
#      (https://<app>.azurewebsites.net/api/<endpoint>) with NO connector
#      webhook path and NO system key. Instead, notificationDetails.authentication
#      = ManagedServiceIdentity makes the connector attach a real Entra ID bearer
#      token (minted from the trigger UAMI) to every callback. App Service
#      built-in authentication (Easy Auth) validates that token at the edge.
#   2. Walk the operator through OAuth consent for each of the three connections
#      (Office 365 Outlook, Teams, Office 365 Users).
#
# Connection access policies for the web-app MI, the trigger MI, and the deployer
# user are created by Bicep (infra/connectorNamespace.bicep), so this script does
# not grant ACLs.

set -e

# Colors
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${YELLOW}Post-deployment configuration...${NC}"

outputs=$(azd env get-values --output json)

if ! command -v jq &> /dev/null; then
    echo -e "${RED}Error: jq is required for this script. Please install jq.${NC}"
    exit 1
fi

subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID')
resourceGroupName=$(echo "$outputs" | jq -r '.resourceGroupName')
connectorNamespaceName=$(echo "$outputs" | jq -r '.connectorNamespaceName')
connectorNamespaceConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceConnectionName')
connectorNamespaceTeamsConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceTeamsConnectionName')
connectorNamespaceOffice365usersConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceOffice365usersConnectionName')
appServiceName=$(echo "$outputs" | jq -r '.appServiceName')
appServiceDefaultHostname=$(echo "$outputs" | jq -r '.appServiceDefaultHostname')
office365EndpointName=$(echo "$outputs" | jq -r '.office365EndpointName')
entraAppClientId=$(echo "$outputs" | jq -r '.entraAppClientId')
triggerIdentityResourceId=$(echo "$outputs" | jq -r '.triggerIdentityResourceId')

# --- Create Connector Namespace trigger config ------------------------------
echo ""
echo -e "${YELLOW}Creating Connector Namespace trigger config...${NC}"

triggerName="${connectorNamespaceConnectionName}-trigger"

# Plain App Service route. No /runtime/webhooks/connector, no code= system key.
# The single enforcement point is App Service built-in authentication validating
# the Entra ID token attached below.
callbackUrl="https://${appServiceDefaultHostname}/api/${office365EndpointName}"

apiUrl="https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/connectorGateways/${connectorNamespaceName}/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

# notificationDetails.authentication tells the connector to mint an Entra ID
# token from the user-assigned MI referenced by `identity` and attach it to the
# callback. The token audience must match an allowedAudience configured on the
# web app's Easy Auth -- we use the Entra app's clientId.
body=$(cat <<JSON
{
  "properties": {
    "description": "Office 365 Outlook trigger config (secured with MI + App Service built-in authentication)",
    "connectionDetails": {
      "connectorName": "office365",
      "connectionName": "${connectorNamespaceConnectionName}"
    },
    "operationName": "OnNewEmailV3",
    "parameters": [
      {
        "name": "folderPath",
        "value": "Inbox"
      }
    ],
    "notificationDetails": {
      "callbackUrl": "${callbackUrl}",
      "httpMethod": "Post",
      "authentication": {
        "type": "ManagedServiceIdentity",
        "audience": "${entraAppClientId}",
        "identity": "${triggerIdentityResourceId}"
      }
    }
  }
}
JSON
)

echo -e "${CYAN}  API URL: ${apiUrl}${NC}"
echo -e "${CYAN}  Callback URL: ${callbackUrl}${NC}"
echo -e "${CYAN}  Token audience: ${entraAppClientId}${NC}"

az rest --method PUT --url "${apiUrl}" --body "${body}"

echo -e "${GREEN}✅ Connector Namespace trigger config created.${NC}"

# --- Install the official connector-namespace az CLI extension --------------
# Resolve the latest released wheel URL from Azure/Connectors GitHub releases.
# All releases are pre-release so /releases/latest 404s; fetch /releases?per_page=1.
# Pin a version by exporting CONNECTOR_NAMESPACE_EXT_URL before running azd up.
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    CONNECTOR_NAMESPACE_EXT_URL=$(curl -fsSL \
        "https://api.github.com/repos/Azure/Connectors/releases?per_page=1" \
        | grep -oE '"browser_download_url"\s*:\s*"[^"]*connector_namespace[^"]*\.whl"' \
        | head -1 \
        | sed 's/.*"\(https[^"]*\)".*/\1/')
fi
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    echo -e "${RED}ERROR: could not resolve connector-namespace extension URL from Azure/Connectors releases${NC}" >&2
    exit 2
fi
if [[ -z "$(az extension show --name connector-namespace --query name -o tsv 2>/dev/null || true)" ]]; then
    echo -e "${CYAN}Installing 'connector-namespace' Azure CLI extension from $CONNECTOR_NAMESPACE_EXT_URL${NC}"
    az extension add --upgrade --yes --source "$CONNECTOR_NAMESPACE_EXT_URL"
fi

# --- Authorize the connector connections (OAuth consent) --------------------
# Portal authorization UX is not yet available for Connector Namespace
# connections, so we drive OAuth consent through the CLI:
#   1. `connection list-consent-links` returns a one-shot consent URL.
#   2. Open it in a browser; the user signs in and consent is persisted.
#   3. Poll properties.overallStatus until it flips to `Connected`.
echo ""
echo -e "${YELLOW}Authorizing connector connections via Azure CLI...${NC}"

open_url() {
    local url="$1"
    if command -v xdg-open >/dev/null 2>&1; then
        xdg-open "$url" >/dev/null 2>&1 || true
    elif command -v open >/dev/null 2>&1; then
        open "$url" >/dev/null 2>&1 || true
    elif command -v wslview >/dev/null 2>&1; then
        wslview "$url" >/dev/null 2>&1 || true
    fi
}

authorize_connection() {
    local connectionName="$1"
    local description="$2"

    echo -e "${CYAN}-> Authorizing ${description} connection: ${connectionName}${NC}"

    local currentStatus
    currentStatus=$(az connector-namespace connection show \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        -n "${connectionName}" --query "properties.overallStatus" -o tsv 2>/dev/null || echo "")
    if [[ "$(echo "$currentStatus" | tr '[:upper:]' '[:lower:]')" == "connected" ]]; then
        echo -e "${GREEN}   already Connected; skipping consent flow${NC}"
        return
    fi

    local consentJson link paramsFile
    paramsFile=$(mktemp)
    echo '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' > "${paramsFile}"
    consentJson=$(az connector-namespace connection list-consent-links \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        --connection-name "${connectionName}" --parameters "@${paramsFile}" -o json 2>/dev/null || echo "")
    rm -f "${paramsFile}"
    link=$(echo "${consentJson}" | jq -r '.value[0].link // empty' 2>/dev/null || echo "")
    if [[ -z "${link}" ]]; then
        echo -e "${RED}   list-consent-links returned no link; skipping${NC}"
        return
    fi

    echo -e "${CYAN}   opening browser for OAuth consent...${NC}"
    echo -e "${CYAN}   (if no tab opens, paste this URL manually:${NC}"
    echo -e "${CYAN}      ${link})${NC}"
    open_url "${link}"

    local deadline=$(($(date +%s) + 300))
    local lastStatus=""
    local s=""
    while [[ $(date +%s) -lt $deadline ]]; do
        s=$(az connector-namespace connection show \
            -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
            -n "${connectionName}" --query "properties.overallStatus" -o tsv 2>/dev/null || echo "")
        if [[ "$s" != "$lastStatus" ]]; then
            echo -e "${CYAN}   status: ${s:-?}${NC}"
            lastStatus="$s"
        fi
        if [[ "$(echo "$s" | tr '[:upper:]' '[:lower:]')" == "connected" ]]; then
            echo -e "${GREEN}   ✓ ${connectionName} authenticated${NC}"
            return
        fi
        sleep 3
    done
    echo -e "${YELLOW}   timed out waiting for consent (5 min). Re-run this script when ready.${NC}"
}

authorize_connection "${connectorNamespaceConnectionName}"               "Office 365 Outlook (trigger + sender history + flag)"
authorize_connection "${connectorNamespaceTeamsConnectionName}"          "Teams (post triage card)"
authorize_connection "${connectorNamespaceOffice365usersConnectionName}" "Office 365 Users (IN-ORG badge + manager enrichment)"

echo ""
echo -e "${GREEN}✅ All connector connections authorized.${NC}"
echo -e "${GREEN}   Tail logs:  az webapp log tail -g ${resourceGroupName} -n ${appServiceName}${NC}"
echo ""
