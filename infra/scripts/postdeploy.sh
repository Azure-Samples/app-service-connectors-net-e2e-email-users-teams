#!/bin/bash
# Post-deployment configuration for the App Service + Connector Namespace sample.
#
# This script walks the operator through OAuth consent for the Office 365
# Outlook, Teams, and Office 365 Users connections. The first-class App Service
# trigger is created afterward in the Connector Namespace portal, where the
# destination wizard binds the trigger to the deployed web app and route.
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
entraAppIdentifierUri=$(echo "$outputs" | jq -r '.entraAppIdentifierUri')

# --- Install the official connector-namespace az CLI extension --------------
# Resolve the latest released wheel URL from Azure/Connectors GitHub releases.
# All releases are pre-release so /releases/latest 404s; fetch /releases?per_page=1.
# Pin a version by exporting CONNECTOR_NAMESPACE_EXT_URL before running azd up.
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    CONNECTOR_NAMESPACE_EXT_URL=$(curl -fsSL \
        "https://api.github.com/repos/Azure/Connectors/releases?per_page=20" \
        | grep -oE '"browser_download_url"\s*:\s*"[^"]*connector_namespace[^"]*\.whl"' \
        | head -1 \
        | sed 's/.*"\(https[^"]*\)".*/\1/')
fi
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    echo -e "${RED}ERROR: could not resolve connector-namespace extension URL from Azure/Connectors releases${NC}" >&2
    exit 2
fi
echo -e "${CYAN}Installing/updating 'connector-namespace' Azure CLI extension from $CONNECTOR_NAMESPACE_EXT_URL${NC}"
az extension add --upgrade --yes --source "$CONNECTOR_NAMESPACE_EXT_URL"

# --- Authorize the connector connections (OAuth consent) --------------------
# Drive OAuth consent through the CLI so authorization is part of the
# post-deployment flow:
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
    if ! currentStatus=$(az connector-namespace connection show \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        -n "${connectionName}" --query "properties.overallStatus" -o tsv); then
        echo -e "${RED}   failed to read connection status for ${connectionName}${NC}" >&2
        return 1
    fi
    if [[ "$(echo "$currentStatus" | tr '[:upper:]' '[:lower:]')" == "connected" ]]; then
        echo -e "${GREEN}   already Connected; skipping consent flow${NC}"
        return
    fi

    local consentJson link paramsFile
    paramsFile=$(mktemp)
    echo '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' > "${paramsFile}"
    if ! consentJson=$(az connector-namespace connection list-consent-links \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        --connection-name "${connectionName}" --parameters "@${paramsFile}" -o json); then
        rm -f "${paramsFile}"
        echo -e "${RED}   failed to create an OAuth consent link for ${connectionName}${NC}" >&2
        return 1
    fi
    rm -f "${paramsFile}"
    link=$(echo "${consentJson}" | jq -r '.value[0].link // empty')
    if [[ -z "${link}" ]]; then
        echo -e "${RED}   list-consent-links returned no link for ${connectionName}${NC}" >&2
        return 1
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
    echo -e "${RED}   timed out waiting for ${connectionName} consent after 5 minutes${NC}" >&2
    return 1
}

authorize_connection "${connectorNamespaceConnectionName}"               "Office 365 Outlook (trigger + sender history + flag)"
authorize_connection "${connectorNamespaceTeamsConnectionName}"          "Teams (post triage card)"
authorize_connection "${connectorNamespaceOffice365usersConnectionName}" "Office 365 Users (IN-ORG badge + manager enrichment)"

echo ""
echo -e "${GREEN}✅ All connector connections authorized.${NC}"
echo -e "${YELLOW}Create the App Service trigger in the Connector Namespace portal:${NC}"
echo -e "${CYAN}   https://connectors.azure.com/${subscriptionId}/${resourceGroupName}/${connectorNamespaceName}/triggers${NC}"
echo -e "${CYAN}   Source: Office 365 Outlook / When a new email arrives (V3)${NC}"
echo -e "${CYAN}   Connection: ${connectorNamespaceConnectionName}${NC}"
echo -e "${CYAN}   Destination: App Service / ${appServiceName}${NC}"
echo -e "${CYAN}   Audience: ${entraAppIdentifierUri}${NC}"
echo -e "${CYAN}   The App Service destination defaults to POST /api/webhook.${NC}"
echo -e "${GREEN}   Tail logs:  az webapp log tail -g ${resourceGroupName} -n ${appServiceName}${NC}"
echo ""
