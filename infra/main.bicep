targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the environment which is used to generate a short unique hash used in all resources.')
param environmentName string

@metadata({
  azd: {
    type: 'location'
  }
})
@description('Location for all resources except the Connector Namespace, which is pinned to westcentralus while in preview.')
param location string

metadata name = 'Managed Connectors on App Service — M365 Email Triage'
metadata description = 'Ports the Azure Functions connectors e2e sample (email -> Office 365 Users -> Teams -> flag) to an App Service Web App. The Connector Namespace trigger calls a plain HTTP endpoint secured by App Service built-in authentication (Easy Auth) validating the trigger UAMI token — no Functions runtime, no system keys.'

@description('Id of the user identity used for local debugging. Granted access to the connections so the same code can be run locally with `az login`.')
@metadata({
  azd: {
    type: 'principalId'
  }
})
param userPrincipalId string = deployer().objectId

@description('Route segment of the App Service endpoint that receives the connector trigger callback. The callback URL is https://<app>.azurewebsites.net/api/<value>.')
param office365EndpointName string = 'onNewEmail'

@description('The Teams Team ID (groupId) to post triage cards to.')
param teamsTeamId string

@description('The Teams Channel ID to post triage cards to.')
param teamsChannelId string

@description('Optional. Comma-separated list of email addresses whose messages always count as important.')
param importantSenders string = ''

@description('Optional. Comma-separated list of internal/in-org email domains for the Office 365 Users IN-ORG badge prefilter. Empty = look up every sender.')
param internalDomains string = ''

@description('Optional. Service Management Reference (e.g. a service tree GUID) attached to the Entra app registration. Required by some tenant policies — see https://aka.ms/service-management-reference-error.')
param serviceManagementReference string = ''

@description('When false, the Connector Namespace is referenced as existing and not re-PUT. Set to false (via `azd env set CREATE_CONNECTOR_NAMESPACE false`) after the first successful provision to make `azd up` idempotent.')
param createConnectorNamespace bool = true

var abbrs = loadJsonContent('./abbreviations.json')
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = { 'azd-env-name': environmentName }

var appServiceName = '${abbrs.webSitesAppService}${resourceToken}'
var appServicePlanName = '${abbrs.webServerFarms}${resourceToken}'
var webAppIdentityName = '${abbrs.managedIdentityUserAssignedIdentities}${resourceToken}'
var triggerIdentityName = '${abbrs.managedIdentityUserAssignedIdentities}trigger-${resourceToken}'
var resourceGroupName = '${abbrs.resourcesResourceGroups}${environmentName}'
var logAnalyticsName = '${abbrs.operationalInsightsWorkspaces}${resourceToken}'
var appInsightsName = '${abbrs.insightsComponents}${resourceToken}'
var connectorNamespaceName = '${abbrs.connectorNamespaces}${resourceToken}'
var connectorNamespaceConnectionName = '${abbrs.connectorNamespacesConnections}${resourceToken}'
var connectorNamespaceTeamsConnectionName = '${abbrs.connectorNamespacesConnections}teams-${resourceToken}'
var connectorNamespaceOffice365usersConnectionName = '${abbrs.connectorNamespacesConnections}o365users-${resourceToken}'
var entraAppUniqueName = 'app-${resourceToken}'

var monitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'

resource rg 'Microsoft.Resources/resourceGroups@2025-04-01' = {
  name: resourceGroupName
  location: location
  tags: tags
}

module logAnalytics 'br/public:avm/res/operational-insights/workspace:0.15.0' = {
  name: '${uniqueString(deployment().name, location)}-loganalytics'
  scope: rg
  params: {
    name: logAnalyticsName
    location: location
    tags: tags
    dataRetention: 30
  }
}

// User-assigned MI for the web app. It is: (a) the identity the connector clients use to
// call the 3 connections, (b) the FIC subject Easy Auth uses to mint client assertions
// (no client secret), and (c) the App Insights publisher.
module webAppUserAssignedIdentity 'br/public:avm/res/managed-identity/user-assigned-identity:0.5.0' = {
  name: 'webAppUserAssignedIdentity'
  scope: rg
  params: {
    location: location
    tags: tags
    name: webAppIdentityName
  }
}

// Dedicated identity attached to the Connector Namespace. The trigger uses it to mint the
// AAD bearer token attached to every callback to the App Service endpoint.
module triggerUserAssignedIdentity 'br/public:avm/res/managed-identity/user-assigned-identity:0.5.0' = {
  name: 'triggerUserAssignedIdentity'
  scope: rg
  params: {
    location: location
    tags: tags
    name: triggerIdentityName
  }
}

module monitoring 'br/public:avm/res/insights/component:0.7.1' = {
  name: '${uniqueString(deployment().name, location)}-appinsights'
  scope: rg
  params: {
    name: appInsightsName
    location: location
    tags: tags
    workspaceResourceId: logAnalytics.outputs.resourceId
    disableLocalAuth: true
    roleAssignments: [
      {
        roleDefinitionIdOrName: monitoringMetricsPublisherRoleId
        principalId: webAppUserAssignedIdentity.outputs.principalId
        principalType: 'ServicePrincipal'
      }
      {
        roleDefinitionIdOrName: monitoringMetricsPublisherRoleId
        principalId: userPrincipalId
        principalType: 'User'
      }
    ]
  }
}

// App Service plan (Linux). Basic B1 with alwaysOn so the app is warm to receive the
// connector callbacks. Easy Auth works on all tiers; Basic+ is needed for alwaysOn.
module appServicePlan 'br/public:avm/res/web/serverfarm:0.7.0' = {
  scope: rg
  name: appServicePlanName
  params: {
    name: appServicePlanName
    location: location
    tags: tags
    skuName: 'B1'
    reserved: true // Linux
  }
}

// Connector Namespace + 3 connections. The trigger UAMI is attached so the trigger can
// mint AAD tokens; Easy Auth on the web app validates those tokens.
module connectorNamespace './connectorNamespace.bicep' = {
  scope: rg
  name: connectorNamespaceName
  params: {
    name: connectorNamespaceName
    location: 'westcentralus' // Connector Namespace preview region.
    tags: tags
    connectionName: connectorNamespaceConnectionName
    teamsConnectionName: connectorNamespaceTeamsConnectionName
    office365usersConnectionName: connectorNamespaceOffice365usersConnectionName
    triggerIdentityResourceId: triggerUserAssignedIdentity.outputs.resourceId
    triggerIdentityPrincipalId: triggerUserAssignedIdentity.outputs.principalId
    webAppPrincipalId: webAppUserAssignedIdentity.outputs.principalId
    userPrincipalId: userPrincipalId
    createConnectorNamespace: createConnectorNamespace
  }
}

// Entra app registration that Easy Auth validates incoming tokens against. The web app MI
// federates against this app so Easy Auth needs no client secret.
module entraApp './app/entra.bicep' = {
  scope: rg
  name: 'entraApp'
  params: {
    appUniqueName: entraAppUniqueName
    appDisplayName: 'M365 Email Triage (App Service) (${appServiceName})'
    serviceManagementReference: serviceManagementReference
    managedIdentityPrincipalId: webAppUserAssignedIdentity.outputs.principalId
    functionAppHostname: '${appServiceName}.azurewebsites.net'
    tags: tags
  }
}

var allAppSettings = {
  APPLICATIONINSIGHTS_CONNECTION_STRING: monitoring.outputs.connectionString
  // Client id of the web app MI — used by DefaultAzureCredential for the connector calls
  // and by the Azure Monitor OpenTelemetry exporter.
  AZURE_CLIENT_ID: webAppUserAssignedIdentity.outputs.clientId
  // Connection runtime URLs the three connector SDK clients hit.
  OFFICE365_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.office365ConnectionRuntimeUrl
  TEAMS_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.teamsConnectionRuntimeUrl
  OFFICE365USERS_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.office365usersConnectionRuntimeUrl
  // Triage config.
  TEAMS_TEAM_ID: teamsTeamId
  TEAMS_CHANNEL_ID: teamsChannelId
  IMPORTANT_SENDERS: importantSenders
  INTERNAL_DOMAINS: internalDomains
  // Magic value: tells Easy Auth to use the named user-assigned MI to mint a federated
  // client assertion against the Entra app, in place of a client secret.
  OVERRIDE_USE_MI_FIC_ASSERTION_CLIENTID: webAppUserAssignedIdentity.outputs.clientId
}

module webApp 'br/public:avm/res/web/site:0.22.0' = {
  scope: rg
  name: appServiceName
  params: {
    name: appServiceName
    location: location
    tags: union(tags, { 'azd-service-name': 'web' })
    kind: 'app,linux'
    serverFarmResourceId: appServicePlan.outputs.resourceId
    httpsOnly: true
    managedIdentities: {
      userAssignedResourceIds: [
        '${webAppUserAssignedIdentity.outputs.resourceId}'
      ]
    }
    siteConfig: {
      linuxFxVersion: 'DOTNETCORE|10.0'
      alwaysOn: true
      ftpsState: 'Disabled'
      http20Enabled: true
    }
    configs: [
      {
        name: 'appsettings'
        properties: allAppSettings
      }
      {
        // Built-in authentication: every incoming request (including the connector
        // callback) must carry a valid Entra ID token whose audience matches our app and
        // whose caller object ID is the Connector Namespace's trigger UAMI.
        name: 'authsettingsV2'
        properties: {
          globalValidation: {
            requireAuthentication: true
            unauthenticatedClientAction: 'Return401'
            redirectToProvider: 'azureactivedirectory'
          }
          httpSettings: {
            requireHttps: true
            routes: {
              apiPrefix: '/.auth'
            }
            forwardProxy: {
              convention: 'NoProxy'
            }
          }
          identityProviders: {
            azureActiveDirectory: {
              enabled: true
              registration: {
                openIdIssuer: '${environment().authentication.loginEndpoint}${tenant().tenantId}/v2.0'
                clientId: entraApp.outputs.applicationId
                // FIC instead of a client secret — Easy Auth reads the user-assigned MI
                // from the named app setting and uses it to mint client assertions.
                clientSecretSettingName: 'OVERRIDE_USE_MI_FIC_ASSERTION_CLIENTID'
              }
              validation: {
                jwtClaimChecks: {}
                allowedAudiences: [
                  entraApp.outputs.applicationId
                  entraApp.outputs.identifierUri
                ]
                defaultAuthorizationPolicy: {
                  allowedPrincipals: {
                    // Only the Connector Namespace's trigger UAMI is allowed in.
                    // Tokens are matched by oid (the UAMI's principalId).
                    identities: [
                      triggerUserAssignedIdentity.outputs.principalId
                    ]
                  }
                }
              }
              isAutoProvisioned: false
            }
          }
          login: {
            tokenStore: {
              enabled: false
            }
            preserveUrlFragmentsForLogins: false
          }
          platform: {
            enabled: true
            runtimeVersion: '~1'
          }
        }
      }
    ]
  }
}

@description('The resource ID of the created Resource Group.')
output resourceGroupResourceId string = rg.id

@description('The name of the created Resource Group.')
output resourceGroupName string = rg.name

@description('The name of the created App Service (web app).')
output appServiceName string = webApp.outputs.name

@description('The default hostname of the created App Service.')
output appServiceDefaultHostname string = webApp.outputs.defaultHostname

@description('The name of the created Connector Namespace.')
output connectorNamespaceName string = connectorNamespace.outputs.name

@description('The name of the Office 365 connection on the Connector Namespace.')
output connectorNamespaceConnectionName string = connectorNamespace.outputs.connectionName

@description('The name of the Teams connection.')
output connectorNamespaceTeamsConnectionName string = connectorNamespace.outputs.teamsConnectionName

@description('The name of the Office 365 Users connection.')
output connectorNamespaceOffice365usersConnectionName string = connectorNamespace.outputs.office365usersConnectionName

@description('Route segment of the App Service endpoint receiving the connector callback.')
output office365EndpointName string = office365EndpointName

@description('App (client) ID of the Entra app registration Easy Auth validates against. Connector Namespace requests tokens for this audience.')
output entraAppClientId string = entraApp.outputs.applicationId

@description('Identifier URI of the Entra app registration (alternative audience value).')
output entraAppIdentifierUri string = entraApp.outputs.identifierUri

@description('Resource ID of the user-assigned MI attached to the Connector Namespace (referenced by the trigger config notificationDetails.authentication.identity).')
output triggerIdentityResourceId string = triggerUserAssignedIdentity.outputs.resourceId
