param name string
param location string
param tags object = {}

@description('Office 365 Outlook connection name (drives the trigger + sender history + flag).')
param connectionName string
@description('Microsoft Teams connection name (post triage card).')
param teamsConnectionName string
@description('Office 365 Users connection name (IN-ORG badge + manager enrichment).')
param office365usersConnectionName string

@description('Resource ID of the user-assigned managed identity attached to the Connector Namespace. The trigger uses this identity to mint the AAD token it attaches to the App Service callback.')
param triggerIdentityResourceId string
@description('Object (principal) ID of the trigger UAMI, granted access to the office365 connection so the runtime can read the mailbox on the trigger config\'s behalf.')
param triggerIdentityPrincipalId string

@description('Object (principal) ID of the App Service (web app) user-assigned MI. Granted access to all three connections so the app can call them at runtime (GetEmails / Flag / PostMessage / UserProfile).')
param webAppPrincipalId string

@description('Optional. AAD object ID of a user (typically the deployer) to also grant access to the connections, so the same code can be debugged locally with `az login` credentials.')
param userPrincipalId string = ''
param tenantId string = tenant().tenantId

@description('When false, reference the namespace as existing instead of creating/updating it. Required workaround for the Connector Namespace RP rejecting identity in update PUTs ("ManagedIdentityInvalid: user assigned identities can not be changed"), even when the body is byte-for-byte identical to current state. Set to false on the second+ provision by running: azd env set CREATE_CONNECTOR_NAMESPACE false')
param createConnectorNamespace bool = true

resource newConnectorNamespace 'Microsoft.Web/connectorGateways@2026-05-01-preview' = if (createConnectorNamespace) {
  name: name
  location: location
  tags: tags
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${triggerIdentityResourceId}': {}
    }
  }
}

resource existingConnectorNamespace 'Microsoft.Web/connectorGateways@2026-05-01-preview' existing = if (!createConnectorNamespace) {
  name: name
}

// ---------------------------------------------------------------------------
// office365 connection — the trigger source, plus GetEmails + Flag client calls.
// ---------------------------------------------------------------------------
resource office365Connection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = {
  name: '${name}/${connectionName}'
  properties: {
    connectorName: 'office365'
  }
  dependsOn: createConnectorNamespace ? [ newConnectorNamespace ] : [ existingConnectorNamespace ]
}

// Web app MI -> office365 (GetEmails + Flag).
resource office365WebAppAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = {
  parent: office365Connection
  name: 'webapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: webAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

// Trigger UAMI -> office365. The namespace runtime impersonates this identity when
// reading from the mailbox on behalf of the trigger config.
resource office365TriggerAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = {
  parent: office365Connection
  name: 'trigger-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: triggerIdentityPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource office365UserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(userPrincipalId)) {
  parent: office365Connection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

// ---------------------------------------------------------------------------
// teams connection — outbound only (PostMessage), called by the web app MI.
// ---------------------------------------------------------------------------
resource teamsConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = {
  name: '${name}/${teamsConnectionName}'
  properties: {
    connectorName: 'teams'
  }
  dependsOn: createConnectorNamespace ? [ newConnectorNamespace ] : [ existingConnectorNamespace ]
}

resource teamsWebAppAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = {
  parent: teamsConnection
  name: 'webapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: webAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource teamsUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(userPrincipalId)) {
  parent: teamsConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

// ---------------------------------------------------------------------------
// office365users connection — outbound only (UserProfile + Manager), web app MI.
// ---------------------------------------------------------------------------
resource office365usersConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = {
  name: '${name}/${office365usersConnectionName}'
  properties: {
    connectorName: 'office365users'
  }
  dependsOn: createConnectorNamespace ? [ newConnectorNamespace ] : [ existingConnectorNamespace ]
}

resource office365usersWebAppAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = {
  parent: office365usersConnection
  name: 'webapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: webAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource office365usersUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(userPrincipalId)) {
  parent: office365usersConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

@description('The resource ID of the Connector Namespace.')
output resourceId string = createConnectorNamespace ? newConnectorNamespace.id : existingConnectorNamespace.id

@description('The name of the Connector Namespace.')
output name string = name

@description('The name of the Office 365 connection.')
output connectionName string = connectionName

@description('Runtime URL for the Office 365 connection.')
output office365ConnectionRuntimeUrl string = office365Connection.properties.connectionRuntimeUrl

@description('The name of the Teams connection.')
output teamsConnectionName string = teamsConnectionName

@description('Runtime URL for the Teams connection.')
output teamsConnectionRuntimeUrl string = teamsConnection.properties.connectionRuntimeUrl

@description('The name of the Office 365 Users connection.')
output office365usersConnectionName string = office365usersConnectionName

@description('Runtime URL for the Office 365 Users connection.')
output office365usersConnectionRuntimeUrl string = office365usersConnection.properties.connectionRuntimeUrl
