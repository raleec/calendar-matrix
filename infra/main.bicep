@description('Name of the Azure Static Web App resource.')
param name string

@description('Location of the Static Web App resource. Static Web Apps are only available in a subset of Azure regions.')
param location string = resourceGroup().location

@description('URL of the GitHub repository backing this Static Web App (e.g. https://github.com/raleec/calendar-matrix).')
param repositoryUrl string

@description('Branch of the repository the Static Web App tracks.')
param branch string = 'main'

@description('SKU tier for the Static Web App. Standard is required for custom auth / Microsoft Graph scenarios.')
param sku string = 'Standard'

@description('Optional GitHub personal access token (repo scope) used to configure the Azure-managed CI/CD integration for repositoryUrl/branch. Leave empty to skip that integration and deploy solely via the GitHub Actions workflow using AZURE_STATIC_WEB_APPS_API_TOKEN — required by the platform if repositoryUrl is set, otherwise the deployment will fail validation.')
@secure()
param repositoryToken string = ''

var useManagedIntegration = !empty(repositoryToken)

resource staticSite 'Microsoft.Web/staticSites@2023-12-01' = {
  name: name
  location: location
  sku: {
    name: sku
    tier: sku
  }
  properties: useManagedIntegration ? {
    repositoryUrl: repositoryUrl
    branch: branch
    repositoryToken: repositoryToken
    buildProperties: {
      appLocation: '/'
      outputLocation: 'dist'
    }
  } : {
    // Deployment is performed via the GitHub Actions workflow using the
    // AZURE_STATIC_WEB_APPS_API_TOKEN secret, so no build provider credentials
    // are configured here. Setting repositoryUrl without repositoryToken
    // would fail ARM validation, so both are omitted in this mode.
    buildProperties: {
      appLocation: '/'
      outputLocation: 'dist'
    }
  }
}

@description('Default hostname of the deployed Static Web App. Use this (prefixed with https://) as the MSAL redirect URI.')
output defaultHostname string = staticSite.properties.defaultHostname

@description('Resource ID of the Static Web App.')
output staticSiteId string = staticSite.id

@description('Name of the Static Web App resource.')
output staticSiteName string = staticSite.name
