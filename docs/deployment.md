# Deployment guide

This document walks through deploying `calendar-matrix` end-to-end: provisioning
the Azure Static Web App with Bicep, registering an Entra ID (Azure AD) app
for MSAL sign-in, wiring up GitHub Actions secrets, and pushing to `main`.

## Prerequisites

- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) `>= 2.60`, logged in with `az login`
- An Azure subscription you can deploy resources into
- Owner/Contributor access on a resource group (or permission to create one)
- Admin access to the `raleec/calendar-matrix` GitHub repository (to add secrets)

## 1. Deploy the infrastructure

Create (or reuse) a resource group, then deploy `infra/main.bicep`:

```bash
az group create \
  --name rg-calendar-matrix \
  --location eastus2

az deployment group create \
  --resource-group rg-calendar-matrix \
  --template-file infra/main.bicep \
  --parameters name=calendar-matrix \
               location=eastus2 \
               repositoryUrl=https://github.com/raleec/calendar-matrix \
               branch=main
```

Alternatively, deploy using the checked-in parameter file:

```bash
az deployment group create \
  --resource-group rg-calendar-matrix \
  --template-file infra/main.bicep \
  --parameters infra/main.bicepparam
```

> This guide deploys the Static Web App without connecting Azure's managed
> GitHub integration — `repositoryUrl` is recorded for reference but
> deployment is instead done via the GitHub Actions workflow using
> `AZURE_STATIC_WEB_APPS_API_TOKEN` (step 4). If you'd rather let Azure manage
> the GitHub Actions workflow for you, pass a GitHub personal access token via
> the optional `repositoryToken` parameter (e.g.
> `--parameters repositoryToken=$GITHUB_PAT`); omitting it while still
> setting `repositoryUrl` is fine, since the template only forwards
> `repositoryUrl`/`branch` to Azure when a `repositoryToken` is supplied.

Capture the `defaultHostname` output — it's the hostname of the deployed
Static Web App (e.g. `calendar-matrix.azurestaticapps.net`):

```bash
az deployment group show \
  --resource-group rg-calendar-matrix \
  --name main \
  --query properties.outputs.defaultHostname.value \
  --output tsv
```

## 2. Register the app in Microsoft Entra ID

1. In the [Azure portal](https://portal.azure.com), go to **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it (e.g. `calendar-matrix`) and register.
3. Under **Authentication**, add a **Single-page application** platform.
4. Under **API permissions**, add the Microsoft Graph delegated permissions
   the app needs (e.g. `User.Read`, `Calendars.Read`) and grant admin consent
   if required by your tenant.
5. Note the **Application (client) ID** and **Directory (tenant) ID** from the
   **Overview** page — these become `VITE_AAD_CLIENT_ID` and
   `VITE_AAD_TENANT_ID`.

## 3. Add the redirect URI

Using the `defaultHostname` output from step 1, add both the production and
`www`-free URL as an SPA redirect URI on the app registration:

```
https://<defaultHostname>
```

For example: `https://calendar-matrix.azurestaticapps.net`.

Static Web Apps also creates preview environments for pull requests at URLs
like `https://<defaultHostname-branch-hash>.<region>.azurestaticapps.net`; add
those individually if you need MSAL sign-in to work on PR previews.

## 4. Get the Static Web Apps deployment token

```bash
az staticwebapp secrets list \
  --name calendar-matrix \
  --resource-group rg-calendar-matrix \
  --query properties.apiKey \
  --output tsv
```

## 5. Configure GitHub repository secrets and variables

In the GitHub repository, go to **Settings → Secrets and variables → Actions**.

Add as a **secret**:

| Secret name                          | Value                                            |
| ------------------------------------ | ------------------------------------------------- |
| `AZURE_STATIC_WEB_APPS_API_TOKEN`    | Output of the `az staticwebapp secrets list` command above |

Add as **variables** (not secrets — these are public client identifiers that
end up embedded in the compiled front-end bundle, not sensitive values):

| Variable name                        | Value                                            |
| ------------------------------------- | ------------------------------------------------- |
| `VITE_AAD_CLIENT_ID`                  | Application (client) ID from step 2                |
| `VITE_AAD_TENANT_ID`                  | Directory (tenant) ID from step 2                  |

These are consumed by [`.github/workflows/azure-static-web-apps.yml`](../.github/workflows/azure-static-web-apps.yml),
which builds the app with the `VITE_AAD_CLIENT_ID` / `VITE_AAD_TENANT_ID`
environment variables set and deploys it using the deployment token.

> The workflow sets `skip_deploy_on_missing_secrets: true`, so runs still
> succeed (building only, without deploying) before
> `AZURE_STATIC_WEB_APPS_API_TOKEN` has been added. Once the secret is
> configured, deployments happen automatically.

## 6. Push

- Pushing to `main` triggers `build_and_deploy_job`, which builds and deploys
  to the production environment.
- Opening a pull request against `main` triggers a build and deploy of a
  temporary preview environment; closing the pull request tears the preview
  environment down via `close_pull_request_job`.

## Verifying the deployment

1. Confirm the workflow run in the **Actions** tab succeeds.
2. Browse to `https://<defaultHostname>` and confirm the app loads.
3. Sign in and confirm the MSAL redirect completes successfully (no
   `AADSTS50011` redirect URI mismatch errors).
