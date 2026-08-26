# App registration for MSAL sign-in

This app authenticates users with [MSAL for React](https://github.com/AzureAD/microsoft-authentication-library-for-js)
against a **single-tenant** Microsoft Entra ID (Azure AD) app registration.
This document covers registering that app; see
[`docs/deployment.md`](deployment.md) for the full infrastructure/CI story.

## 1. Create the registration

1. In the [Azure portal](https://portal.azure.com), go to **Microsoft Entra ID**
   → **App registrations** → **New registration**.
2. Name it (e.g. `calendar-matrix`).
3. Under **Supported account types**, choose **Accounts in this organizational
   directory only (Single tenant)**.
4. Leave **Redirect URI** blank for now — it's added as a platform in the next
   step — and select **Register**.

## 2. Add the SPA platform and redirect URIs

1. Open the new registration → **Authentication** → **Add a platform** →
   **Single-page application**.
2. Add a redirect URI for each environment that needs to sign in:
   - `http://localhost:5173` — Vite dev server (`npm run dev`)
   - `http://localhost:4280` — Static Web Apps CLI emulator
     (`npx @azure/static-web-apps-cli start`)
   - `https://<defaultHostname>` — the production Static Web App hostname
     (see `docs/deployment.md`, e.g. `https://calendar-matrix.azurestaticapps.net`)
   - one entry per PR preview environment hostname, if MSAL sign-in needs to
     work on previews (Static Web Apps creates a new hostname per PR)
3. Do **not** create a client secret — this is a public client (browser SPA)
   and MSAL never uses one. The "Certificates & secrets" blade should stay
   empty.

## 3. Add delegated API permissions

Under **API permissions** → **Add a permission** → **Microsoft Graph** →
**Delegated permissions**, add:

- `User.Read`
- `User.ReadBasic.All`
- `Calendars.Read.Shared`
- `GroupMember.Read.All`
- `People.Read`

Select **Grant admin consent for `<tenant>`** so users aren't prompted to
consent individually (some of these permissions require admin consent in
most tenants).

## 4. Record the identifiers

From the registration's **Overview** page, note:

- **Application (client) ID** → `VITE_AAD_CLIENT_ID`
- **Directory (tenant) ID** → `VITE_AAD_TENANT_ID`

Use these to populate `.env.local` for local development (copy
`.env.example`) and the `VITE_AAD_CLIENT_ID` / `VITE_AAD_TENANT_ID` GitHub
Actions repository variables for deployment — see
[`docs/deployment.md`](deployment.md#5-configure-github-repository-secrets-and-variables).

Both values are public client identifiers embedded in the compiled front-end
bundle; they are configured as repository **variables**, not secrets, and no
client secret exists anywhere in this repository or the app registration.
