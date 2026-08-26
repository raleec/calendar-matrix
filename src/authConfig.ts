import type { Configuration } from '@azure/msal-browser'

const clientId = import.meta.env.VITE_AAD_CLIENT_ID
const tenantId = import.meta.env.VITE_AAD_TENANT_ID

/**
 * MSAL configuration for the single-page app registration.
 *
 * The authority is scoped to the single tenant the app is registered in
 * (rather than `common`/`organizations`), and the token cache lives in
 * `sessionStorage` so signed-in state does not persist across browser tabs
 * or after the tab is closed.
 */
export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: '/',
    postLogoutRedirectUri: '/',
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
}

/**
 * Delegated Microsoft Graph scopes requested at sign-in.
 */
export const loginRequest = {
  scopes: [
    'User.Read',
    'User.ReadBasic.All',
    'Calendars.Read.Shared',
    'GroupMember.Read.All',
    'People.Read',
  ],
}

export const graphConfig = {
  meEndpoint: 'https://graph.microsoft.com/v1.0/me',
  photoEndpoint: 'https://graph.microsoft.com/v1.0/me/photo/$value',
}
