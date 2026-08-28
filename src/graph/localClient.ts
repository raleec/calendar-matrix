/**
 * Local-mode API client — used when the app is running outside Power Apps
 * (e.g. `npm run local` / `npm run dev`).
 *
 * Uses @azure/msal-browser for auth (popup flow, Microsoft SSO) and calls
 * Microsoft Graph directly from the browser — no Express proxy required.
 * Works with Microsoft corporate CA policies because the browser handles
 * the full interactive sign-in flow.
 */

import * as msal from '@azure/msal-browser'
import type { GraphGroup, GraphPerson } from './types'
import type { GraphScheduleInformation, TimeWindow } from './schedule'

// Client ID for local MSAL browser auth.
// Set VITE_LOCAL_CLIENT_ID in .env.local to use your own app registration.
// See README §"Local mode" for setup instructions.
const CLIENT_ID: string = import.meta.env.VITE_LOCAL_CLIENT_ID as string || ''
const TENANT_ID = '72f988bf-86f1-41af-91ab-2d7cd011db47' // Microsoft corp tenant
const GRAPH = 'https://graph.microsoft.com/v1.0'

const SCOPES = [
  'Calendars.Read',
  'User.Read',
  'User.Read.All',
  'Group.Read.All',
]

// ---------------------------------------------------------------------------
// MSAL browser setup (lazy-initialised, singleton)
// ---------------------------------------------------------------------------

let _msalApp: msal.PublicClientApplication | null = null

async function getMsalApp(): Promise<msal.PublicClientApplication> {
  if (!CLIENT_ID) {
    throw new Error(
      'VITE_LOCAL_CLIENT_ID is not set.\n\n' +
      'Create an Entra app registration and add it to .env.local:\n' +
      '  VITE_LOCAL_CLIENT_ID=<your-app-client-id>\n\n' +
      'See README § "Local mode" for step-by-step instructions.',
    )
  }
  if (!_msalApp) {
    _msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        // The Microsoft-native loopback redirect — registered in the app for
        // native/desktop use. Popup flow uses this URI so no custom page needed.
        redirectUri: 'https://login.microsoftonline.com/common/oauth2/nativeclient',
      },
      cache: { cacheLocation: 'sessionStorage' },
    })
    await _msalApp.initialize()
    // Handle any redirect response on page load
    await _msalApp.handleRedirectPromise()
  }
  return _msalApp
}

async function getToken(): Promise<string> {
  const app = await getMsalApp()
  const accounts = app.getAllAccounts()

  if (accounts.length > 0) {
    try {
      const result = await app.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] })
      return result.accessToken
    } catch {
      // Fall through to popup
    }
  }

  const result = await app.acquireTokenPopup({ scopes: SCOPES })
  return result.accessToken
}

// ---------------------------------------------------------------------------
// Graph fetch helper
// ---------------------------------------------------------------------------

async function graphFetch<T>(
  path: string,
  init?: RequestInit & { extraHeaders?: Record<string, string> },
): Promise<T> {
  const token = await getToken()
  const res = await fetch(`${GRAPH}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init?.extraHeaders ?? {}),
    },
  })
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText)
    throw new Error(`Graph ${init?.method ?? 'GET'} ${path} failed (${res.status}): ${text}`)
  }
  return res.json() as Promise<T>
}

/**
 * Returns true only when explicitly running in standalone local mode
 * (`npm run local`). When using `pac code run` the app still runs on
 * localhost, but VITE_LOCAL_MODE is not set so connectors are used normally.
 */
export function isLocalMode(): boolean {
  return import.meta.env.VITE_LOCAL_MODE === 'true'
}

// ---------------------------------------------------------------------------
// Identity
// ---------------------------------------------------------------------------

export interface LocalMe {
  id: string
  displayName: string
  userPrincipalName: string
  mail?: string
}

export function getLocalMe(): Promise<LocalMe> {
  return graphFetch<LocalMe>('/me?$select=id,displayName,userPrincipalName,mail')
}

// ---------------------------------------------------------------------------
// People
// ---------------------------------------------------------------------------

export async function localSearchUsers(query: string, top = 15): Promise<GraphPerson[]> {
  const filter = `"displayName:${query}" OR "mail:${query}"`
  const data = await graphFetch<{ value: Array<{ id?: string; displayName?: string; mail?: string }> }>(
    `/users?$search=${encodeURIComponent(filter)}&$top=${top}&$select=id,displayName,mail`,
    { extraHeaders: { ConsistencyLevel: 'eventual' } },
  )
  return (data.value ?? [])
    .filter((u): u is typeof u & { id: string } => !!u.id)
    .map((u) => ({ id: u.id, displayName: u.displayName ?? u.mail ?? u.id, mail: u.mail ?? null }))
}

export async function localSearchGroups(query: string, top = 15): Promise<GraphGroup[]> {
  const safe = query.replace(/'/g, "''")
  const data = await graphFetch<{ value: Array<{ id?: string; displayName?: string }> }>(
    `/groups?$filter=startswith(displayName,'${safe}')&$top=${top}&$select=id,displayName`,
  )
  return (data.value ?? [])
    .filter((g): g is typeof g & { id: string } => !!g.id)
    .map((g) => ({ id: g.id, displayName: g.displayName ?? g.id }))
}

export async function localGetGroupMembers(groupId: string): Promise<GraphPerson[]> {
  const members: GraphPerson[] = []
  let nextLink: string | undefined

  do {
    const url = nextLink
      ? nextLink.replace(GRAPH, '')
      : `/groups/${groupId}/members?$top=999&$select=id,displayName,mail`
    // eslint-disable-next-line no-await-in-loop
    const data = await graphFetch<{
      value?: Array<{ id?: string; displayName?: string; mail?: string }>
      '@odata.nextLink'?: string
    }>(url)
    for (const m of data.value ?? []) {
      if (m.id) members.push({ id: m.id, displayName: m.displayName ?? m.mail ?? m.id, mail: m.mail ?? null })
    }
    nextLink = data['@odata.nextLink']
  } while (nextLink)

  return members
}

export async function localGetDirectReports(userId: string): Promise<GraphPerson[]> {
  const data = await graphFetch<{ value?: Array<{ id?: string; displayName?: string; mail?: string }> }>(
    `/users/${userId}/directReports?$select=id,displayName,mail`,
  )
  return (data.value ?? [])
    .filter((u): u is typeof u & { id: string } => !!u.id)
    .map((u) => ({ id: u.id, displayName: u.displayName ?? u.mail ?? u.id, mail: u.mail ?? null }))
}

// ---------------------------------------------------------------------------
// Schedule
// ---------------------------------------------------------------------------

export async function localGetSchedule(
  schedules: string[],
  timeWindow: TimeWindow,
): Promise<GraphScheduleInformation[]> {
  const data = await graphFetch<{ value?: GraphScheduleInformation[] }>('/me/calendar/getSchedule', {
    method: 'POST',
    extraHeaders: { Prefer: 'outlook.timezone="UTC"' },
    body: JSON.stringify({
      schedules,
      startTime: timeWindow.startTime,
      endTime: timeWindow.endTime,
      availabilityViewInterval: 1440,
    }),
  })
  return data.value ?? []
}
