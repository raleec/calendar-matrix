/**
 * Local Graph proxy server for development without Power Platform.
 *
 * Uses MSAL WAM broker (Windows Account Manager) to acquire delegated Graph
 * tokens — the same mechanism Teams and Office use on corporate Windows
 * machines. No custom app registration required; WAM handles all CA policies
 * transparently via the OS credential store.
 *
 * Start with: `npm run server`  (or `npm run local` to start both together)
 */

import express, { type Request, type Response } from 'express'
import { PublicClientApplication, type Configuration, type AuthenticationResult } from '@azure/msal-node'
import { NativeBrokerPlugin } from '@azure/msal-node-extensions'

const PORT = 3001
const GRAPH = 'https://graph.microsoft.com/v1.0'

// Microsoft Graph Command Line Tools — well-known public client with
// pre-consented Calendar + User + Group scopes.
const CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
const TENANT_ID = '72f988bf-86f1-41af-91ab-2d7cd011db47' // Microsoft corp tenant

const SCOPES = [
  'https://graph.microsoft.com/Calendars.Read',
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/User.Read.All',
  'https://graph.microsoft.com/Group.Read.All',
  'offline_access',
]

// ---------------------------------------------------------------------------
// MSAL setup — WAM broker for transparent Windows SSO
// ---------------------------------------------------------------------------

const msalConfig: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
  broker: {
    nativeBrokerPlugin: new NativeBrokerPlugin(),
  },
}

const msalApp = new PublicClientApplication(msalConfig)

let cachedResult: AuthenticationResult | null = null
const REFRESH_BUFFER_MS = 5 * 60 * 1000

async function getGraphToken(): Promise<string> {
  // Try silent acquisition (WAM can often do this without any user interaction)
  const accounts = await msalApp.getTokenCache().getAllAccounts()
  if (accounts.length > 0) {
    try {
      const silent = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] })
      if (silent && silent.expiresOn && silent.expiresOn.getTime() - REFRESH_BUFFER_MS > Date.now()) {
        return silent.accessToken
      }
    } catch {
      // Fall through to interactive
    }
  }

  // WAM interactive — uses Windows OS auth dialog (handles all CA policies)
  console.log('\n' + '─'.repeat(60))
  console.log('  A Windows sign-in dialog should appear shortly.')
  console.log('  Check your taskbar if it does not come to the foreground.')
  console.log('─'.repeat(60) + '\n')
  const result = await msalApp.acquireTokenInteractive({
    scopes: SCOPES,
    // Null window handle — WAM shows dialog without a parent window
    windowHandle: Buffer.from(Int32Array.of(0).buffer),
  })

  if (!result) throw new Error('WAM auth returned no token')
  cachedResult = result
  console.log(`Signed in as ${result.account?.username ?? 'unknown'} — proxy ready.\n`)
  return result.accessToken
}

// ---------------------------------------------------------------------------
// Graph proxy helpers
// ---------------------------------------------------------------------------

async function graphGet<T>(path: string, extraHeaders?: Record<string, string>): Promise<T> {
  const token = await getGraphToken()
  const res = await fetch(`${GRAPH}${path}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...extraHeaders,
    },
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`Graph GET ${path} failed (${res.status}): ${text}`)
  }
  return res.json() as Promise<T>
}

async function graphPost<T>(path: string, body: unknown, extraHeaders?: Record<string, string>): Promise<T> {
  const token = await getGraphToken()
  const res = await fetch(`${GRAPH}${path}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Prefer': 'outlook.timezone="UTC"',
      ...extraHeaders,
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) {
    const text = await res.text()
    throw new Error(`Graph POST ${path} failed (${res.status}): ${text}`)
  }
  return res.json() as Promise<T>
}

// ---------------------------------------------------------------------------
// Express app
// ---------------------------------------------------------------------------

const app = express()
app.use(express.json())
app.use((_req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', 'http://localhost:5173')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  next()
})

function send(res: Response, fn: () => Promise<unknown>) {
  fn()
    .then((data) => res.json(data))
    .catch((err: unknown) => {
      const msg = err instanceof Error ? err.message : String(err)
      console.error(msg)
      res.status(500).json({ error: msg })
    })
}

// Current user identity
app.get('/api/me', (_req, res) => {
  send(res, () => graphGet('/me'))
})

// Schedule availability
app.post('/api/getSchedule', (req: Request, res: Response) => {
  send(res, () =>
    graphPost('/me/calendar/getSchedule', {
      ...req.body,
      availabilityViewInterval: 1440,
    }),
  )
})

// User typeahead search
app.get('/api/users/search', (req: Request, res: Response) => {
  const q = String(req.query.q ?? '').trim()
  const top = Number(req.query.top ?? 15)
  if (!q) { res.json({ value: [] }); return }
  const filter = encodeURIComponent(`"displayName:${q}" OR "mail:${q}"`)
  send(res, () =>
    graphGet(
      `/users?$search=${filter}&$top=${top}&$select=id,displayName,mail`,
      { ConsistencyLevel: 'eventual', '$count': 'true' },
    ),
  )
})

// Group typeahead search
app.get('/api/groups/search', (req: Request, res: Response) => {
  const q = String(req.query.q ?? '').trim()
  const top = Number(req.query.top ?? 15)
  if (!q) { res.json({ value: [] }); return }
  const safe = q.replace(/'/g, "''")
  send(res, () =>
    graphGet(
      `/groups?$filter=startswith(displayName,'${safe}')&$top=${top}&$select=id,displayName`,
    ),
  )
})

// Group members
app.get('/api/groups/:id/members', (req: Request, res: Response) => {
  const { id } = req.params
  const top = Number(req.query.top ?? 999)
  const skipToken = req.query.skipToken ? `&$skiptoken=${req.query.skipToken}` : ''
  send(res, () =>
    graphGet(`/groups/${id}/members?$top=${top}&$select=id,displayName,mail${skipToken}`),
  )
})

// Direct reports
app.get('/api/users/:id/directReports', (req: Request, res: Response) => {
  const { id } = req.params
  send(res, () =>
    graphGet(`/users/${id}/directReports?$select=id,displayName,mail`),
  )
})

app.listen(PORT, () => {
  console.log(`Graph proxy listening on http://localhost:${PORT}`)
  console.log('Authenticating with Microsoft Graph (browser)...')
  getGraphToken()
    .then(() => console.log('Proxy ready.\n'))
    .catch((err) => {
      console.error('Authentication failed:', err instanceof Error ? err.message : err)
      process.exit(1)
    })
})
