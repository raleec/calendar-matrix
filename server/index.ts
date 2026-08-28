/**
 * Local Graph proxy server for development without Power Platform.
 *
 * Obtains Microsoft Graph access tokens via the `az` CLI (`az account
 * get-access-token`) so no additional Entra app registration is needed.
 * The server runs on port 3001; Vite proxies `/api/*` to it in dev mode.
 *
 * Start with: `npm run server`  (or `npm run local` to start both together)
 */

import { execSync } from 'child_process'
import express, { type Request, type Response } from 'express'

const PORT = 3001
const GRAPH = 'https://graph.microsoft.com/v1.0'

// ---------------------------------------------------------------------------
// Token management
// ---------------------------------------------------------------------------

interface CachedToken {
  token: string
  expiresAt: number // ms epoch
}

let tokenCache: CachedToken | null = null
const REFRESH_BUFFER_MS = 60_000 // refresh 60 s before expiry

function getGraphToken(): string {
  const now = Date.now()
  if (tokenCache && tokenCache.expiresAt - REFRESH_BUFFER_MS > now) {
    return tokenCache.token
  }

  try {
    const raw = execSync(
      'az account get-access-token --resource https://graph.microsoft.com --output json',
      { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] },
    )
    const parsed = JSON.parse(raw) as { accessToken: string; expiresOn: string }
    tokenCache = {
      token: parsed.accessToken,
      expiresAt: new Date(parsed.expiresOn).getTime(),
    }
    return tokenCache.token
  } catch (err) {
    throw new Error(
      `Failed to get Graph token via az CLI. Make sure you are signed in with "az login".\n${String(err)}`,
    )
  }
}

// ---------------------------------------------------------------------------
// Graph proxy helper
// ---------------------------------------------------------------------------

async function graphGet<T>(path: string, extraHeaders?: Record<string, string>): Promise<T> {
  const token = getGraphToken()
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
  const token = getGraphToken()
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
  console.log('Fetching initial token...')
  try {
    getGraphToken()
    console.log('Token OK — ready.')
  } catch (err) {
    console.error('Token fetch failed:', err instanceof Error ? err.message : err)
    console.error('Run "az login" then restart the server.')
  }
})
