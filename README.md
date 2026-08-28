# Calendar Matrix

A Microsoft Power Apps code app that renders a month-grid availability matrix
for your team, powered by Microsoft Graph. Rows are people, columns are days of
the selected month, and cells are colour-coded by the person's calendar status
(OOO, DTO, TR, WE) with per-person and per-day absence totals.

## Status codes

| Code | Label | Source |
|------|-------|--------|
| `OOO` | Out of Office | OOF calendar event (generic subject) |
| `DTO` | Vacation | OOF event with *vacation / PTO / annual leave / DTO* in the subject |
| `TR` | Travel | OOF event with *travel / trip / onsite / offsite* in the subject |
| `WE` | Working Elsewhere | "Working Elsewhere" calendar status |

Busy events are treated as free (shown blank). The `availabilityView` from the
Graph `getSchedule` API drives the base status; `scheduleItems` subject
heuristics further refine OOF days into `DTO` or `TR`.

## Architecture

This app runs entirely inside Power Platform as a **Power Apps code app** — no
separate Azure hosting, no Entra app registration, and no admin tenant consent
required. Authentication is handled transparently by the Power Platform
runtime.

```
Power Apps runtime
  └── calendar-matrix (code app)
        ├── Office 365 Users connector  →  user/group search, getSchedule (via HttpRequest)
        ├── Office 365 Outlook connector
        └── Office 365 Groups connector →  group search & membership expansion
```

The Vite build output (`dist/`) is packaged and deployed to Power Platform via
the [Power Apps CLI (`pac`)](https://learn.microsoft.com/power-platform/developer/pac-cli-overview).

## Local development

Prerequisites: [Node.js](https://nodejs.org) 18+, [pac CLI](https://learn.microsoft.com/power-platform/developer/pac-cli-overview).

```bash
npm install
```

Start the connection proxy (in one terminal):

```bash
pac code run
```

Start the Vite dev server (in another terminal):

```bash
npm run dev
```

Then open the local play URL shown by `pac code run` (e.g.
`https://apps.powerapps.com/play/{envId}/app/local?_localAppUrl=http://localhost:5173&_localConnectionUrl=http://localhost:8080`).

### Available scripts

| Script | Purpose |
|--------|---------|
| `npm run dev` | Start the Vite dev server (port 5173) |
| `npm run build` | Type-check and compile into `dist/` |
| `npm run preview` | Preview the production build locally |
| `npm run lint` | Run ESLint |
| `npm run format` | Format the repository with Prettier |
| `npm run format:check` | Check formatting without writing files |
| `npm run server` | Start the local Graph proxy (port 3001) |
| `npm run local` | Start both the proxy and Vite together (Scout / standalone mode) |

## Local mode (Scout / standalone)

The app can also run without Power Platform using only an active
[Azure CLI (`az`)](https://learn.microsoft.com/cli/azure/) session — no app
registration required.

### Prerequisites

[Node.js](https://nodejs.org) 18+ and `npm install` (already done if you followed the dev setup above). No Azure CLI or app registration needed.

### Starting local mode

```bash
npm run local
```

This starts two processes concurrently:
- **Express proxy** on port 3001 — authenticates via MSAL device-code flow
  (Microsoft Graph PowerShell public client — no app registration needed), then
  forwards Graph API calls with the cached token
- **Vite dev server** on port 5173 — proxies `/api/*` to the Express proxy

On first run the terminal will print a URL and one-time code to sign in:

```
──────────────────────────────────────────────────────────────
  Sign in to authorize Calendar Matrix (local mode)
──────────────────────────────────────────────────────────────
  To sign in, use a web browser to open https://microsoft.com/devicelogin
  and enter the code XXXXXXXXX to authenticate.
──────────────────────────────────────────────────────────────
```

After signing in the token is cached for the session (and refreshed automatically).

Open **http://localhost:5173** in your browser (or point Scout at that URL).

The app auto-detects it is running on `localhost` and routes all Graph calls
through the proxy instead of Power Apps connectors. No `pac code run` is
needed in this mode.

### How it works

```
Browser / Scout (localhost:5173)
  └── Vite dev server
        └── /api/* proxy → Express proxy (localhost:3001)
              └── az account get-access-token → Bearer token
                    └── Microsoft Graph API (graph.microsoft.com)
```

## Deployment

Build and push to the connected Power Platform environment in one step:

```bash
npm run build && pac code push
```

`pac code push` packages `dist/` and deploys to the environment you are
authenticated against (`pac auth list` / `pac auth select`).

> **Note:** The portal launch button on `make.powerapps.com` sometimes returns
> an HTTP 0 error for code apps. Use the direct play URL instead:
> `https://apps.powerapps.com/play/{environmentId}/app/{appId}?tenantId={tenantId}`

## Data & caching

- Schedule data is fetched from Graph via the **Office 365 Users** connector's
  `HttpRequest` operation, calling `https://graph.microsoft.com/v1.0/me/calendar/getSchedule`.
- Results are cached for **30 minutes** in both an in-memory map and
  `sessionStorage`, keyed by `personId:year:month`.
- The **Refresh** button invalidates the cache for all visible people and
  re-fetches immediately.
- The **Export CSV** button downloads the current grid as a UTF-8 CSV
  (BOM-prefixed for Excel compatibility).

## People picker

- Search by name or email (debounced 300 ms).
- Expand a group to add all its members at once.
- Load your own direct reports with one click.
- Selected people are persisted to `localStorage` and restored on reload.
- Individual people can be removed via their chip's × button; **Clear all**
  removes everyone at once.

## Project structure

```
src/
  components/
    PeoplePicker.tsx     # Search bar, result list, selected-people chips
    MatrixGrid.tsx       # Availability table with loading/error states, CSV export
  graph/
    people.ts            # User/group search via Office365 connectors
    schedule.ts          # getSchedule fetching, availabilityView parsing, OOF heuristics
    scheduleCache.ts     # Two-tier (memory + sessionStorage) cache with 30-min TTL
    types.ts             # GraphPerson, GraphGroup interfaces
  hooks/
    usePeopleSelection.ts  # Selected-people state, localStorage persistence
    useGraphToken.ts       # Power Apps context (user identity)
  utils/
    csv.ts               # RFC 4180 CSV builder + browser download helper
    date.ts              # Date arithmetic helpers
  status.ts              # StatusKey type, LEGEND, statusLabel
  generated/             # Auto-generated connector service classes (do not edit)
power.config.json        # Power Apps code app manifest (app ID, connector references)
vite.config.ts           # Vite config — base: './' is required for Power Apps CDN hosting
```

## Key constraints & gotchas

| Constraint | Detail |
|---|---|
| `base: './'` in `vite.config.ts` | **Required** — without it Vite emits absolute `/assets/` paths that break Power Apps CDN hosting |
| `Body` as plain object | The Office365Users `HttpRequest` schema declares `Body` as `object`; passing `JSON.stringify()` output causes double-encoding and an HTTP 400 |
| Absolute Graph URL | The connector path validator blocks paths with more than 2 segments; use the full `https://graph.microsoft.com/v1.0/…` URL |
| No `setConfig({})` call | Not needed in the Power Apps code app template |
| `pac code run` required locally | Without the connection proxy the connector calls fail silently |
