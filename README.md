# Calendar Matrix

Azure Static Web App to view team out-of-office / calendar availability in a
month-grid matrix, powered by Microsoft Graph.

Rows are people, columns are the days of the selected month, and cells are
colour-coded by status (V = Vacation, P = Personal Leave, T = Travel,
WE = Working Elsewhere), with per-person and per-day totals.

This repository currently contains the project skeleton only: a Vite + React +
TypeScript app with placeholder components. Authentication (MSAL), the people
picker, the Microsoft Graph `getSchedule` data layer, and the full grid are
tracked in separate issues.

## Local development

```bash
npm install
npm run dev
```

The app is served at http://localhost:5173.

Useful scripts:

| Script                 | Purpose                                |
| ---------------------- | -------------------------------------- |
| `npm run dev`          | Start the Vite dev server              |
| `npm run build`        | Type-check and build into `dist/`      |
| `npm run preview`      | Preview the production build           |
| `npm run lint`         | Run ESLint                             |
| `npm run format`       | Format the repository with Prettier    |
| `npm run format:check` | Check formatting without writing files |

## Running under the Static Web Apps emulator

To exercise `staticwebapp.config.json` (SPA fallback routing, headers) locally:

```bash
npm run build
npx @azure/static-web-apps-cli start dist --run "npm run dev"
```

The emulator is served at http://localhost:4280. Deep links such as `/matrix`
are rewritten to `/index.html` rather than returning a 404.

## Configuration

Copy `.env.example` to `.env.local` and fill in the values for your Entra ID app
registration:

```bash
cp .env.example .env.local
```

| Variable             | Purpose                                     |
| -------------------- | ------------------------------------------- |
| `VITE_AAD_CLIENT_ID` | Application (client) ID of the SPA          |
| `VITE_AAD_TENANT_ID` | Directory (tenant) ID the app is registered |

`.env` and `.env.local` are gitignored — never commit real tenant identifiers.

> **Note:** a single-tenant Entra ID app registration (SPA platform, with the
> local dev and deployed redirect URIs plus consented delegated Graph scopes) is
> required before sign-in will work. Until that exists and the MSAL work lands,
> the app renders a static shell only.

## Deployment

`.github/workflows/azure-static-web-apps.yml` builds the app and deploys it to
Azure Static Web Apps on pushes to `main`, and creates preview environments for
pull requests. It requires:

- a repository secret `AZURE_STATIC_WEB_APPS_API_TOKEN` (the deployment token of
  the Static Web App resource)
- repository variables `VITE_AAD_CLIENT_ID` and `VITE_AAD_TENANT_ID`, which are
  passed to the build as environment variables

See [`docs/deployment.md`](docs/deployment.md) for the end-to-end guide to
provisioning the Azure infrastructure (`infra/main.bicep`), registering the
Entra ID app, and wiring up the GitHub Actions deployment pipeline.
