# SharePoint Knowledge Base Front End

This repository includes a static, Azure Static Web Apps-ready knowledge base front end that reads synchronized SharePoint output from:

- `sharepoint_sync/index.json`
- `sharepoint_sync/content/**/*.json`

## Local preview

Because the app fetches JSON files, serve the repository over HTTP (do not open `index.html` directly from disk).

### Option 1: Python

```bash
python -m http.server 8080
```

Open `http://localhost:8080`.

### Option 2: Node (if available)

```bash
npx serve .
```

## Azure Static Web Apps deployment notes

- This app is plain static HTML/CSS/JS and supports an optional Azure Static Web Apps API.

- This repo now includes an Azure Static Web Apps API endpoint at `api/contact` (served at `/api/contact`) with GET and POST handling for quick browser validation and form submissions.
- `staticwebapp.config.json` is included for SPA deep-link support (for routes like `/doc/{id}`) and navigation fallback.
- In Azure Static Web Apps, set app location to the repository root (`/`) and output location to empty (or `/`, depending on portal input).
- Ensure your SharePoint sync workflow continues committing updated JSON files under `sharepoint_sync/`.

## Deployment runbook

Use this checklist when you want to deploy a front-end change (including removing debug UI elements).

1. **Create and merge your PR into `main`.**
   - This repo deploys to Azure Static Web Apps from pushes to `main`.
2. **Confirm the Azure deploy workflow runs successfully.**
   - Workflow file: `.github/workflows/azure-static-web-apps-orange-plant-05fa93210.yml`
   - Trigger: `push` on `main`
3. **Verify the production site after deployment.**
   - Hard refresh the page (Ctrl/Cmd+Shift+R) to avoid stale JS.
   - Check the document list, filters, and `/doc/{id}` deep links.
4. **If content looks stale, run the SharePoint sync workflow manually.**
   - Workflow file: `.github/workflows/sync.yml`
   - Trigger: `workflow_dispatch` (manual) or every 5 minutes (scheduled cron).
   - This regenerates `sharepoint_sync/index.json` and `sharepoint_sync/content/**/*.json`.
5. **Re-check production once sync commits are pushed.**
   - Any new commit to `main` from the sync workflow triggers a fresh static app deployment.

## Front-end behavior summary

- Loads index records from `sharepoint_sync/index.json`.
- Displays searchable document list with:
  - search (name, path, extracted content once indexed)
  - drive filter
  - file type filter
  - sort options (name and last modified)
  - pagination for larger document sets
- Shows selected document details and content in a viewer panel.
- Handles missing/unsupported extracted content gracefully.
- Includes copy-link support for document deep links.
- Uses query-parameter links (`/?doc={id}`) for shareable links to avoid 404s on hosts without SPA rewrites, while still supporting legacy `/doc/{id}` links.
