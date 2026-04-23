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

- This app is plain static HTML/CSS/JS and does not require a custom API.
- `staticwebapp.config.json` is included for SPA deep-link support (for routes like `/doc/{id}`) and navigation fallback.
- In Azure Static Web Apps, set app location to the repository root (`/`) and output location to empty (or `/`, depending on portal input).
- Ensure your SharePoint sync workflow continues committing updated JSON files under `sharepoint_sync/`.

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
