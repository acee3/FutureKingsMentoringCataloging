# fkm

Internal catalog app for FKM workshop PowerPoints.

The app scans configured SharePoint document libraries for `.pptx` files, extracts slide text, asks OpenAI to generate configured spreadsheet fields, stores the current catalog in SQLite, provides semantic search, and exports the current catalog as Excel.

## App Shape

- [`catalog_app/api.py`](catalog_app/api.py): FastAPI app, static UI, reload/search/export endpoints
- [`catalog_app/catalog_sync.py`](catalog_app/catalog_sync.py): manual SharePoint reload workflow
- [`catalog_app/db/`](catalog_app/db): schema and current-state SQLite helpers
- [`catalog_app/generation/configuration.py`](catalog_app/generation/configuration.py): SharePoint sources and official Excel column configuration
- [`catalog_app/generation/`](catalog_app/generation): Microsoft Graph, PowerPoint parsing, OpenAI metadata, and Excel export helpers
- [`catalog_app/static/`](catalog_app/static): browser UI

## Current-State Storage

This app is designed for a small Azure App Service deployment, so SQLite stores current state only, not history.

- `presentations` stores one row per current PowerPoint. If Microsoft Graph reports a file was deleted or removed from the configured folders, the row is hard-deleted so search/export do not show stale files and the database does not grow forever.
- `presentation_sources` stores one row per configured SharePoint source plus its latest Microsoft Graph delta link. Delta links let Reload ask Graph for changes since the last successful reload.
- `sync_status` contains exactly one row. Each Reload overwrites it with the latest status, counts, timestamps, and error. Historical reload runs are intentionally not retained.

The full spreadsheet row is stored in `presentations.metadata` as JSON text. Stable operational fields such as source key, drive ID, item ID, web URL, embedding, and timestamps also have normal columns so the app can update efficiently. Embeddings are stored as JSON arrays, and semantic search computes cosine similarity in Python.

## Local Run

Start the app with a persistent SQLite volume:

```bash
docker compose up --build
```

Open [http://localhost:8000](http://localhost:8000).

Useful local command without Docker:

```bash
uv run python app.py
```

Without Docker, the database defaults to [`catalog.sqlite3`](catalog.sqlite3) in the project root. Set `DATABASE_PATH` to use another SQLite file.

## Workflow

1. Click **Reload from SharePoint**.
2. The app uses Microsoft Graph delta queries to find changed/deleted PowerPoints.
3. Changed PowerPoints are parsed, classified with OpenAI, embedded, and upserted into SQLite.
4. Deleted or removed PowerPoints are hard-deleted from SQLite.
5. Use the table, semantic search, or **Export Excel**.

The first successful reload establishes the baseline delta links. Later reloads use those links to fetch only changes.

## Required Environment Variables

Application:

- `DATABASE_PATH`: SQLite database path, defaults to `catalog.sqlite3`
- `OPENAI_API_KEY`: OpenAI API key
- `EMBEDDING_MODEL`: optional, defaults to `text-embedding-3-small`

Microsoft Graph app credentials:

- `TENANT_ID`: Microsoft Entra tenant ID
- `CLIENT_ID`: app registration client ID used for Microsoft Graph
- `CLIENT_SECRET_VALUE`: app registration client secret value, not the secret ID
- `SITE_HOSTNAME`: SharePoint hostname, for example `contoso.sharepoint.com`
- `SITE_PATH`: SharePoint site path, for example `/sites/MySite`

## Azure App Service Authentication

Access control is handled by Azure App Service Authentication, not by FastAPI code.

In Azure App Service > Authentication, configure Microsoft as the identity provider:

- Tenant type: Workforce/current tenant
- Supported account types: Current tenant / single tenant
- Authentication: Require authentication
- Unauthenticated requests: HTTP 302 redirect
- Tenant requirement: Allow requests only from the issuer tenant

With that setup, requests from outside your organization are blocked before they reach the container.

## Azure Container Notes

The production image is built from [`Dockerfile.catalog-app`](Dockerfile.catalog-app).

The container starts:

```bash
uv run uvicorn catalog_app.api:app --host 0.0.0.0 --port ${PORT:-${WEBSITES_PORT:-8000}}
```

For Azure App Service custom containers, configure the app settings above and set the container port setting if your plan requires it.

For Docker Compose, the SQLite database is stored at `/data/catalog.sqlite3` inside the container and persisted in the `sqlite-data` volume. For Azure custom containers, mount persistent storage and set `DATABASE_PATH` to a path on that mount, such as `/data/catalog.sqlite3`.

For Azure App Service's Python build/runtime flow, use `python app.py` as the startup command. Set `DATABASE_PATH` to a persistent path such as `/home/data/catalog.sqlite3`.

## Changing The Catalog

Edit [`catalog_app/generation/configuration.py`](catalog_app/generation/configuration.py).

- `DRIVE_SOURCES` controls which SharePoint drives/folders are scanned.
- `get_presentation_columns(...)` controls the official Excel export columns.
- Direct columns use generators that read Microsoft Graph or parsed slide data.
- AI columns use `registry.ai_generator(field_name, output_type, description)`.

After changing sources or columns, run Reload from the web UI. Column changes may require regenerating many rows and can take time and OpenAI tokens.
