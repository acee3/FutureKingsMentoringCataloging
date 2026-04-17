"""Azure App Service entrypoint for the catalog app."""

import os

import uvicorn

from catalog_app.api import app


if __name__ == "__main__":
    port = int(os.getenv("PORT") or os.getenv("WEBSITES_PORT") or "8000")
    uvicorn.run(app, host="0.0.0.0", port=port)
