"""FastAPI app for workbook-backed semantic presentation search."""

from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from vector_search_app.db import (
    count_presentations,
    ensure_schema,
    fetch_presentations,
    search_presentations,
)
from vector_search_app.embeddings import embed_text
from vector_search_app.models import SearchRequest
from vector_search_app.service import index_uploaded_workbook
from vector_search_app.settings import get_auto_index_on_startup

STATIC_DIR = Path(__file__).with_name("static")


@asynccontextmanager
async def lifespan(_: FastAPI):
    ensure_schema()
    yield


app = FastAPI(title="FKM Vector Search", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/assets", StaticFiles(directory=str(STATIC_DIR)), name="assets")


@app.get("/")
def root() -> RedirectResponse:
    return RedirectResponse(url="/index.html", status_code=307)


@app.get("/index.html")
def index_html() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health() -> dict[str, str | int]:
    return {"status": "ok", "indexed_rows": count_presentations()}


@app.get("/api/presentations")
def presentations(limit: int = 10, offset: int = 0) -> dict[str, object]:
    total = count_presentations()
    items = fetch_presentations(limit=limit, offset=offset)
    columns: list[str] = []
    if items:
        columns = list(items[0].metadata.keys())
    return {
        "items": [item.model_dump() for item in items],
        "columns": columns,
        "total": total,
        "limit": limit,
        "offset": offset,
    }


@app.post("/api/index")
async def index_endpoint(
    workbook: UploadFile = File(...),
    worksheet_name: str | None = Form(default=None),
) -> dict[str, int | str]:
    if not workbook.filename:
        raise HTTPException(status_code=400, detail="Uploaded workbook needs a filename.")
    if not workbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx workbook uploads are supported.")

    try:
        workbook_bytes = await workbook.read()
        if not workbook_bytes:
            raise HTTPException(status_code=400, detail="Uploaded workbook is empty.")
        return index_uploaded_workbook(
            workbook_bytes,
            workbook.filename,
            worksheet_name,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/search")
def semantic_search(request: SearchRequest) -> dict[str, object]:
    if count_presentations() == 0:
        raise HTTPException(
            status_code=400,
            detail="No indexed rows found. Run indexing first.",
        )

    query_embedding = embed_text(request.query)
    results = search_presentations(query_embedding, request.top_k)
    return {"items": [item.model_dump() for item in results]}
