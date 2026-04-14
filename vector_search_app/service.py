"""Shared indexing workflow for workbook-to-pgvector sync."""

from vector_search_app.db import upsert_presentations
from vector_search_app.embeddings import embed_texts
from vector_search_app.excel_loader import load_index_rows


def index_uploaded_workbook(
    workbook_bytes: bytes,
    workbook_name: str,
    worksheet_name: str | None = None,
) -> dict[str, int | str]:
    rows = load_index_rows(workbook_bytes, workbook_name, worksheet_name)
    if not rows:
        return {"indexed_rows": 0, "workbook_name": workbook_name}

    embeddings = embed_texts([row.searchable_text for row in rows])
    indexed_count = upsert_presentations(rows, embeddings)
    return {"indexed_rows": indexed_count, "workbook_name": workbook_name}
