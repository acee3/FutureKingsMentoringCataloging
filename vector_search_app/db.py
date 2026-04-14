"""Database helpers for pgvector-backed presentation search."""

from collections.abc import Sequence
import json

import psycopg
from psycopg.rows import dict_row

from vector_search_app.models import PresentationRecord, SearchResult
from vector_search_app.settings import get_database_url, get_embedding_dimension
from vector_search_app.types import IndexedWorkbookRow


def _vector_literal(values: Sequence[float]) -> str:
    return "[" + ",".join(f"{value:.12f}" for value in values) + "]"


def get_connection() -> psycopg.Connection:
    return psycopg.connect(get_database_url(), row_factory=dict_row)


def ensure_schema() -> None:
    dimension = get_embedding_dimension()
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute("CREATE EXTENSION IF NOT EXISTS vector")
            cursor.execute(
                f"""
                CREATE TABLE IF NOT EXISTS presentations (
                    source_id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    workbook_name TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    row_number INTEGER NOT NULL,
                    metadata JSONB NOT NULL,
                    searchable_text TEXT NOT NULL,
                    embedding VECTOR({dimension}) NOT NULL,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS presentations_sheet_row_idx
                ON presentations (sheet_name, row_number)
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS presentations_embedding_idx
                ON presentations
                USING ivfflat (embedding vector_cosine_ops)
                WITH (lists = 100)
                """
            )
        connection.commit()


def upsert_presentations(
    rows: Sequence[IndexedWorkbookRow],
    embeddings: Sequence[Sequence[float]],
) -> int:
    if len(rows) != len(embeddings):
        raise ValueError("rows and embeddings must have matching lengths")
    if not rows:
        return 0

    with get_connection() as connection:
        with connection.cursor() as cursor:
            for row, embedding in zip(rows, embeddings, strict=True):
                cursor.execute(
                    """
                    INSERT INTO presentations (
                        source_id,
                        title,
                        workbook_name,
                        sheet_name,
                        row_number,
                        metadata,
                        searchable_text,
                        embedding,
                        updated_at
                    )
                    VALUES (
                        %(source_id)s,
                        %(title)s,
                        %(workbook_name)s,
                        %(sheet_name)s,
                        %(row_number)s,
                        %(metadata)s::jsonb,
                        %(searchable_text)s,
                        %(embedding)s::vector,
                        NOW()
                    )
                    ON CONFLICT (source_id) DO UPDATE SET
                        title = EXCLUDED.title,
                        workbook_name = EXCLUDED.workbook_name,
                        sheet_name = EXCLUDED.sheet_name,
                        row_number = EXCLUDED.row_number,
                        metadata = EXCLUDED.metadata,
                        searchable_text = EXCLUDED.searchable_text,
                        embedding = EXCLUDED.embedding,
                        updated_at = NOW()
                    """,
                    {
                        "source_id": row.source_id,
                        "title": row.title,
                        "workbook_name": row.workbook_name,
                        "sheet_name": row.sheet_name,
                        "row_number": row.row_number,
                        "metadata": json.dumps(row.metadata),
                        "searchable_text": row.searchable_text,
                        "embedding": _vector_literal(embedding),
                    },
                )
        connection.commit()

    return len(rows)


def fetch_presentations(limit: int = 10, offset: int = 0) -> list[PresentationRecord]:
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT source_id, title, workbook_name, sheet_name, row_number, metadata
                FROM presentations
                ORDER BY sheet_name, row_number
                LIMIT %s
                OFFSET %s
                """,
                (limit, offset),
            )
            results = cursor.fetchall()

    return [PresentationRecord(**row) for row in results]


def count_presentations() -> int:
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) AS count FROM presentations")
            row = cursor.fetchone()
    return int(row["count"])


def search_presentations(
    query_embedding: Sequence[float],
    top_k: int,
) -> list[SearchResult]:
    vector = _vector_literal(query_embedding)
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT
                    source_id,
                    title,
                    workbook_name,
                    sheet_name,
                    row_number,
                    metadata,
                    1 - (embedding <=> %s::vector) AS score
                FROM presentations
                ORDER BY embedding <=> %s::vector
                LIMIT %s
                """,
                (vector, vector, top_k),
            )
            results = cursor.fetchall()

    return [SearchResult(**row) for row in results]
