"""Current-state catalog persistence helpers."""

from collections.abc import Sequence
import json
import math

from catalog_app.db.connection import get_connection
from catalog_app.models import PresentationRecord, SearchResult, SyncStatus
from catalog_app.app_types import IndexedWorkbookRow


def _empty_to_none(value: str | None) -> str | None:
    return value or None


def _embedding_json(values: Sequence[float]) -> str:
    return json.dumps(list(values), separators=(",", ":"))


def _loads_metadata(value: str | bytes | bytearray | dict) -> dict:
    if isinstance(value, dict):
        return value
    loaded = json.loads(value)
    if not isinstance(loaded, dict):
        raise ValueError("presentation metadata must be a JSON object")
    return loaded


def _cosine_similarity(left: Sequence[float], right: Sequence[float]) -> float:
    if len(left) != len(right):
        return 0.0

    dot_product = 0.0
    left_magnitude = 0.0
    right_magnitude = 0.0
    for left_value, right_value in zip(left, right, strict=True):
        dot_product += left_value * right_value
        left_magnitude += left_value * left_value
        right_magnitude += right_value * right_value

    denominator = math.sqrt(left_magnitude) * math.sqrt(right_magnitude)
    if denominator == 0:
        return 0.0
    return dot_product / denominator


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
                source_key = row.source_key or ""
                drive_id = row.drive_id or ""
                item_id = row.item_id or str(row.metadata.get("id") or row.source_id)
                web_url = row.web_url or str(row.metadata.get("web_url") or "")
                cursor.execute(
                    """
                    INSERT INTO presentations (
                        source_id,
                        title,
                        workbook_name,
                        sheet_name,
                        row_number,
                        source_key,
                        drive_id,
                        item_id,
                        web_url,
                        last_modified_at,
                        metadata,
                        searchable_text,
                        embedding,
                        updated_at
                    )
                    VALUES (
                        :source_id,
                        :title,
                        :workbook_name,
                        :sheet_name,
                        :row_number,
                        :source_key,
                        :drive_id,
                        :item_id,
                        :web_url,
                        :last_modified_at,
                        :metadata,
                        :searchable_text,
                        :embedding,
                        CURRENT_TIMESTAMP
                    )
                    ON CONFLICT (source_id) DO UPDATE SET
                        title = EXCLUDED.title,
                        workbook_name = EXCLUDED.workbook_name,
                        sheet_name = EXCLUDED.sheet_name,
                        row_number = EXCLUDED.row_number,
                        source_key = EXCLUDED.source_key,
                        drive_id = EXCLUDED.drive_id,
                        item_id = EXCLUDED.item_id,
                        web_url = EXCLUDED.web_url,
                        last_modified_at = EXCLUDED.last_modified_at,
                        metadata = EXCLUDED.metadata,
                        searchable_text = EXCLUDED.searchable_text,
                        embedding = EXCLUDED.embedding,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    {
                        "source_id": row.source_id,
                        "title": row.title,
                        "workbook_name": row.workbook_name,
                        "sheet_name": row.sheet_name,
                        "row_number": row.row_number,
                        "source_key": _empty_to_none(source_key),
                        "drive_id": _empty_to_none(drive_id),
                        "item_id": _empty_to_none(item_id),
                        "web_url": _empty_to_none(web_url),
                        "last_modified_at": _empty_to_none(row.last_modified_at),
                        "metadata": json.dumps(row.metadata),
                        "searchable_text": row.searchable_text,
                        "embedding": _embedding_json(embedding),
                    },
                )
        connection.commit()

    return len(rows)


def delete_presentations(source_ids: Sequence[str]) -> int:
    """Hard-delete current catalog rows that no longer exist in SharePoint."""
    if not source_ids:
        return 0

    with get_connection() as connection:
        with connection.cursor() as cursor:
            placeholders = ",".join("?" for _ in source_ids)
            cursor.execute(
                f"DELETE FROM presentations WHERE source_id IN ({placeholders})",
                tuple(source_ids),
            )
            deleted_count = cursor.rowcount
        connection.commit()

    return deleted_count


def upsert_presentation_source(
    *,
    source_key: str,
    source_name: str,
    drive_id: str,
    folder_id: str | None = None,
    folder_path: str | None = None,
    delta_link: str | None = None,
) -> None:
    """Save one configured SharePoint source without retaining source history."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO presentation_sources (
                    source_key,
                    source_name,
                    drive_id,
                    folder_id,
                    folder_path,
                    delta_link,
                    updated_at
                )
                VALUES (
                    :source_key,
                    :source_name,
                    :drive_id,
                    :folder_id,
                    :folder_path,
                    :delta_link,
                    CURRENT_TIMESTAMP
                )
                ON CONFLICT (source_key) DO UPDATE SET
                    source_name = EXCLUDED.source_name,
                    drive_id = EXCLUDED.drive_id,
                    folder_id = EXCLUDED.folder_id,
                    folder_path = EXCLUDED.folder_path,
                    delta_link = EXCLUDED.delta_link,
                    updated_at = CURRENT_TIMESTAMP
                """,
                {
                    "source_key": source_key,
                    "source_name": source_name,
                    "drive_id": drive_id,
                    "folder_id": folder_id,
                    "folder_path": folder_path,
                    "delta_link": delta_link,
                },
            )
        connection.commit()


def get_source_delta_links() -> dict[str, str]:
    """Return saved Microsoft Graph delta links keyed by configured source."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT source_key, delta_link
                FROM presentation_sources
                WHERE delta_link IS NOT NULL
                """
            )
            rows = cursor.fetchall()

    return {str(row["source_key"]): str(row["delta_link"]) for row in rows}


def update_sync_status(
    *,
    status: str,
    total_items: int = 0,
    processed_items: int = 0,
    indexed_rows: int = 0,
    removed_rows: int = 0,
    started: bool = False,
    finished: bool = False,
    error: str | None = None,
) -> None:
    """Overwrite the single current reload status row."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO sync_status (
                    id,
                    status,
                    total_items,
                    processed_items,
                    indexed_rows,
                    removed_rows,
                    started_at,
                    finished_at,
                    error,
                    updated_at
                )
                VALUES (
                    1,
                    :status,
                    :total_items,
                    :processed_items,
                    :indexed_rows,
                    :removed_rows,
                    CASE WHEN :started THEN CURRENT_TIMESTAMP ELSE NULL END,
                    CASE WHEN :finished THEN CURRENT_TIMESTAMP ELSE NULL END,
                    :error,
                    CURRENT_TIMESTAMP
                )
                ON CONFLICT (id) DO UPDATE SET
                    status = EXCLUDED.status,
                    total_items = EXCLUDED.total_items,
                    processed_items = EXCLUDED.processed_items,
                    indexed_rows = EXCLUDED.indexed_rows,
                    removed_rows = EXCLUDED.removed_rows,
                    started_at = CASE
                        WHEN :started THEN CURRENT_TIMESTAMP
                        ELSE sync_status.started_at
                    END,
                    finished_at = CASE
                        WHEN :finished THEN CURRENT_TIMESTAMP
                        ELSE NULL
                    END,
                    error = EXCLUDED.error,
                    updated_at = CURRENT_TIMESTAMP
                """,
                {
                    "status": status,
                    "total_items": total_items,
                    "processed_items": processed_items,
                    "indexed_rows": indexed_rows,
                    "removed_rows": removed_rows,
                    "started": int(started),
                    "finished": int(finished),
                    "error": error,
                },
            )
        connection.commit()


def get_sync_status() -> SyncStatus:
    """Return the current reload status."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT
                    status,
                    total_items,
                    processed_items,
                    indexed_rows,
                    removed_rows,
                    started_at,
                    finished_at,
                    error
                FROM sync_status
                WHERE id = 1
                """
            )
            row = cursor.fetchone()

    if row is None:
        return SyncStatus(
            status="idle",
            total_items=0,
            processed_items=0,
            indexed_rows=0,
            removed_rows=0,
        )
    return SyncStatus(**dict(row))


def fetch_presentations(limit: int = 10, offset: int = 0) -> list[PresentationRecord]:
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT source_id, title, workbook_name, sheet_name, row_number, metadata
                FROM presentations
                ORDER BY sheet_name, row_number
                LIMIT ?
                OFFSET ?
                """,
                (limit, offset),
            )
            results = cursor.fetchall()

    records: list[PresentationRecord] = []
    for row in results:
        values = dict(row)
        values["metadata"] = _loads_metadata(row["metadata"])
        records.append(PresentationRecord(**values))
    return records


def fetch_all_presentation_metadata() -> list[dict]:
    """Return all current catalog rows as configured spreadsheet metadata."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT metadata
                FROM presentations
                ORDER BY sheet_name, row_number
                """
            )
            results = cursor.fetchall()

    return [_loads_metadata(row["metadata"]) for row in results]


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
                    embedding
                FROM presentations
                """,
            )
            results = cursor.fetchall()

    scored_results: list[SearchResult] = []
    for row in results:
        embedding = json.loads(row["embedding"])
        if not isinstance(embedding, list):
            continue
        score = _cosine_similarity(query_embedding, embedding)
        scored_results.append(
            SearchResult(
                source_id=row["source_id"],
                title=row["title"],
                workbook_name=row["workbook_name"],
                sheet_name=row["sheet_name"],
                row_number=row["row_number"],
                metadata=_loads_metadata(row["metadata"]),
                score=score,
            )
        )

    scored_results.sort(key=lambda result: result.score, reverse=True)
    return scored_results[:top_k]
