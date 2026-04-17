"""Database schema setup for the current-state presentation catalog."""

from catalog_app.db.connection import get_connection


def ensure_schema() -> None:
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS presentation_sources (
                    source_key TEXT PRIMARY KEY,
                    source_name TEXT NOT NULL,
                    drive_id TEXT NOT NULL,
                    folder_id TEXT,
                    folder_path TEXT,
                    delta_link TEXT,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS presentations (
                    source_id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    workbook_name TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    row_number INTEGER NOT NULL,
                    source_key TEXT REFERENCES presentation_sources(source_key)
                        ON UPDATE CASCADE
                        ON DELETE SET NULL,
                    drive_id TEXT,
                    item_id TEXT,
                    web_url TEXT,
                    last_modified_at TEXT,
                    metadata TEXT NOT NULL,
                    searchable_text TEXT NOT NULL,
                    embedding TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """
            )

            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS presentations_sheet_row_idx
                ON presentations (sheet_name, row_number);
                """
            )

            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS presentations_item_idx
                ON presentations (drive_id, item_id);
                """
            )

            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS presentations_source_key_idx
                ON presentations (source_key);
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS sync_status (
                    id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
                    status TEXT NOT NULL,
                    total_items INTEGER NOT NULL DEFAULT 0,
                    processed_items INTEGER NOT NULL DEFAULT 0,
                    indexed_rows INTEGER NOT NULL DEFAULT 0,
                    removed_rows INTEGER NOT NULL DEFAULT 0,
                    started_at TEXT,
                    finished_at TEXT,
                    error TEXT,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """
            )

            cursor.execute(
                """
                INSERT INTO sync_status (
                    id,
                    status,
                    total_items,
                    processed_items,
                    indexed_rows,
                    removed_rows
                )
                VALUES (1, 'idle', 0, 0, 0, 0)
                ON CONFLICT (id) DO NOTHING;
                """
            )
        connection.commit()
