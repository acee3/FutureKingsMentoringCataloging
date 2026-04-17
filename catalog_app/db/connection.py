"""Shared database connection helpers."""

from pathlib import Path
import sqlite3

from catalog_app.settings import get_database_path


class CatalogCursor(sqlite3.Cursor):
    def __enter__(self) -> "CatalogCursor":
        return self

    def __exit__(self, *args: object) -> None:
        self.close()


class CatalogConnection(sqlite3.Connection):
    def cursor(self, *args, **kwargs):  # type: ignore[no-untyped-def]
        kwargs.setdefault("factory", CatalogCursor)
        return super().cursor(*args, **kwargs)


def get_connection() -> sqlite3.Connection:
    database_path = get_database_path()
    if database_path != ":memory:":
        Path(database_path).parent.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(database_path, timeout=30, factory=CatalogConnection)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA journal_mode = WAL")
    return connection
