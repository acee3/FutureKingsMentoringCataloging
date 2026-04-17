"""Configuration helpers for the catalog application."""

import os


def get_database_path() -> str:
    return os.getenv("DATABASE_PATH", "catalog.sqlite3")


def get_embedding_model() -> str:
    return os.getenv("EMBEDDING_MODEL", "text-embedding-3-small")
