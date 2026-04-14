"""Configuration helpers for the vector search application."""

import os


def get_database_url() -> str:
    return os.getenv(
        "DATABASE_URL",
        "postgresql://postgres:postgres@localhost:5432/presentations",
    )


def get_embedding_model() -> str:
    return os.getenv("EMBEDDING_MODEL", "text-embedding-3-small")


def get_embedding_dimension() -> int:
    return int(os.getenv("EMBEDDING_DIMENSION", "1536"))


def get_auto_index_on_startup() -> bool:
    return os.getenv("AUTO_INDEX_ON_STARTUP", "false").lower() == "true"
