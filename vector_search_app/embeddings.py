"""Embedding helpers for workbook rows and search queries."""

from collections.abc import Sequence

from llm_work import get_openai_client
from openai import OpenAI

from vector_search_app.settings import get_embedding_model


def _get_client() -> OpenAI:
    return get_openai_client()


def embed_texts(texts: Sequence[str]) -> list[list[float]]:
    if not texts:
        return []

    response = _get_client().embeddings.create(
        model=get_embedding_model(),
        input=list(texts),
    )
    return [item.embedding for item in response.data]


def embed_text(text: str) -> list[float]:
    embeddings = embed_texts([text])
    if not embeddings:
        raise ValueError("Expected one embedding result")
    return embeddings[0]
