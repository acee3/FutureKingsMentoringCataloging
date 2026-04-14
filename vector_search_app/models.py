"""Pydantic request and response models for the vector search API."""

from pydantic import BaseModel, Field


class SearchRequest(BaseModel):
    """Natural-language search request."""

    query: str = Field(min_length=1)
    top_k: int = Field(default=10, ge=1, le=100)


class PresentationRecord(BaseModel):
    """Workbook row returned for the spreadsheet UI."""

    source_id: str
    title: str
    workbook_name: str
    sheet_name: str
    row_number: int
    metadata: dict[str, object]


class SearchResult(PresentationRecord):
    """Workbook row plus similarity score."""

    score: float
