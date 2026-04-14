"""Vector-search-specific shared types."""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class IndexedWorkbookRow:
    """A workbook row prepared for vector indexing."""

    source_id: str
    title: str
    workbook_name: str
    sheet_name: str
    row_number: int
    metadata: dict[str, Any]
    searchable_text: str
