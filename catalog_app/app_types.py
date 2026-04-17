"""Shared project types.

The AGENTS instructions for this repository say that project-local shared types
should live here. If you need a new `TypedDict` or another reusable type, add it
to this file and import it from the rest of the project.
"""

from dataclasses import dataclass
from typing import Any, Callable, NotRequired, TypedDict

from catalog_app.generation.microsoft.types import GraphDriveItem, GraphHeaders


class ConfiguredDriveSource(TypedDict):
    """A drive/folder entry exactly as the user writes it in configuration."""

    name: str
    folder: NotRequired[str]


class PresentationColumn(TypedDict):
    """One Excel column definition.

    The generator receives the source PowerPoint item metadata and returns the
    value that should be written into that column.
    """

    name: str
    generator: Callable[[dict[str, Any]], Any]


class DriveSource(TypedDict):
    """A fully resolved drive source used at runtime.

    Unlike `ConfiguredDriveSource`, this structure includes the Graph drive ID
    and, when applicable, the resolved folder ID.
    """

    name: str
    drive_id: str
    is_default: bool
    folder: NotRequired[str]
    folder_id: NotRequired[str]


class ExcelSetup(TypedDict):
    """The authentication headers and resolved drive list needed for a run."""

    headers: GraphHeaders
    drive_sources: list[DriveSource]


class DeltaCollectionResult(TypedDict):
    """Changed Graph items plus the new delta link for a source."""

    items: list[GraphDriveItem]
    delta_link: str


@dataclass(frozen=True)
class IndexedWorkbookRow:
    """A generated catalog row prepared for embedding-backed indexing."""

    source_id: str
    title: str
    workbook_name: str
    sheet_name: str
    row_number: int
    metadata: dict[str, Any]
    searchable_text: str
    source_key: str | None = None
    drive_id: str | None = None
    item_id: str | None = None
    web_url: str | None = None
    last_modified_at: str | None = None
