"""Shared project types.

The AGENTS instructions for this repository say that project-local shared types
should live here. If you need a new `TypedDict` or another reusable type, add it
to this file and import it from the rest of the project.
"""

from typing import Any, Callable, NotRequired, TypedDict

from microsoft.types import GraphDriveItem, GraphHeaders


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


class RunCheckpoint(TypedDict):
    """Saved progress for resuming a partially completed export."""

    processed_rows: list[dict[str, Any]]
    pending_items: list[GraphDriveItem]
    delta_links: NotRequired[dict[str, str]]


class DeltaSourceState(TypedDict):
    """Saved Microsoft Graph delta state for one configured drive source."""

    source_name: str
    drive_id: str
    delta_link: str
    source_folder: NotRequired[str]
    folder_id: NotRequired[str]


class DeltaState(TypedDict):
    """Saved Microsoft Graph delta state for all configured drive sources."""

    sources: dict[str, DeltaSourceState]


class DeltaCollectionResult(TypedDict):
    """Changed Graph items plus the new delta link for a source."""

    items: list[GraphDriveItem]
    delta_link: str
