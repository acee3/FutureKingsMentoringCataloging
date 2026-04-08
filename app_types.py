from typing import Any, Callable, NotRequired, TypedDict

from microsoft.types import GraphDriveItem, GraphHeaders


class ConfiguredDriveSource(TypedDict):
    name: str
    folder: NotRequired[str]


class PresentationColumn(TypedDict):
    name: str
    generator: Callable[[dict[str, Any]], Any]


class DriveSource(TypedDict):
    name: str
    drive_id: str
    folder: NotRequired[str]
    folder_id: NotRequired[str]


class ExcelSetup(TypedDict):
    headers: GraphHeaders
    drive_sources: list[DriveSource]


class RunCheckpoint(TypedDict):
    processed_rows: list[dict[str, Any]]
    pending_items: list[GraphDriveItem]
