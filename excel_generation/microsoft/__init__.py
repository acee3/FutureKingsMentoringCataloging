"""Convenience exports for Microsoft-related helpers and types."""

from .graph import (
    collect_drive_delta,
    download_pptx_file_content,
    get_all_pptx_files,
    get_drive_delta_page,
    get_drive_id,
    get_drive_item_by_path,
    get_pptx_file,
    get_site_id,
)
from .types import (
    GraphDriveItem,
    GraphDeletedFacet,
    GraphFileFacet,
    GraphFileSystemInfo,
    GraphFolderFacet,
    GraphHashes,
    GraphHeaders,
    GraphIdentity,
    GraphParentReference,
    GraphShared,
    GraphUser,
)


def excel_setup():
    """Import auth code lazily so type-only imports stay lightweight."""
    # Import lazily so graph/type consumers do not require auth dependencies
    # like msal just to import the microsoft package.
    from .auth import excel_setup as _excel_setup

    return _excel_setup()


__all__ = [
    "download_pptx_file_content",
    "excel_setup",
    "collect_drive_delta",
    "get_all_pptx_files",
    "get_drive_delta_page",
    "get_drive_id",
    "get_drive_item_by_path",
    "get_pptx_file",
    "get_site_id",
    "GraphDriveItem",
    "GraphDeletedFacet",
    "GraphFileFacet",
    "GraphFileSystemInfo",
    "GraphFolderFacet",
    "GraphHashes",
    "GraphHeaders",
    "GraphIdentity",
    "GraphParentReference",
    "GraphShared",
    "GraphUser",
]
