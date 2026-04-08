"""Convenience exports for Microsoft-related helpers and types."""

from .graph import (
    download_pptx_file_content,
    get_all_pptx_files,
    get_drive_id,
    get_drive_item_by_path,
    get_pptx_file,
    get_site_id,
)
from .types import (
    GraphDriveItem,
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
    "get_all_pptx_files",
    "get_drive_id",
    "get_drive_item_by_path",
    "get_pptx_file",
    "get_site_id",
    "GraphDriveItem",
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
