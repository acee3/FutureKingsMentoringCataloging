"""TypedDict definitions for the Microsoft Graph data the app uses."""

from typing import Any, NotRequired, TypedDict


type GraphHeaders = dict[str, str]


class GraphUser(TypedDict, total=False):
    """Subset of a Microsoft Graph user object."""

    email: str
    displayName: str


class GraphIdentity(TypedDict, total=False):
    """Identity wrapper used by several Graph fields."""

    user: GraphUser


class GraphParentReference(TypedDict, total=False):
    """Where an item lives inside Microsoft Graph."""

    driveType: str
    driveId: str
    id: str
    name: str
    path: str
    siteId: str


class GraphHashes(TypedDict, total=False):
    """File hash values provided by Graph when available."""

    quickXorHash: str


class GraphFileFacet(TypedDict, total=False):
    """File-specific metadata attached to drive items."""

    fileExtension: str
    hashes: GraphHashes
    mimeType: str


class GraphFileSystemInfo(TypedDict, total=False):
    """Filesystem timestamps for a drive item."""

    createdDateTime: str
    lastModifiedDateTime: str


class GraphShared(TypedDict, total=False):
    """Sharing metadata for an item."""

    scope: str


class GraphFolderFacet(TypedDict, total=False):
    """Folder-specific metadata attached to drive items."""

    childCount: int


class GraphDeletedFacet(TypedDict, total=False):
    """Deletion marker returned by Microsoft Graph delta queries."""

    state: str


class GraphDriveItem(TypedDict, total=False):
    """Subset of the Graph drive item schema used by this project."""

    id: str
    name: str
    size: int
    createdDateTime: str
    lastModifiedDateTime: str
    webUrl: str
    eTag: str
    cTag: str
    createdBy: GraphIdentity
    lastModifiedBy: GraphIdentity
    parentReference: GraphParentReference
    file: NotRequired[GraphFileFacet]
    fileSystemInfo: NotRequired[GraphFileSystemInfo]
    shared: NotRequired[GraphShared]
    folder: NotRequired[GraphFolderFacet]
    deleted: NotRequired[GraphDeletedFacet]
    configuredSourceFolder: NotRequired[str]
    configuredSourceName: NotRequired[str]
    odata_context: NotRequired[str]
    download_url: NotRequired[str]
    additional_data: NotRequired[dict[str, Any]]
