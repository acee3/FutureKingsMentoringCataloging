"""Thin wrappers around the Microsoft Graph HTTP API."""

from __future__ import annotations

import logging
import time
from typing import TYPE_CHECKING, Any

import requests
from requests import Response
from requests.exceptions import HTTPError, RequestException

from .types import GraphDriveItem, GraphHeaders

if TYPE_CHECKING:
    from app_types import DeltaCollectionResult, DriveSource

logger = logging.getLogger(__name__)

DEFAULT_REQUEST_TIMEOUT = (10, 15)
DOWNLOAD_RETRY_ATTEMPTS = 4
DOWNLOAD_RETRY_BACKOFF_SECONDS = 2
DELTA_SELECT_FIELDS = (
    "id,name,webUrl,lastModifiedDateTime,eTag,cTag,size,file,folder,"
    "parentReference,deleted"
)


def _should_retry(response: Response | None, error: Exception | None) -> bool:
    """Decide whether a failed download attempt is worth retrying."""
    if isinstance(error, HTTPError):
        error_response = error.response or response
        if error_response is None:
            return False
        return error_response.status_code == 429 or 500 <= error_response.status_code < 600
    if error is not None:
        return True
    if response is None:
        return False
    return response.status_code == 429 or 500 <= response.status_code < 600


def get_site_id(site_hostname: str, site_path: str, headers: GraphHeaders) -> str:
    """Resolve a SharePoint site into its Graph site ID."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    site_data = resp.json()
    return site_data["id"]


def get_drive_id(site_id: str, drive_name: str, headers: GraphHeaders) -> str:
    """Find the Graph drive ID for a human-readable drive name."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    drives_data = resp.json()
    for drive in drives_data["value"]:
        if drive["name"] == drive_name:
            return drive["id"]
    raise ValueError("Drive not found")


def get_drive_item_by_path(
    drive_id: str, item_path: str, headers: GraphHeaders
) -> GraphDriveItem:
    """Look up a file or folder by its path inside a drive."""
    normalized_item_path = item_path.strip("/")
    if not normalized_item_path:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root"
    else:
        url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:"
            f"/{normalized_item_path}"
        )
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def _get_delta_url_for_source(source: "DriveSource") -> str:
    """Build the initial Microsoft Graph delta URL for a configured source."""
    drive_id = source["drive_id"]
    folder_id = source.get("folder_id")
    item_path = f"items/{folder_id}" if folder_id else "root"
    return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{item_path}/delta"


def get_drive_delta_page(
    url_or_source: str | "DriveSource",
    headers: GraphHeaders,
    *,
    token: str | None = None,
) -> dict[str, Any]:
    """Fetch one Microsoft Graph delta page.

    A saved `@odata.nextLink` or `@odata.deltaLink` already contains all query
    parameters, so extra `$select` parameters are only used for initial source
    URLs.
    """
    params: dict[str, str] | None = None
    if isinstance(url_or_source, str):
        url = url_or_source
    else:
        url = _get_delta_url_for_source(url_or_source)
        params = {"$select": DELTA_SELECT_FIELDS}
        if token:
            params["token"] = token

    resp = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=DEFAULT_REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    return resp.json()


def collect_drive_delta(
    source: "DriveSource",
    headers: GraphHeaders,
    previous_delta_link: str | None = None,
    *,
    token: str | None = None,
) -> "DeltaCollectionResult":
    """Collect all changed items for one source and return its next delta link."""
    items: list[GraphDriveItem] = []
    next_request: str | DriveSource = previous_delta_link or source

    while True:
        page = get_drive_delta_page(
            next_request,
            headers,
            token=token if next_request is source else None,
        )
        items.extend(page.get("value", []))

        next_link = page.get("@odata.nextLink")
        if next_link:
            next_request = next_link
            token = None
            continue

        delta_link = page.get("@odata.deltaLink")
        if not delta_link:
            raise ValueError("Microsoft Graph delta response did not include a delta link")
        return {
            "items": items,
            "delta_link": delta_link,
        }


def get_all_pptx_files(
    drive_id: str,
    headers: GraphHeaders,
    item_id: str = "",
    configured_source_name: str = "",
    configured_source_folder: str = "",
) -> list[GraphDriveItem]:
    """Recursively collect every `.pptx` file below a drive or folder.

    Args:
        drive_id: Microsoft Graph drive ID.
        headers: Auth headers for Graph requests.
        item_id: Optional folder ID. If omitted, scanning starts at the drive
            root.
        configured_source_name: Human-readable drive name from configuration.
        configured_source_folder: Human-readable folder path from configuration.
    """
    item_path = f"items/{item_id}" if item_id else "root"
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{item_path}/children"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    items: list[GraphDriveItem] = resp.json()["value"]

    subfolders = [x for x in items if "folder" in x]
    subfolder_pptx_files = [
        get_all_pptx_files(
            drive_id,
            headers,
            x["id"],
            configured_source_name,
            configured_source_folder,
        )
        for x in subfolders
    ]
    pptx_files = [x for x in items if x["name"].lower().endswith(".pptx")]
    for pptx_file in pptx_files:
        if configured_source_name:
            pptx_file["configuredSourceName"] = configured_source_name
        if configured_source_folder:
            pptx_file["configuredSourceFolder"] = configured_source_folder
    return [f for f in pptx_files] + [
        f for subfolder_files in subfolder_pptx_files for f in subfolder_files
    ]


def get_pptx_file(
    drive_id: str,
    item_id: str,
    headers: GraphHeaders,
    configured_source_name: str = "",
    configured_source_folder: str = "",
) -> GraphDriveItem:
    """Fetch metadata for one PowerPoint item by ID."""
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    item: GraphDriveItem = resp.json()
    if configured_source_name:
        item["configuredSourceName"] = configured_source_name
    if configured_source_folder:
        item["configuredSourceFolder"] = configured_source_folder
    return item


def download_pptx_file_content(
    drive_id: str, item_id: str, headers: GraphHeaders
) -> bytes:
    """Download raw bytes for one PowerPoint file.

    Downloads can fail transiently because of throttling or temporary network
    issues, so this function retries a few times before giving up.
    """
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    last_error: Exception | None = None

    for attempt in range(1, DOWNLOAD_RETRY_ATTEMPTS + 1):
        response: Response | None = None
        try:
            response = requests.get(
                url,
                headers=headers,
                timeout=DEFAULT_REQUEST_TIMEOUT,
                stream=True,
            )
            response.raise_for_status()

            content = bytearray()
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    content.extend(chunk)
            return bytes(content)
        except RequestException as exc:
            last_error = exc
            if not _should_retry(response, exc) or attempt == DOWNLOAD_RETRY_ATTEMPTS:
                raise
            sleep_seconds = DOWNLOAD_RETRY_BACKOFF_SECONDS * attempt
            logger.warning(
                (
                    "Retrying PPTX download for item %s on drive %s "
                    "(attempt %s/%s) after %s: %s"
                ),
                item_id,
                drive_id,
                attempt,
                DOWNLOAD_RETRY_ATTEMPTS,
                sleep_seconds,
                exc,
            )
            time.sleep(sleep_seconds)
        finally:
            if response is not None:
                response.close()

    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Failed to download PPTX content for item {item_id}")
