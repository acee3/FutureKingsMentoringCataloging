import logging
import time

import requests
from requests import Response
from requests.exceptions import HTTPError, RequestException

from .types import GraphDriveItem, GraphHeaders

logger = logging.getLogger(__name__)

DEFAULT_REQUEST_TIMEOUT = (10, 15)
DOWNLOAD_RETRY_ATTEMPTS = 4
DOWNLOAD_RETRY_BACKOFF_SECONDS = 2


def _should_retry(response: Response | None, error: Exception | None) -> bool:
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
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    site_data = resp.json()
    return site_data["id"]


def get_drive_id(site_id: str, drive_name: str, headers: GraphHeaders) -> str:
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


def get_all_pptx_files(
    drive_id: str, headers: GraphHeaders, item_id: str = ""
) -> list[GraphDriveItem]:
    item_path = f"items/{item_id}" if item_id else "root"
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{item_path}/children"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    items: list[GraphDriveItem] = resp.json()["value"]

    subfolders = [x for x in items if "folder" in x]
    subfolder_pptx_files = [
        get_all_pptx_files(drive_id, headers, x["id"]) for x in subfolders
    ]
    pptx_files = [x for x in items if x["name"].lower().endswith(".pptx")]
    return [f for f in pptx_files] + [
        f for subfolder_files in subfolder_pptx_files for f in subfolder_files
    ]


def get_pptx_file(
    drive_id: str, item_id: str, headers: GraphHeaders
) -> GraphDriveItem:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    resp = requests.get(url, headers=headers, timeout=DEFAULT_REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def download_pptx_file_content(
    drive_id: str, item_id: str, headers: GraphHeaders
) -> bytes:
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
