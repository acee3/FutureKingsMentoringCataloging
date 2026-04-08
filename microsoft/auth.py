import os

import msal

from app_types import DriveSource, ExcelSetup
from configuration import DRIVE_SOURCES

from .graph import get_drive_id, get_drive_item_by_path, get_site_id
from .types import GraphHeaders


def excel_setup() -> ExcelSetup:
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET_VALUE")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    token = app.acquire_token_for_client(scopes=scopes)
    access_token = token["access_token"]
    headers: GraphHeaders = {"Authorization": f"Bearer {access_token}"}
    site_id = get_site_id(site_hostname, site_path, headers)

    drive_sources: list[DriveSource] = []
    for raw_source in DRIVE_SOURCES:
        drive_id = get_drive_id(site_id, raw_source["name"], headers)
        source: DriveSource = {
            "name": raw_source["name"],
            "drive_id": drive_id,
        }
        folder = raw_source.get("folder")
        if folder:
            folder_item = get_drive_item_by_path(drive_id, folder, headers)
            source["folder"] = folder
            source["folder_id"] = folder_item["id"]
        drive_sources.append(source)

    return {
        "headers": headers,
        "drive_sources": drive_sources,
    }
