import os

import msal
import requests


def excel_setup():
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET_VALUE")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]
    site_hostname = os.getenv("SITE_HOSTNAME")
    site_path = os.getenv("SITE_PATH")
    drive_name = os.getenv("DRIVE_NAME")

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    token = app.acquire_token_for_client(scopes=scopes)
    access_token = token["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}
    site_id = get_site_id(site_hostname, site_path, headers)
    library_drive_id = get_drive_id(site_id, drive_name, headers)

    return headers, library_drive_id


def get_site_id(site_hostname, site_path, headers):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    site_data = resp.json()
    return site_data["id"]


def get_drive_id(site_id, drive_name, headers):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    drives_data = resp.json()
    for drive in drives_data["value"]:
        if drive["name"] == drive_name:
            return drive["id"]
    raise ValueError("Drive not found")


def get_all_pptx_files(drive_id, headers, item_id="") -> list[str]:
    item_path = f"items/{item_id}" if item_id else "root"
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{item_path}/children"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    items = resp.json()["value"]

    subfolders = [x for x in items if "folder" in x]
    subfolder_pptx_files = [
        get_all_pptx_files(drive_id, headers, x["id"]) for x in subfolders
    ]
    pptx_files = [x for x in items if x["name"].lower().endswith(".pptx")]
    return [f for f in pptx_files] + [
        f for subfolder_files in subfolder_pptx_files for f in subfolder_files
    ]


def get_file(drive_id, item_id, headers):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()


def download_pptx_file_content(drive_id, item_id, headers):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.content
