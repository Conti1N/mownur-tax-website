"""
tools/onedrive_upload.py
Upload a local file to a specific path in OneDrive using Microsoft Graph API.

Usage (CLI):
    python onedrive_upload.py <local_file_path> <onedrive_destination_path>

    Example:
        python onedrive_upload.py /tmp/w2.pdf "Mownur Clients/Jane Doe - 1234/w2.pdf"

Usage (module):
    from tools.onedrive_upload import upload_file
    result = upload_file("/tmp/w2.pdf", "Mownur Clients/Jane Doe - 1234/w2.pdf")

Credentials (from .env):
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, ONEDRIVE_USER_ID
"""

import os
import sys
import mimetypes
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
CHUNK_SIZE = 5 * 1024 * 1024  # 5 MB


def _get_token() -> str:
    tenant = os.environ["AZURE_TENANT_ID"]
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type": "client_credentials",
        "client_id": os.environ["AZURE_CLIENT_ID"],
        "client_secret": os.environ["AZURE_CLIENT_SECRET"],
        "scope": "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


def upload_file(local_file_path: str, onedrive_destination_path: str) -> dict:
    """
    Upload a local file to OneDrive at the given destination path.

    Args:
        local_file_path: Absolute or relative path to the file on disk.
        onedrive_destination_path: Path inside OneDrive, e.g.
            "Mownur Clients/Jane Doe - 1234/w2.pdf"

    Returns:
        Graph API response dict with 'id', 'name', 'webUrl', etc.
    """
    path = Path(local_file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {local_file_path}")

    user_id = os.environ.get("ONEDRIVE_USER_ID", os.environ.get("EMAIL_FROM", ""))
    if not user_id:
        raise ValueError("ONEDRIVE_USER_ID or EMAIL_FROM must be set in .env")

    token = _get_token()
    headers = {"Authorization": f"Bearer {token}"}
    file_size = path.stat().st_size
    content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"

    # Encode each path segment separately to preserve slashes
    encoded_dest = "/".join(
        requests.utils.quote(seg, safe="") for seg in onedrive_destination_path.split("/")
    )
    item_path = f"/users/{requests.utils.quote(user_id, safe='')}/drive/root:/{encoded_dest}"

    if file_size <= 4 * 1024 * 1024:
        # Simple PUT upload (≤ 4 MB)
        upload_url = f"{GRAPH_BASE}{item_path}:/content"
        with open(path, "rb") as f:
            resp = requests.put(
                upload_url,
                headers={**headers, "Content-Type": content_type},
                data=f,
            )
        resp.raise_for_status()
        return resp.json()
    else:
        # Large file: create upload session then send in chunks
        session_url = f"{GRAPH_BASE}{item_path}:/createUploadSession"
        session_resp = requests.post(
            session_url,
            headers={**headers, "Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": "rename"}},
        )
        session_resp.raise_for_status()
        upload_url = session_resp.json()["uploadUrl"]

        result = None
        with open(path, "rb") as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(CHUNK_SIZE)
                end = offset + len(chunk) - 1
                chunk_headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {offset}-{end}/{file_size}",
                    "Content-Type": "application/octet-stream",
                }
                r = requests.put(upload_url, headers=chunk_headers, data=chunk)
                r.raise_for_status()
                result = r.json() if r.content else {}
                offset += len(chunk)

        return result


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) == 3:
        local, dest = sys.argv[1], sys.argv[2]
        print(f"Uploading '{local}' → OneDrive:'{dest}' ...")
        item = upload_file(local, dest)
        print(f"Success! OneDrive item ID : {item.get('id')}")
        print(f"Web URL                   : {item.get('webUrl')}")
    else:
        # ── Quick self-test (creates a tiny test file) ─────────────────────
        print("Running self-test upload...")
        test_file = Path("/tmp/_od_test.txt")
        test_file.write_text("Mownur Services — OneDrive upload test")
        result = upload_file(str(test_file), "Mownur Clients/_test/_od_test.txt")
        print("Test upload OK:", result.get("webUrl", result))
        test_file.unlink()
