"""
tools/excel_update.py
Add / update rows in the OneDrive Excel client database via Microsoft Graph API.
The Excel file must contain a table named "ClientDatabase".

Expected columns (order matters for new row inserts):
    firstName | lastName | email | phone | filingStatus | incomeTypes | lifeChanges
    dependentsCount | submittedAt | status | statusUpdatedAt | underReviewAt
    filedAt | completedAt | oneDriveFolderUrl | notes | last4

Usage (CLI):
    # Update a field
    python excel_update.py update <first> <last> <last4> <field> <value>

    # Lookup a client
    python excel_update.py lookup <first> <last> <last4>

    # Add a new row  (provide values as JSON string)
    python excel_update.py add '{"firstName":"Jane","lastName":"Doe","last4":"1234",...}'

Usage (module):
    from tools.excel_update import update_client, lookup_client, add_client_row

Credentials (from .env):
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, EXCEL_FILE_ID,
    ONEDRIVE_USER_ID
"""

import os
import sys
import json
import requests
from datetime import datetime, timezone
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TABLE_NAME = "ClientDatabase"

# Canonical column order — must match the Excel table exactly
COLUMNS = [
    "firstName", "lastName", "email", "phone", "filingStatus", "incomeTypes",
    "lifeChanges", "dependentsCount", "submittedAt", "status", "statusUpdatedAt",
    "underReviewAt", "filedAt", "completedAt", "oneDriveFolderUrl", "notes", "last4",
]


def _get_token() -> str:
    tenant = os.environ["AZURE_TENANT_ID"]
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": os.environ["AZURE_CLIENT_ID"],
            "client_secret": os.environ["AZURE_CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
        },
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _file_base(token: str) -> str:
    user_id = os.environ.get("ONEDRIVE_USER_ID", os.environ.get("EMAIL_FROM", ""))
    file_id = os.environ["EXCEL_FILE_ID"]
    return f"{GRAPH_BASE}/users/{requests.utils.quote(user_id, safe='')}/drive/items/{file_id}/workbook/tables/{TABLE_NAME}"


def _get_all_rows(token: str) -> list[dict]:
    """Return all table rows as a list of dicts keyed by column name."""
    base = _file_base(token)
    resp = requests.get(f"{base}/rows?$select=index,values", headers=_headers(token))
    resp.raise_for_status()
    rows = []
    for row in resp.json().get("value", []):
        values = row["values"][0]  # Each row is [[v1, v2, ...]]
        rows.append({"_index": row["index"], **dict(zip(COLUMNS, values))})
    return rows


def _match_client(rows: list[dict], first: str, last: str, last4: str) -> dict | None:
    """Find the first row matching first+last+last4 (case-insensitive)."""
    for row in rows:
        if (
            str(row.get("firstName", "")).strip().lower() == first.strip().lower()
            and str(row.get("lastName", "")).strip().lower() == last.strip().lower()
            and str(row.get("last4", "")).strip() == last4.strip()
        ):
            return row
    return None


def lookup_client(first: str, last: str, last4: str) -> dict | None:
    """
    Look up a client by first name, last name, and last 4 SSN digits.

    Returns the row dict (including '_index') or None if not found.
    """
    token = _get_token()
    rows = _get_all_rows(token)
    return _match_client(rows, first, last, last4)


def update_client(first: str, last: str, last4: str, field: str, value: str) -> dict:
    """
    Update a single field for a client identified by first+last+last4.

    Args:
        first, last, last4: Client identifier.
        field: Column name to update (must be in COLUMNS).
        value: New value for the field.

    Returns:
        The updated row dict.

    Raises:
        KeyError if client not found.
        ValueError if field is not a valid column.
    """
    if field not in COLUMNS:
        raise ValueError(f"Invalid field '{field}'. Valid columns: {COLUMNS}")

    token = _get_token()
    rows = _get_all_rows(token)
    row = _match_client(rows, first, last, last4)
    if row is None:
        raise KeyError(f"Client not found: {first} {last} last4={last4}")

    # Build updated values array in column order
    updated = {**row, field: value}
    if field == "status":
        updated["statusUpdatedAt"] = datetime.now(timezone.utc).isoformat()

    values = [[updated.get(col, "") for col in COLUMNS]]

    base = _file_base(token)
    idx = row["_index"]
    resp = requests.patch(
        f"{base}/rows/itemAt(index={idx})",
        headers=_headers(token),
        json={"values": values},
    )
    resp.raise_for_status()
    return {**updated, "_index": idx}


def add_client_row(data: dict) -> dict:
    """
    Append a new client row to the Excel table.

    Args:
        data: Dict with any subset of COLUMNS keys.

    Returns:
        Graph API response dict.
    """
    token = _get_token()
    if not data.get("submittedAt"):
        data["submittedAt"] = datetime.now(timezone.utc).isoformat()
    if not data.get("status"):
        data["status"] = "Received"
    if not data.get("statusUpdatedAt"):
        data["statusUpdatedAt"] = data["submittedAt"]

    values = [[data.get(col, "") for col in COLUMNS]]
    base = _file_base(token)
    resp = requests.post(
        f"{base}/rows/add",
        headers=_headers(token),
        json={"values": values},
    )
    resp.raise_for_status()
    return resp.json()


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = sys.argv[1:]

    if not args:
        # Self-test: lookup a dummy client
        print("Running self-test lookup...")
        result = lookup_client("Test", "Client", "0000")
        if result:
            print("Found:", result)
        else:
            print("Client not found (expected for test data).")

    elif args[0] == "lookup" and len(args) == 4:
        row = lookup_client(args[1], args[2], args[3])
        if row:
            print(json.dumps({k: v for k, v in row.items() if k != "_index"}, indent=2))
        else:
            print("Client not found.")

    elif args[0] == "update" and len(args) == 6:
        _, first, last, last4, field, value = args
        updated = update_client(first, last, last4, field, value)
        print(f"Updated '{field}' for {first} {last}:")
        print(json.dumps({k: v for k, v in updated.items() if k != "_index"}, indent=2))

    elif args[0] == "add" and len(args) == 2:
        data = json.loads(args[1])
        result = add_client_row(data)
        print("Row added:", result)

    else:
        print(__doc__)
