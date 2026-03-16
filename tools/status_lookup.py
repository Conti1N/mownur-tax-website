"""
tools/status_lookup.py
Query the Excel client database for a client's full status timeline.

Usage (CLI):
    python status_lookup.py <first> <last> <last4>

    Example:
        python status_lookup.py Jane Doe 1234

Usage (module):
    from tools.status_lookup import get_client_status

Credentials (from .env):
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID,
    EXCEL_FILE_ID, ONEDRIVE_USER_ID
"""

import os
import sys
import json
from dotenv import load_dotenv

load_dotenv()

# Import the lookup function from the sibling module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_update import lookup_client  # noqa: E402

# Status stages in chronological order
STATUS_STAGES = [
    ("submittedAt",    "Received"),
    ("underReviewAt",  "Under Review"),
    ("filedAt",        "Filed"),
    ("completedAt",    "Completed"),
]


def get_client_status(first: str, last: str, last4: str) -> dict:
    """
    Look up a client and return their full status timeline.

    Returns a dict with:
        found (bool)
        client (dict | None)  — all fields except _index
        current_status (str)
        timeline (list of {stage, timestamp} in chronological order)

    Raises nothing — returns found=False on miss.
    """
    row = lookup_client(first, last, last4)
    if row is None:
        return {
            "found": False,
            "client": None,
            "current_status": None,
            "timeline": [],
        }

    timeline = []
    for field, label in STATUS_STAGES:
        ts = row.get(field, "")
        if ts:
            timeline.append({"stage": label, "timestamp": ts})

    # Current status is the explicit status field; fallback to last completed stage
    current = row.get("status") or (timeline[-1]["stage"] if timeline else "Unknown")

    return {
        "found": True,
        "client": {k: v for k, v in row.items() if k != "_index"},
        "current_status": current,
        "timeline": timeline,
    }


def format_status_report(result: dict) -> str:
    """Return a human-readable status report string."""
    if not result["found"]:
        return "Client not found."

    c = result["client"]
    name = f"{c.get('firstName', '')} {c.get('lastName', '')}".strip()
    lines = [
        f"Client: {name}",
        f"Email : {c.get('email', '—')}",
        f"Phone : {c.get('phone', '—')}",
        f"Status: {result['current_status']}",
        "",
        "Timeline:",
    ]
    for entry in result["timeline"]:
        lines.append(f"  ✓ {entry['stage']:15s}  {entry['timestamp']}")

    if not result["timeline"]:
        lines.append("  (no timestamps recorded yet)")

    if c.get("notes"):
        lines += ["", f"Notes : {c['notes']}"]
    if c.get("oneDriveFolderUrl"):
        lines += ["", f"Folder: {c['oneDriveFolderUrl']}"]

    return "\n".join(lines)


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) == 4:
        _, first, last, last4 = sys.argv
        result = get_client_status(first, last, last4)
        print(format_status_report(result))
        print()
        print("Raw JSON:")
        print(json.dumps(result, indent=2, default=str))
    else:
        # Self-test
        print("Running self-test lookup (expecting not-found)...")
        result = get_client_status("Test", "Client", "0000")
        print(format_status_report(result))
        print("Self-test complete.")
