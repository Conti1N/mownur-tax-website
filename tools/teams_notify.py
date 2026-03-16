"""
tools/teams_notify.py
Send a formatted Teams message via incoming webhook.

Usage (CLI):
    python teams_notify.py                        # sends a test message
    python teams_notify.py '{"client_name":"Jane Doe","email":"jane@example.com",...}'

Usage (module):
    from tools.teams_notify import notify_new_submission, notify_additional_docs

Credentials (from .env):
    TEAMS_WEBHOOK_URL
"""

import os
import sys
import json
from datetime import datetime, timezone
import pytz
import requests
from dotenv import load_dotenv

load_dotenv()

CT = pytz.timezone("America/Chicago")


def _ct_now() -> str:
    return datetime.now(CT).strftime("%B %d, %Y %I:%M %p CT")


def _send(payload: dict) -> None:
    url = os.environ.get("TEAMS_WEBHOOK_URL", "")
    if not url:
        raise ValueError("TEAMS_WEBHOOK_URL not set in .env")
    resp = requests.post(url, json=payload, timeout=15)
    resp.raise_for_status()


def notify_new_submission(
    client_name: str,
    email: str = "",
    phone: str = "",
    filing_status: str = "",
    income_types: str = "",
    life_changes: str = "",
    dependents_count: str = "0",
    doc_count: int = 0,
    onedrive_link: str = "",
    submitted_at: str = "",
) -> None:
    """
    Send a Teams notification for a new client intake submission.

    Args:
        client_name:     Full name of the client.
        email:           Client email address.
        phone:           Client phone number.
        filing_status:   e.g. "Single", "Married Filing Jointly"
        income_types:    Comma-separated list of income types.
        life_changes:    Notable life changes this tax year.
        dependents_count: Number of dependents (as string).
        doc_count:       Number of documents uploaded.
        onedrive_link:   URL to the client's OneDrive folder.
        submitted_at:    ISO timestamp; defaults to now (CT).
    """
    ts = submitted_at or _ct_now()
    lines = [
        "📋 **New Tax Client Submission — Mownur Services**",
        "",
        f"**Client:** {client_name}",
        f"**Email:** {email or '—'} | **Phone:** {phone or '—'}",
        f"**Filing Status:** {filing_status or '—'}",
        f"**Income Types:** {income_types or '—'}",
    ]
    if life_changes:
        lines.append(f"**Life Changes:** {life_changes}")
    if dependents_count and dependents_count != "0":
        lines.append(f"**Dependents:** {dependents_count}")
    lines += [
        f"**Documents Uploaded:** {doc_count} file{'s' if doc_count != 1 else ''}",
        "",
    ]
    if onedrive_link:
        lines.append(f"📁 [OneDrive Folder]({onedrive_link})")
    lines.append(f"**Submitted:** {ts}")

    _send({"text": "\n".join(lines)})


def notify_additional_docs(
    client_name: str,
    uploaded_count: int = 0,
    onedrive_link: str = "",
) -> None:
    """
    Send a Teams notification when a client uploads additional documents.
    """
    lines = [
        "📎 **Additional Documents Received — Mownur Services**",
        "",
        f"**Client:** {client_name}",
        f"**Files Uploaded:** {uploaded_count}",
        "",
    ]
    if onedrive_link:
        lines.append(f"📁 [OneDrive Folder]({onedrive_link})")
    lines.append(f"**Received:** {_ct_now()}")
    _send({"text": "\n".join(lines)})


def notify_status_update(client_name: str, new_status: str, notes: str = "") -> None:
    """Send a Teams notification when a client's status is updated."""
    lines = [
        "🔄 **Client Status Updated — Mownur Services**",
        "",
        f"**Client:** {client_name}",
        f"**New Status:** {new_status}",
    ]
    if notes:
        lines.append(f"**Notes:** {notes}")
    lines.append(f"**Updated:** {_ct_now()}")
    _send({"text": "\n".join(lines)})


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) == 2:
        opts = json.loads(sys.argv[1])
        notify_new_submission(
            client_name=opts.get("client_name", ""),
            email=opts.get("email", ""),
            phone=opts.get("phone", ""),
            filing_status=opts.get("filing_status", ""),
            income_types=opts.get("income_types", ""),
            life_changes=opts.get("life_changes", ""),
            dependents_count=str(opts.get("dependents_count", "0")),
            doc_count=int(opts.get("doc_count", 0)),
            onedrive_link=opts.get("onedrive_link", ""),
        )
        print("Teams notification sent.")
    else:
        # Self-test
        print("Sending test Teams notification...")
        notify_new_submission(
            client_name="Test Client",
            email="test@example.com",
            phone="612-555-0100",
            filing_status="Single",
            income_types="W-2, Freelance",
            life_changes="Bought a home",
            dependents_count="1",
            doc_count=3,
            onedrive_link="https://onedrive.live.com/test",
        )
        print("Done.")
