"""
tools/send_email.py
Send email via Microsoft Graph API (Outlook / M365).

Usage (CLI):
    python send_email.py <to_email> <subject> <body_html>

    Example:
        python send_email.py "jane@example.com" "Your return is ready" "<p>Hi Jane...</p>"

Usage (module):
    from tools.send_email import send_email

Credentials (from .env):
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID,
    EMAIL_FROM, ONEDRIVE_USER_ID (falls back to EMAIL_FROM)
"""

import os
import sys
import requests
from dotenv import load_dotenv

load_dotenv()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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


def send_email(
    to_email: str,
    subject: str,
    body_html: str,
    cc_emails: list[str] | None = None,
    save_to_sent: bool = True,
) -> None:
    """
    Send an HTML email from the configured EMAIL_FROM mailbox.

    Args:
        to_email:     Recipient email address.
        subject:      Email subject line.
        body_html:    Full HTML body of the email.
        cc_emails:    Optional list of CC addresses.
        save_to_sent: Whether to save to Sent Items (default True).

    Raises:
        requests.HTTPError on Graph API failure.
    """
    from_addr = os.environ.get("EMAIL_FROM", "")
    if not from_addr:
        raise ValueError("EMAIL_FROM not set in .env")

    user_id = os.environ.get("ONEDRIVE_USER_ID", from_addr)
    token = _get_token()

    message: dict = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body_html},
        "toRecipients": [{"emailAddress": {"address": to_email}}],
    }

    if cc_emails:
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_emails
        ]

    payload = {"message": message, "saveToSentItems": save_to_sent}

    url = f"{GRAPH_BASE}/users/{requests.utils.quote(user_id, safe='')}/sendMail"
    resp = requests.post(
        url,
        json=payload,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )
    resp.raise_for_status()


def send_status_update_email(
    to_email: str,
    client_first: str,
    new_status: str,
    notes: str = "",
) -> None:
    """Convenience wrapper: send a pre-formatted status update email."""
    status_descriptions = {
        "Received":    "We have received your tax documents and your return is in our queue.",
        "Under Review": "Our team is currently reviewing your documents.",
        "Filed":       "Your tax return has been filed with the IRS.",
        "Completed":   "Your tax return is complete. Please review and confirm receipt.",
    }
    description = status_descriptions.get(new_status, f"Your status has been updated to: {new_status}.")
    notes_block = f"<p><strong>Notes from your preparer:</strong> {notes}</p>" if notes else ""

    body = f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #1a1a2e;">Tax Return Status Update</h2>
      <p>Hi {client_first},</p>
      <p>{description}</p>
      {notes_block}
      <p>If you have any questions, reply to this email or call us.</p>
      <p style="color: #666; font-size: 0.85em;">
        Mownur Services — Minneapolis Tax Professionals<br>
        Most returns completed within 5 business days.
      </p>
    </div>
    """
    send_email(
        to_email=to_email,
        subject=f"Mownur Services: Your Tax Return — {new_status}",
        body_html=body,
    )


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) == 4:
        _, to, subject, body = sys.argv
        send_email(to, subject, body)
        print(f"Email sent to {to}")
    else:
        # Self-test
        test_to = os.environ.get("EMAIL_FROM", "")
        if not test_to:
            print("Set EMAIL_FROM in .env to run self-test.")
            sys.exit(1)
        print(f"Sending test email to {test_to} ...")
        send_email(
            to_email=test_to,
            subject="[TEST] Mownur Services — send_email.py self-test",
            body_html="<p>This is an automated self-test from <code>tools/send_email.py</code>.</p>",
        )
        print("Done. Check the inbox.")
