#!/usr/bin/env python3
"""Minimal Gmail API helper for sending an email with file attachments via OAuth."""

import base64
import mimetypes
import os
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CREDENTIALS_PATH = PROJECT_ROOT / "credentials.json"
TOKEN_PATH = PROJECT_ROOT / "token.json"
TOKEN_URI = "https://oauth2.googleapis.com/token"
MAX_TOTAL_ATTACHMENT_BYTES = 20 * 1024 * 1024  # stay under Gmail's ~25MB per-message limit


class GmailAuthError(Exception):
    pass


class GmailSendError(Exception):
    pass


def _get_gmail_service_from_env():
    """Cloud/stateless path: build credentials directly from env vars and refresh
    them — no local files, no interactive consent (there's no browser available
    in a serverless function). Used when GMAIL_REFRESH_TOKEN is set."""
    creds = Credentials(
        None,
        refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
        client_id=os.environ["GMAIL_CLIENT_ID"],
        client_secret=os.environ["GMAIL_CLIENT_SECRET"],
        token_uri=TOKEN_URI,
        scopes=SCOPES,
    )
    try:
        creds.refresh(Request())
    except Exception as e:
        raise GmailAuthError(
            f"Failed to refresh Gmail credentials from GMAIL_REFRESH_TOKEN env vars: {e}. "
            f"The refresh token may have expired or been revoked — redo the local OAuth "
            f"consent flow and update the GMAIL_REFRESH_TOKEN env var."
        ) from e
    return build("gmail", "v1", credentials=creds)


def get_gmail_service():
    if os.environ.get("GMAIL_REFRESH_TOKEN"):
        return _get_gmail_service_from_env()

    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                raise GmailAuthError(
                    f"{CREDENTIALS_PATH} not found. Download an OAuth client (Desktop app type) "
                    f"from Google Cloud Console and save it there."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def send_email_with_attachments(to: str, subject: str, body: str, attachment_paths: list) -> None:
    attachment_paths = [Path(p) for p in attachment_paths]
    total_size = sum(p.stat().st_size for p in attachment_paths)
    if total_size > MAX_TOTAL_ATTACHMENT_BYTES:
        raise GmailSendError(
            f"Total attachment size ({total_size / 1024 / 1024:.1f} MB) exceeds the "
            f"{MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024:.0f} MB safety limit for one email."
        )

    service = get_gmail_service()

    message = MIMEMultipart()
    message["to"] = to
    message["subject"] = subject
    message.attach(MIMEText(body))

    for path in attachment_paths:
        ctype, encoding = mimetypes.guess_type(str(path))
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        _, subtype = ctype.split("/", 1)
        part = MIMEApplication(path.read_bytes(), _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=path.name)
        message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    try:
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
    except Exception as e:
        raise GmailSendError(f"Gmail API send failed: {e}") from e
