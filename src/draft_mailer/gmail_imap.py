"""
gmail_imap.py

Creates a real Gmail draft over IMAP -- no SMTP, nothing is ever sent.
Logs in, builds a MIME message with the invoice PDF attached, and uploads
it to the account's Drafts folder with the IMAP \\Draft flag set, which is
what makes Gmail treat it as a draft rather than a delivered message.

Credentials (EMAIL_ADDRESS, EMAIL_APP_PASSWORD -- a Gmail App Password,
not the account password) come from the environment (see utils/env_loader.py
/ .env.example). Connection settings (host/port/timeout) come from the
poller's own config.
"""
import email.utils
import imaplib
import mimetypes
import os
import time
from email.message import EmailMessage


class ImapAuthError(Exception):
    """Raised when IMAP login fails -- missing/wrong credentials, or the
    account rejected the connection."""


def connect(cfg):
    """Logs into IMAP and returns (imap_connection, from_addr). Raises
    ImapAuthError on any login failure."""
    address = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    if not address or not app_password:
        raise ImapAuthError("EMAIL_ADDRESS / EMAIL_APP_PASSWORD are not set in the environment")

    try:
        imap = imaplib.IMAP4_SSL(cfg["imap_host"], cfg["imap_port"], timeout=cfg.get("imap_timeout_seconds", 45))
        imap.login(address, app_password)
    except (imaplib.IMAP4.error, OSError) as exc:
        raise ImapAuthError(f"IMAP login failed for {address}: {exc}") from exc

    return imap, address


def find_drafts_folder(imap, override=None):
    """Returns the mailbox name to APPEND drafts into. Uses `override` if
    given; otherwise looks for the folder IMAP reports with the \\Drafts
    special-use attribute (RFC 6154); falls back to Gmail's default
    English name if that lookup doesn't turn anything up."""
    if override:
        return override

    try:
        status, folders = imap.list()
        if status == "OK":
            for raw in folders:
                decoded = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
                if "\\Drafts" in decoded:
                    # Folder name is the last quoted or bare token on the line.
                    parts = decoded.split(' "/" ')
                    if len(parts) == 2:
                        return parts[1].strip('"')
    except imaplib.IMAP4.error:
        pass

    return "[Gmail]/Drafts"


def build_draft_mime(from_addr, to_list, cc_list, subject, body_text, attachment_path=None, attachment_name=None):
    """Builds a MIME message (plain-text body, PDF attached if given) and
    returns it ready for append_draft()."""
    msg = EmailMessage()
    msg["From"] = from_addr
    if to_list:
        msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg["Date"] = email.utils.formatdate(localtime=True)
    msg.set_content(body_text)

    if attachment_path:
        content_type, _ = mimetypes.guess_type(str(attachment_name or attachment_path))
        maintype, subtype = (content_type or "application/pdf").split("/", 1)
        with open(attachment_path, "rb") as f:
            msg.add_attachment(
                f.read(), maintype=maintype, subtype=subtype,
                filename=attachment_name or str(attachment_path),
            )

    return msg


def append_draft(imap, drafts_folder, mime_msg):
    """Uploads `mime_msg` into `drafts_folder` flagged \\Draft -- a Gmail
    draft, never sent."""
    status, _ = imap.append(drafts_folder, "\\Draft", imaplib.Time2Internaldate(time.time()), mime_msg.as_bytes())
    if status != "OK":
        raise RuntimeError(f"IMAP APPEND to {drafts_folder!r} failed: {status}")
