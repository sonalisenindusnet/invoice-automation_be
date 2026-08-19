"""IMAP helpers for creating Gmail drafts with invoice PDF attachments."""

import imaplib
import os
import time
from email.message import EmailMessage


class ImapAuthError(RuntimeError):
    """Raised when the Gmail IMAP connection cannot be established."""


def connect(config):
    email_address = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    if not email_address or not app_password:
        raise ImapAuthError("Set EMAIL_ADDRESS and EMAIL_APP_PASSWORD before running the server.")

    timeout = config.get("imap_timeout_seconds", 20)
    try:
        imap = imaplib.IMAP4_SSL(config["imap_host"], config["imap_port"], timeout=timeout)
        imap.login(email_address, app_password)
    except (OSError, imaplib.IMAP4.error) as exc:
        raise ImapAuthError(
            f"Could not connect to Gmail IMAP at {config['imap_host']}:{config['imap_port']}: {exc}"
        ) from exc
    return imap, email_address


def find_drafts_folder(imap, override=None):
    if override:
        return override
    typ, data = imap.list()
    if typ != "OK":
        return "[Gmail]/Drafts"
    for line in data:
        line_str = line.decode(errors="ignore") if isinstance(line, bytes) else line
        if "\\Drafts" in line_str:
            return line_str.split('"')[-2] if '"' in line_str else line_str.rsplit(" ", 1)[-1]
    return "[Gmail]/Drafts"


def build_draft_mime(from_addr, to_list, cc_list, subject, body_text, attachment_path=None, attachment_name=None):
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.set_content(body_text)

    if attachment_path:
        with open(attachment_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=attachment_name or os.path.basename(attachment_path),
            )
    return msg


def append_draft(imap, drafts_folder, mime_message):
    """Append a message to Gmail's Drafts folder; it is never sent."""
    typ, response = imap.append(
        drafts_folder,
        "\\Draft",
        imaplib.Time2Internaldate(time.time()),
        mime_message.as_bytes(),
    )
    if typ != "OK":
        raise RuntimeError(f"Failed to append draft to {drafts_folder!r}: {response}")
    return response
