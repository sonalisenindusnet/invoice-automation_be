"""
gmail_imap.py

All the raw IMAP plumbing for the server: connect, find unread matching
emails, pull out subject/body, and APPEND a finished draft (with the
invoice PDF attached) straight into the account's Drafts folder.

Uses IMAP + a Google App Password — deliberately NOT the Gmail OAuth
connector, since that's the piece that hit the account-mismatch problem.
This only needs:
  1. 2-Step Verification turned on for the Gmail account you want to use
  2. An App Password generated for it (myaccount.google.com/apppasswords)
  3. EMAIL_ADDRESS and EMAIL_APP_PASSWORD set as environment variables

No Anthropic/Claude connector, no OAuth consent screen — this talks to
Gmail exactly the way Outlook/Thunderbird would.
"""
import email
import imaplib
import os
import smtplib
from email.header import decode_header
from email.message import EmailMessage
from email.utils import formataddr


class ImapAuthError(RuntimeError):
    pass


def connect(config):
    email_address = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    if not email_address or not app_password:
        raise ImapAuthError(
            "Set EMAIL_ADDRESS and EMAIL_APP_PASSWORD environment variables before running "
            "(see README 'Running the server' section)."
        )
    timeout = config.get("imap_timeout_seconds", 20)
    try:
        imap = imaplib.IMAP4_SSL(config["imap_host"], config["imap_port"], timeout=timeout)
    except OSError as e:
        # Without an explicit timeout, a network hang here would block forever
        # instead of failing loudly — this is what actually happens if the host
        # is unreachable (wrong network, firewall, DNS issue, etc.).
        raise ImapAuthError(
            f"Could not reach {config['imap_host']}:{config['imap_port']} within {timeout}s "
            f"({e}). Check your network connection / firewall."
        )
    try:
        imap.login(email_address, app_password)
    except imaplib.IMAP4.error as e:
        raise ImapAuthError(
            f"IMAP login failed for {email_address}: {e}. "
            "Double-check the App Password and that IMAP is enabled on the account."
        )
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
            # Folder name is the last quoted/unquoted token
            if '"' in line_str:
                return line_str.split('"')[-2]
            return line_str.rsplit(" ", 1)[-1]
    return "[Gmail]/Drafts"


def _decode_subject(raw_subject):
    if not raw_subject:
        return ""
    parts = decode_header(raw_subject)
    decoded = ""
    for text, enc in parts:
        if isinstance(text, bytes):
            decoded += text.decode(enc or "utf-8", errors="ignore")
        else:
            decoded += text
    return decoded


def _extract_body(msg):
    """Prefer text/plain; fall back to text/html. Returns (body, kind)."""
    if msg.is_multipart():
        html_fallback = None
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if "attachment" in disp:
                continue
            if ctype == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    return payload.decode(part.get_content_charset() or "utf-8", errors="ignore"), "text"
            elif ctype == "text/html" and html_fallback is None:
                payload = part.get_payload(decode=True)
                if payload:
                    html_fallback = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
        if html_fallback:
            return html_fallback, "html"
        return "", "text"
    else:
        payload = msg.get_payload(decode=True)
        body = payload.decode(msg.get_content_charset() or "utf-8", errors="ignore") if payload else ""
        return body, ("html" if msg.get_content_type() == "text/html" else "text")


def fetch_matching_emails(imap, mailbox, subject_contains, processed_message_ids, unseen_only=True):
    """
    Returns a list of dicts: {uid, message_id, subject, from, body, body_kind, date}
    for messages whose Subject contains subject_contains (case-insensitive) and
    whose Message-ID isn't already in processed_message_ids.
    """
    imap.select(mailbox)
    criterion = "UNSEEN" if unseen_only else "ALL"
    typ, data = imap.search(None, criterion)
    if typ != "OK":
        return []

    results = []
    for uid in data[0].split():
        typ, msg_data = imap.fetch(uid, "(RFC822)")
        if typ != "OK" or not msg_data or not msg_data[0]:
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        message_id = msg.get("Message-ID", "").strip()
        if message_id and message_id in processed_message_ids:
            continue

        subject = _decode_subject(msg.get("Subject", ""))
        if subject_contains.lower() not in subject.lower():
            continue

        body, body_kind = _extract_body(msg)
        results.append({
            "uid": uid,
            "message_id": message_id,
            "subject": subject,
            "from": msg.get("From", ""),
            "date": msg.get("Date", ""),
            "body": body,
            "body_kind": body_kind,
        })
    return results


def mark_seen(imap, uid):
    imap.store(uid, "+FLAGS", "\\Seen")


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
            data = f.read()
        msg.add_attachment(
            data, maintype="application", subtype="pdf",
            filename=attachment_name or os.path.basename(attachment_path),
        )
    return msg


def append_draft(imap, drafts_folder, mime_message):
    """Insert mime_message into the Drafts mailbox with the \\Draft flag — this
    is what makes it show up as a draft in the Gmail UI, ready for a human to
    review and hit Send, with no SMTP send happening at all."""
    typ, resp = imap.append(
        drafts_folder, "\\Draft", imaplib.Time2Internaldate(__import__("time").time()),
        mime_message.as_bytes(),
    )
    if typ != "OK":
        raise RuntimeError(f"Failed to append draft to {drafts_folder!r}: {resp}")
    return resp
