"""
email_server.py

The actual server you asked for: runs continuously, polls the inbox on an
interval, and for every unread email whose Subject matches
config/email_server_config.json's "subject_contains":

  1. Parses it                          (parse_invoice_summary.py)
  2. Appends a row to the India tab      (append_invoice_to_excel.py)
  3. Generates the invoice PDF           (generate_invoice_pdf.py)
  4. Builds the client email             (draft_email_from_excel_row.py's compose_email)
  5. Uploads it as a real Gmail draft, PDF attached (gmail_imap.py)
  6. Writes EMAIL STATUS='Draft Created' back into that Excel row (mark_email_sent.py)
  7. Marks the source email read and remembers its Message-ID, so it's
     never processed twice even across restarts

Setup (one time):
  1. Turn on 2-Step Verification for the Gmail account to monitor
  2. Generate an App Password: myaccount.google.com/apppasswords
  3. Copy .env.example (in the project root) to ".env" and fill in your real
     EMAIL_ADDRESS and EMAIL_APP_PASSWORD there — this file stays on your
     machine only, is in .gitignore, and this script reads it automatically.
     (A real environment variable set in your shell still overrides it, if
     you'd rather do it that way for a one-off test.)

Run:
  python email_server.py                 # runs forever, polling on the configured interval
  python email_server.py --once          # single pass, then exit (good for testing)
  python email_server.py --interval 30   # override the poll interval (seconds)
"""
import argparse
import json
import logging
import sys
import time
from pathlib import Path

from scripts.env_loader import load_env_file
from scripts.parse_invoice_summary import parse_invoice_summary
from src.excel.append_invoice_to_excel import append_invoice, update_email_status
from src.mailer.draft_email_from_excel_row import find_row, compose_email
from scripts.generate_invoice_pdf import render_invoice_pdf
from scripts.gmail_imap import (
    connect, find_drafts_folder, fetch_matching_emails, mark_seen,
    build_draft_mime, append_draft, ImapAuthError,
)

BASE = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE / "config" / "email_server_config.json"

logger = logging.getLogger("email_server")


def _resolve(path_str):
    """Paths in the config are relative to the config/ directory itself."""
    p = Path(path_str)
    return str(p if p.is_absolute() else (CONFIG_PATH.parent / p).resolve())


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["tracker_xlsx_path"] = _resolve(cfg["tracker_xlsx_path"])
    cfg["state_file"] = _resolve(cfg["state_file"])
    cfg["log_file"] = _resolve(cfg["log_file"])
    return cfg


def setup_logging(log_file):
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )


def load_processed_ids(state_file):
    p = Path(state_file)
    if p.exists():
        return set(json.loads(p.read_text(encoding="utf-8")))
    return set()


def save_processed_ids(state_file, ids):
    Path(state_file).write_text(json.dumps(sorted(ids), indent=2), encoding="utf-8")


def _dump_debug_body(output_dir, email_item):
    """Always saves the most recently seen matching email's raw body, so a
    parsing mismatch can be diagnosed by looking at exactly what came in —
    without needing to resend or forward anything."""
    ext = "html" if email_item["body_kind"] == "html" else "txt"
    debug_path = Path(output_dir) / f"debug_last_email_body.{ext}"
    header = (
        f"Subject: {email_item.get('subject')}\n"
        f"From: {email_item.get('from')}\n"
        f"Date: {email_item.get('date')}\n"
        f"body_kind: {email_item.get('body_kind')}\n"
        f"{'-' * 60}\n"
    )
    debug_path.write_text(header + email_item["body"], encoding="utf-8")
    return debug_path


def process_one_email(cfg, from_addr, drafts_folder, imap, email_item, output_dir):
    """Runs stages 1-6 for a single matching email. Returns a summary dict."""
    debug_path = _dump_debug_body(output_dir, email_item)
    logger.info("  raw body saved to %s (for debugging parse mismatches)", debug_path)

    body = email_item["body"]
    source_type = "html" if email_item["body_kind"] == "html" else "text"

    data, warnings = parse_invoice_summary(body, source_type=source_type)

    for w in warnings:
        logger.warning("  parse warning: %s", w)

    append_result = append_invoice(
        cfg["tracker_xlsx_path"], data, entity_key=data.get("tab"),
        requested_by=email_item.get("from") or cfg.get("requested_by_fallback", ""),
    )
    if append_result["status"] == "duplicate_flagged":
        logger.info("  DUPLICATE — not appended: %s", append_result["message"])
        return {"status": "duplicate", "detail": append_result}

    invoice_no = append_result["row"]["invoice_no"]
    logger.info("  appended row, Invoice No %s", invoice_no)

    _, row = find_row(cfg["tracker_xlsx_path"], invoice_no=invoice_no)
    safe_inv = invoice_no.replace("/", "-")
    pdf_path = output_dir / f"invoice_{safe_inv}.pdf"
    render_invoice_pdf(row, pdf_path)

    draft = compose_email(row)

    mime_msg = build_draft_mime(
        from_addr=from_addr, to_list=draft["to"], cc_list=draft["cc"],
        subject=draft["subject"], body_text=draft["body"],
        attachment_path=pdf_path, attachment_name=f"Invoice_{safe_inv}.pdf",
    )
    append_draft(imap, drafts_folder, mime_msg)
    logger.info("  draft created in %s", drafts_folder)

    update_email_status(
        cfg["tracker_xlsx_path"], invoice_no, "Draft Created",
        sent_date=time.strftime("%Y-%m-%d"), sent_to=draft["to"], sent_cc=draft["cc"],
    )

    return {"status": "drafted", "invoice_no": invoice_no, "pdf_path": str(pdf_path), "draft": draft}


def run_once(cfg):
    output_dir = Path(cfg["tracker_xlsx_path"]).resolve().parent
    processed_ids = load_processed_ids(cfg["state_file"])

    imap, from_addr = connect(cfg)
    try:
        drafts_folder = find_drafts_folder(imap, cfg.get("drafts_folder_override"))
        matches = fetch_matching_emails(
            imap, cfg["mailbox"], cfg["subject_contains"], processed_ids, unseen_only=True
        )
        logger.info("Poll: %d matching unread email(s) found", len(matches))

        for item in matches:
            logger.info("Processing: %r from %s", item["subject"], item["from"])
            try:
                result = process_one_email(cfg, from_addr, drafts_folder, imap, item, output_dir)
                logger.info("  result: %s", result["status"])
            except Exception as e:
                logger.exception("  FAILED processing this email: %s", e)
                continue  # one bad email must not kill the loop or skip the others

            if item["message_id"]:
                processed_ids.add(item["message_id"])
            if cfg.get("mark_source_email_as_read", True):
                mark_seen(imap, item["uid"])

        save_processed_ids(cfg["state_file"], processed_ids)
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--once", action="store_true", help="Run a single poll pass and exit")
    ap.add_argument("--interval", type=int, default=None, help="Override poll_interval_seconds from config")
    args = ap.parse_args()

    env_found = load_env_file()  # loads .env (if present) into os.environ; never logs values
    cfg = load_config()
    setup_logging(cfg["log_file"])
    interval = args.interval or cfg["poll_interval_seconds"]

    if "EMAIL_ADDRESS" in env_found:
        logger.info("Loaded credentials from .env (EMAIL_ADDRESS=%s)", env_found["EMAIL_ADDRESS"])
    logger.info("Starting email_server (mailbox=%s, filter=%r, interval=%ss)",
                cfg["mailbox"], cfg["subject_contains"], interval)

    try:
        if args.once:
            try:
                run_once(cfg)
            except ImapAuthError as e:
                logger.error("Connection/auth error: %s", e)
                sys.exit(1)
            return
        while True:
            try:
                run_once(cfg)
            except ImapAuthError as e:
                logger.error("Auth error, will retry next interval: %s", e)
            except Exception as e:
                logger.exception("Unexpected error in poll loop: %s", e)
            time.sleep(interval)
    except KeyboardInterrupt:
        logger.info("Stopped by user (Ctrl+C).")


if __name__ == "__main__":
    main()
