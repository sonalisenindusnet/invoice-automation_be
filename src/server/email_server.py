"""
email_server.py

The server polls the Google Sheet on an interval. Across the USA, UK, and
Poland 2026 tabs, it selects rows where "Email Drafted" is explicitly
False, then:
  1. Generates the invoice PDF, using that entity's own template
  2. Builds the client email — no tax lines for USA, VAT for UK/Poland
  3. Uploads it as a real Gmail draft, PDF attached (gmail_imap.py)
  4. Flips that row's "Email Drafted" to True so it's never picked up
     again.

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

BASE = Path(__file__).resolve().parent.parent.parent
SRC_DIR = BASE / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from utils.env_loader import load_env_file
from excel.append_invoice_to_excel import (
    update_email_status, find_rows_ready_for_invoicing, mark_email_drafted,
)
from mailer.draft_email_from_excel_row import compose_email, render_invoice_for_entity
from mailer.gmail_imap import (
    connect, find_drafts_folder, build_draft_mime, append_draft, ImapAuthError,
)

CONFIG_PATH = BASE / "config" / "email_server_config.json"

# Tabs scanned for rows whose Email Drafted value is explicitly False.
REVIEWABLE_ENTITIES = ["usa", "uk", "poland"]

logger = logging.getLogger("email_server")


def _resolve(path_str):
    """Paths in the config are relative to the config/ directory itself."""
    p = Path(path_str)
    return str(p if p.is_absolute() else (CONFIG_PATH.parent / p).resolve())


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # tracker_xlsx_path is kept (and still resolved) purely as the anchor
    # for output_dir (PDFs/debug files/logs) below -- the real local
    # "Invoice Traker.xlsx" file still exists in that folder, it's just no
    # longer read/written by the live pipeline once tracker_source is
    # "google_sheets".
    cfg["tracker_xlsx_path"] = _resolve(cfg["tracker_xlsx_path"])
    cfg["log_file"] = _resolve(cfg["log_file"])

    # 2026-08-14: added as part of the Google Sheets migration (see
    # tracker_io.py's module docstring for why the local .xlsx file could
    # no longer be used for live reads/writes). "tracker_ref" is what every
    # actual data read/write call below uses -- either the Sheets dict, or
    # (if tracker_source is absent/"excel") the same local path as before,
    # for backward compatibility.
    if cfg.get("tracker_source") == "google_sheets":
        cfg["google_service_account_json"] = _resolve(cfg["google_service_account_json"])
        cfg["tracker_ref"] = {
            "type": "google_sheets",
            "sheet_id": cfg["google_sheet_id"],
            "service_account_json": cfg["google_service_account_json"],
        }
    else:
        cfg["tracker_ref"] = cfg["tracker_xlsx_path"]

    return cfg


def setup_logging(log_file):
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )


def process_pending_row(cfg, from_addr, drafts_folder, imap, entity_key, row, output_dir):
    """Create a PDF and Gmail draft for one row with Email Drafted=False."""
    invoice_no = row["invoice_no"]
    safe_inv = (invoice_no or "unknown").replace("/", "-")
    pdf_path = output_dir / f"invoice_{safe_inv}.pdf"
    render_invoice_for_entity(entity_key, row, pdf_path)

    draft = compose_email(row, entity_key)

    mime_msg = build_draft_mime(
        from_addr=from_addr, to_list=draft["to"], cc_list=draft["cc"],
        subject=draft["subject"], body_text=draft["body"],
        attachment_path=pdf_path, attachment_name=f"Invoice_{safe_inv}.pdf",
    )
    append_draft(imap, drafts_folder, mime_msg)
    logger.info("  draft created in %s for Invoice No %s", drafts_folder, invoice_no)

    drafted_result = mark_email_drafted(
        cfg["tracker_ref"], invoice_no, entity_key=entity_key,
    )
    if drafted_result["status"] not in ("updated", "not_supported"):
        logger.warning("  could not flip Email Drafted for %s: %s", invoice_no, drafted_result)

    status_result = update_email_status(
        cfg["tracker_ref"], invoice_no, "Draft Created",
        sent_date=time.strftime("%Y-%m-%d"), sent_to=draft["to"], sent_cc=draft["cc"],
        entity_key=entity_key,
    )
    if status_result["status"] == "not_supported":
        logger.info("  %s", status_result["message"])

    return {"status": "drafted", "invoice_no": invoice_no, "entity_key": entity_key,
            "pdf_path": str(pdf_path), "draft": draft}


def process_pending_rows(cfg, from_addr, drafts_folder, imap, output_dir):
    """Process every USA/UK/Poland row whose Email Drafted flag is False."""
    total = 0
    for entity_key in REVIEWABLE_ENTITIES:
        try:
            ready_rows = find_rows_ready_for_invoicing(cfg["tracker_ref"], entity_key)
        except Exception as e:
            logger.exception("  FAILED scanning '%s' tab for undrafted rows: %s", entity_key, e)
            continue

        if ready_rows:
            logger.info("Found %d row(s) with Email Drafted=False in '%s' tab", len(ready_rows), entity_key)

        for row in ready_rows:
            invoice_no = row.get("invoice_no")
            logger.info("Processing undrafted row: entity=%s Invoice No=%s", entity_key, invoice_no)
            try:
                result = process_pending_row(cfg, from_addr, drafts_folder, imap, entity_key, row, output_dir)
                logger.info("  result: %s", result["status"])
                total += 1
            except Exception as e:
                logger.exception("  FAILED generating invoice/draft for %s: %s", invoice_no, e)
                continue  # one bad row must not kill the loop or skip the others
    return total


def run_once(cfg):
    output_dir = Path(cfg["tracker_xlsx_path"]).resolve().parent

    imap, from_addr = connect(cfg)
    try:
        drafts_folder = find_drafts_folder(imap, cfg.get("drafts_folder_override"))
        drafted_count = process_pending_rows(cfg, from_addr, drafts_folder, imap, output_dir)
        logger.info("Invoice generation: %d row(s) drafted this cycle", drafted_count)
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
    logger.info("Starting invoice draft server (Google Sheets scan, interval=%ss)", interval)

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
