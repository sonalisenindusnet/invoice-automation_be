"""
email_server.py

The actual server you asked for: runs continuously, polls on an interval,
and does TWO SEPARATE passes every cycle — split apart so a human reviews
every new row in Excel before any invoice/email ever gets created:

PASS 1 — INTAKE (for every unread email whose Subject matches
config/email_server_config.json's "subject_contains"):
  1. Parses it                          (parse_invoice_summary.py)
  2. Figures out WHICH tab the named client belongs to — USA, UK, or
     Poland 2026 (foreign clients only, no GST/India right now)
                                          (entity_resolver.py)
  3. Checks it's not a duplicate, then calls the real MIS project-info API
     with the CP email's PF ID (added 2026-08-12 — replaces the earlier
     "bypass MIS validation" stance now that a real endpoint exists) and
     appends a row to that tab in ONE shot: email-parsed fields + whatever
     the MIS API added, Review Status = "Pending Review", and a new
     "MIS Verified" boolean — True only if that MIS call actually
     succeeded. A failed/timed-out/unreachable MIS call does NOT block the
     append: the row still goes in, just with MIS Verified = False, so the
     accountant knows to check that project ID by hand before reviewing it.
                                          (append_invoice_to_excel.py, mis/mis_api.py)
  4. Marks the source email read and remembers its Message-ID, so it's
     never processed twice even across restarts
  STOPS THERE. No PDF, no draft email, no send — that only happens once a
  human changes that row's "Review Status" cell to "Reviewed" in Excel.

If step 2 can't determine the client's tab (not in
config/company_tab_map.json and not found in any existing tab), the email
is logged as "entity_unresolved" and left UNREAD/unprocessed — nothing is
appended anywhere. Add the company to company_tab_map.json and the next
poll will pick it up automatically; no data is ever written to a guessed
tab.

PASS 2 — INVOICE GENERATION (every cycle, across USA/UK/Poland 2026):
  Scans each tab's "Review Status" column for rows a human has already set
  to exactly "Reviewed" AND whose "Email Drafted" column isn't already
  True — every other combination (blank, "Pending Review", historical rows
  that predate this column, or a "Reviewed" row already drafted) is
  skipped, so nothing is ever double-processed or guessed into being ready.
  For each such row, reading the row BACK OUT of Excel (not from memory —
  Pass 1 may have appended it minutes, hours, or poll-cycles ago):
  1. Generates the invoice PDF, using that entity's own template
                                          (generate_invoice_pdf_intl.py)
  2. Builds the client email — no tax lines for USA, VAT for UK/Poland
                                          (draft_email_from_excel_row.py's compose_email)
  3. Uploads it as a real Gmail draft, PDF attached (gmail_imap.py)
  4. Flips that row's "Email Drafted" to True so it's never picked up
     again — "Review Status" is deliberately left untouched, a permanent
     record of the human's own approval, separate from the pipeline's own
     progress          (append_invoice_to_excel.py)
  5. Tries to write EMAIL STATUS='Draft Created' back into that Excel row —
     none of USA/UK/Poland 2026 have this column in your real tracker, so
     this is currently always a no-op (kept generic in case a tracked tab
     is added later)

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
from parsing.parse_invoice_summary import parse_invoice_summary
from excel.append_invoice_to_excel import (
    append_invoice, update_email_status,
    find_rows_ready_for_invoicing, mark_email_drafted,
)
from excel.entity_resolver import resolve_entity_from_email_text
from mailer.draft_email_from_excel_row import compose_email, render_invoice_for_entity
from mailer.gmail_imap import (
    connect, find_drafts_folder, fetch_matching_emails, mark_seen,
    build_draft_mime, append_draft, ImapAuthError,
)

CONFIG_PATH = BASE / "config" / "email_server_config.json"

# The only entities live in the Excel routing right now (India isn't wired
# into entity_resolver.ENTITY_SCHEMA_PATHS yet) — Pass 2 scans exactly these
# tabs for rows a human has marked "Reviewed".
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
    cfg["state_file"] = _resolve(cfg["state_file"])
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


def intake_one_email(cfg, email_item, output_dir):
    """PASS 1 for a single matching email: parse + append ONLY. No PDF, no
    draft, no send — the row lands with Review Status = 'Pending Review'
    and waits for a human. Returns a summary dict."""
    debug_path = _dump_debug_body(output_dir, email_item)
    logger.info("  raw body saved to %s (for debugging parse mismatches)", debug_path)

    body = email_item["body"]
    source_type = "html" if email_item["body_kind"] == "html" else "text"

    data, warnings = parse_invoice_summary(body, source_type=source_type)
    for w in warnings:
        logger.warning("  parse warning: %s", w)

    # Tab routing (2026-08-13): the account team now states the tab
    # explicitly, as the FIRST line of the email (e.g. "Tab: UK") -- this
    # REPLACES the old company-name-based lookup entirely for the live
    # intake path. Never falls back to it, never guesses. See
    # entity_resolver.py's resolve_entity_from_email_text() for the exact
    # required format.
    entity_key, tab_resolution, tab_message = resolve_entity_from_email_text(body)
    if entity_key is None:
        logger.warning(
            "  TAB NOT RESOLVED (%s) — %s (email left unread so it retries once fixed)",
            tab_resolution, tab_message,
        )
        return {
            "status": "entity_unresolved",
            "detail": {
                "message": tab_message,
                "resolution": tab_resolution,
                "client_company": data.get("client_company"),
            },
        }

    append_result = append_invoice(
        cfg["tracker_ref"], data,
        requested_by=email_item.get("from") or cfg.get("requested_by_fallback", ""),
        entity_key=entity_key,
        debug_dir=output_dir,
    )

    if append_result["status"] == "entity_unresolved":
        logger.warning(
            "  UNRESOLVED CLIENT %r — %s (email left unread so it retries once you fix the mapping)",
            append_result.get("client_company"), append_result["message"],
        )
        return {"status": "entity_unresolved", "detail": append_result}

    if append_result["status"] == "duplicate_flagged":
        logger.info("  DUPLICATE — not appended: %s", append_result["message"])
        return {"status": "duplicate", "detail": append_result}

    entity_key = append_result["entity_key"]
    row = append_result["row"]
    invoice_no = row["invoice_no"]
    logger.info(
        "  appended row to '%s' tab (entity=%s, resolution=%s), Invoice No %s — "
        "Review Status = 'Pending Review', waiting for manual review",
        append_result["sheet"], entity_key, append_result.get("resolution"), invoice_no,
    )

    mis_check = append_result.get("mis_verification") or {}
    if mis_check.get("verified"):
        logger.info("  MIS check OK for PF ID %r — MIS Verified = True", row.get("pf_id"))
    else:
        logger.warning(
            "  MIS check NOT verified for PF ID %r (%s) — row appended anyway with MIS Verified = "
            "False; accountant should check the project ID manually",
            row.get("pf_id"), mis_check.get("error") or "unknown reason",
        )

    return {"status": "appended", "invoice_no": invoice_no, "entity_key": entity_key, "row": row}


def process_reviewed_row(cfg, from_addr, drafts_folder, imap, entity_key, row, output_dir):
    """PASS 2 for a single row a human has already marked 'Reviewed': render
    the PDF, draft the client email, upload it, then flip Review Status to
    'Invoice Generated' so it's never picked up again. `row` here was read
    straight back OUT of Excel by find_rows_ready_for_invoicing — NOT held
    in memory from Pass 1, which may have run in a different poll cycle,
    or even before a server restart."""
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


def process_reviewed_rows(cfg, from_addr, drafts_folder, imap, output_dir):
    """PASS 2 top level: scans every reviewable tab (USA/UK/Poland) for rows
    marked 'Reviewed' and processes each one. One bad row must not stop the
    rest — same resilience rule as Pass 1's per-email try/except."""
    total = 0
    for entity_key in REVIEWABLE_ENTITIES:
        try:
            ready_rows = find_rows_ready_for_invoicing(cfg["tracker_ref"], entity_key)
        except Exception as e:
            logger.exception("  FAILED scanning '%s' tab for reviewed rows: %s", entity_key, e)
            continue

        if ready_rows:
            logger.info("Pass 2: %d row(s) marked 'Reviewed' in '%s' tab", len(ready_rows), entity_key)

        for row in ready_rows:
            invoice_no = row.get("invoice_no")
            logger.info("Processing reviewed row: entity=%s Invoice No=%s", entity_key, invoice_no)
            try:
                result = process_reviewed_row(cfg, from_addr, drafts_folder, imap, entity_key, row, output_dir)
                logger.info("  result: %s", result["status"])
                total += 1
            except Exception as e:
                logger.exception("  FAILED generating invoice/draft for %s: %s", invoice_no, e)
                continue  # one bad row must not kill the loop or skip the others
    return total


def run_once(cfg):
    output_dir = Path(cfg["tracker_xlsx_path"]).resolve().parent
    processed_ids = load_processed_ids(cfg["state_file"])

    imap, from_addr = connect(cfg)
    try:
        drafts_folder = find_drafts_folder(imap, cfg.get("drafts_folder_override"))

        # ---- PASS 1: read new CP emails, append rows only ----
        matches = fetch_matching_emails(
            imap, cfg["mailbox"], cfg["subject_contains"], processed_ids, unseen_only=True
        )
        logger.info("Pass 1 (intake): %d matching unread email(s) found", len(matches))

        for item in matches:
            logger.info("Processing: %r from %s", item["subject"], item["from"])
            try:
                result = intake_one_email(cfg, item, output_dir)
                logger.info("  result: %s", result["status"])
            except Exception as e:
                logger.exception("  FAILED processing this email: %s", e)
                continue  # one bad email must not kill the loop or skip the others

            if result["status"] == "entity_unresolved":
                # Not the email's fault — a config/mapping gap. Leave it
                # unread and unprocessed so it's retried automatically once
                # the company is added to config/company_tab_map.json.
                continue

            if item["message_id"]:
                processed_ids.add(item["message_id"])
            if cfg.get("mark_source_email_as_read", True):
                mark_seen(imap, item["uid"])

        save_processed_ids(cfg["state_file"], processed_ids)

        # ---- PASS 2: generate invoice + draft email for reviewed rows ----
        drafted_count = process_reviewed_rows(cfg, from_addr, drafts_folder, imap, output_dir)
        logger.info("Pass 2 (invoice generation): %d row(s) drafted this cycle", drafted_count)
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
