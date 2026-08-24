"""
poller.py

Scans the real tracker's USA/UK/Poland tabs for rows an accountant has
marked "Reviewed" that haven't been drafted yet, and for each: renders the
invoice PDF, composes the email, uploads it as a real Gmail draft with the
PDF attached, then flips that row's Email Drafted flag so it's never
picked up again.

Reuses the schema loader from save_api.excel_writer (same config/tabs/
files) and the untouched PDF renderer in pdf.generate_invoice_pdf_intl.
"""
import json
import logging
import os
import re
from pathlib import Path

from pdf.generate_invoice_pdf_intl import render_international_invoice
from save_api.excel_writer import load_schema
from utils.xlsx_io import TRACKER_LOCK
from utils.tracker_io import load_tracker_with_retry, save_tracker, tracker_ref_from_config
from tax.tax_calculator import tax_result_from_stored

from draft_mailer.email_composer import compose_email
from draft_mailer.gmail_imap import connect, find_drafts_folder, build_draft_mime, append_draft

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "draft_poller_config.json"
OUTPUT_DIR = PROJECT_ROOT / "output"

# Tabs scanned for rows whose Review Status is "Reviewed".
REVIEWABLE_ENTITIES = ["usa", "uk", "poland", "singapore"]

# Entities whose PDF actually shows a client-conditional tax line (their
# real-world rate depends on the client's own country -- see
# tax.tax_calculator). Poland and USA are always 0% regardless of client,
# so their PDF keeps using its static config/entities/*.json rate/label
# untouched -- see _row_to_intl_row() and generate_invoice_pdf_intl.py's
# _totals_block().
DYNAMIC_TAX_PDF_ENTITIES = {"uk", "singapore"}

POLL_INTERVAL_ENV = "DRAFT_POLL_INTERVAL_SECONDS"
DEFAULT_POLL_INTERVAL_SECONDS = 300

logger = logging.getLogger("draft_mailer.poller")


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def poll_interval_seconds():
    """How often (in seconds) the draft loop scans the tracker, read from
    the DRAFT_POLL_INTERVAL_SECONDS environment variable. Falls back to
    DEFAULT_POLL_INTERVAL_SECONDS if unset or not a valid integer. (Same
    pattern as save_api.excel_writer._payment_due_days().)"""
    raw = os.environ.get(POLL_INTERVAL_ENV)
    if raw is None or not raw.strip():
        return DEFAULT_POLL_INTERVAL_SECONDS
    try:
        return int(raw.strip())
    except ValueError:
        return DEFAULT_POLL_INTERVAL_SECONDS


def _tracker_path(cfg):
    """Despite the name (kept for call-site stability), this returns a
    `tracker_ref` dict pointing at the live Google Sheet -- see
    utils.tracker_io.tracker_ref_from_config for the resolution logic."""
    return tracker_ref_from_config(cfg, CONFIG_PATH.parent)


def _is_reviewed(value):
    return str(value or "").strip().lower() == "reviewed"


def _is_drafted(value):
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in ("true", "1", "yes")


def _is_mis_verified(value):
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in ("true", "1", "yes")


def _iter_rows(ws, schema):
    """Yields {key: value} for each row that actually has data."""
    cols = schema["columns"]
    start = schema.get("header_row_index", 1) + 1
    for row in ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        yield {cols[i]["key"]: row[i] if i < len(row) else None for i in range(len(cols))}


def _pending_rows(ws, schema):
    """Rows that are Reviewed, MIS-verified, and not yet drafted -- all
    three conditions must hold before a draft is ever created."""
    for row in _iter_rows(ws, schema):
        if (
            _is_reviewed(row.get("review_status"))
            and _is_mis_verified(row.get("mis_verification_done"))
            and not _is_drafted(row.get("email_drafted"))
        ):
            yield row


def _find_row_index_by_invoice_no(ws, schema, invoice_no):
    cols = schema["columns"]
    inv_col = next(i for i, c in enumerate(cols) if c["key"] == "invoice_no") + 1
    start = schema.get("header_row_index", 1) + 1
    for row_idx in range(start, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=inv_col).value
        if (val or "").strip() == (invoice_no or "").strip():
            return row_idx
    return None


def _scan_pending(tracker_ref, entity_key):
    """Returns (schema, [pending row dicts]) for one entity's tab."""
    with TRACKER_LOCK:
        schema = load_schema(entity_key)
        wb = load_tracker_with_retry(tracker_ref)
        ws = wb[schema["sheet_name"]]
        return schema, list(_pending_rows(ws, schema))


def mark_drafted(tracker_ref, entity_key, invoice_no):
    """Flips Email Drafted to True for one row by Invoice No. Returns False
    if the row can't be found (draft was still created either way)."""
    with TRACKER_LOCK:
        schema = load_schema(entity_key)
        wb = load_tracker_with_retry(tracker_ref)
        ws = wb[schema["sheet_name"]]

        row_idx = _find_row_index_by_invoice_no(ws, schema, invoice_no)
        if row_idx is None:
            return False

        col_idx = next(i for i, c in enumerate(schema["columns"]) if c["key"] == "email_drafted") + 1
        ws.cell(row=row_idx, column=col_idx, value=True)
        save_tracker(wb, tracker_ref)
        return True


_MONTH_OF_RE = re.compile(r"for\s+the\s+month\s+of\s+([A-Za-z]+'?\d{2,4})", re.IGNORECASE)


def _derive_month_label(description):
    if not description:
        return None
    m = _MONTH_OF_RE.search(description)
    return m.group(1) if m else None


def _coerce_amount(value):
    """The gspread-backed tracker adapter always returns cell values as
    strings (unlike openpyxl, which preserves numeric types) -- coerce the
    total amount back to a float here, right before it's handed to the PDF
    renderer, which sums/multiplies it. A missing or unparseable value
    becomes 0.0 rather than raising, so a blank cell never crashes
    drafting."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _row_to_intl_row(entity_key, row, tax_result):
    """Adapts a tracker row into the shape
    generate_invoice_pdf_intl.render_international_invoice() expects.
    client_address_lines/po_no/po_date/due_date aren't tracked by the
    current minimal schema, so they're simply omitted -- the renderer
    already treats a missing value as "don't show this line".

    The row's own "Total Amount" column is the POST-tax grand total (see
    excel_writer.build_row()), so the PDF's subtotal comes from
    `tax_result["subtotal"]` (reconstructed from the row -- see
    tax_result_from_stored()), never straight from row["total"] -- feeding
    the grand total in as if it were the subtotal would double the tax on
    the rendered PDF.

    `tax_result` is only attached as intl_row["tax"] for
    DYNAMIC_TAX_PDF_ENTITIES -- Poland and USA's PDF keeps rendering from
    their static, always-correct config/entities/*.json rate/label,
    exactly as before this feature."""
    description = row.get("invoice_description") or ""
    intl_row = {
        "invoice_no": row.get("invoice_no"),
        "invoice_date": row.get("invoice_date"),
        "due_date": row.get("due_date") or None,
        "client_name": row.get("client_company") or "",
        "client_address_lines": [],
        "month_label": _derive_month_label(description),
        "currency": row.get("currency") or None,
    }
    if entity_key == "usa":
        intl_row["line_items"] = [{"label": description or "Services", "amount": tax_result["subtotal"]}]
    else:
        intl_row["description"] = description
        intl_row["subtotal"] = tax_result["subtotal"]
    if entity_key in DYNAMIC_TAX_PDF_ENTITIES:
        intl_row["tax"] = tax_result
    return intl_row


def process_row(entity_key, row, imap, from_addr, drafts_folder, tracker_ref):
    invoice_no = row.get("invoice_no")
    safe_inv = (invoice_no or "unknown").replace("/", "-")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    pdf_path = OUTPUT_DIR / f"invoice_{safe_inv}.pdf"

    # Reconstructed ONCE per row, from the row's own already-saved Total
    # Amount + VAT/GST columns (NOT recomputed from a fresh rate lookup --
    # see tax_result_from_stored()'s own docstring for why), and reused for
    # both the PDF (UK/Singapore only, see DYNAMIC_TAX_PDF_ENTITIES) and
    # the drafted email (every entity) -- so the two documents can never
    # disagree with each other, or with what's already sitting in the
    # sheet, about the tax charged on this invoice.
    tax_result = tax_result_from_stored(
        entity_key, row.get("country"), _coerce_amount(row.get("total")), _coerce_amount(row.get("vat")),
    )

    intl_row = _row_to_intl_row(entity_key, row, tax_result)
    render_international_invoice(entity_key, intl_row, pdf_path)

    row_for_email = dict(row)
    row_for_email.update({
        # NOTE: "subtotal" here (pre-tax) is deliberately distinct from the
        # row's own "total" key (post-tax grand total, per
        # excel_writer.build_row()) -- email_composer.py/llm_drafter.py
        # must show THIS as the Sub-Total line, never row["total"].
        "subtotal": tax_result["subtotal"],
        "tax_name": tax_result["tax_name"],
        "tax_rate": tax_result["rate"],
        "tax_amount": tax_result["tax_amount"],
        "total_with_tax": tax_result["total"],
    })
    draft = compose_email(row_for_email)
    mime_msg = build_draft_mime(
        from_addr=from_addr, to_list=draft["to"], cc_list=draft["cc"],
        subject=draft["subject"], body_text=draft["body"],
        attachment_path=pdf_path, attachment_name=f"Invoice_{safe_inv}.pdf",
    )
    append_draft(imap, drafts_folder, mime_msg)
    logger.info("Draft created in %s for Invoice No %s", drafts_folder, invoice_no)

    if not mark_drafted(tracker_ref, entity_key, invoice_no):
        logger.warning("Could not find row for %s to flip Email Drafted -- draft was still created", invoice_no)


def run_once(cfg):
    """One poll cycle: scan every reviewable tab, draft anything pending."""
    tracker_ref = _tracker_path(cfg)

    imap, from_addr = connect(cfg)
    try:
        drafts_folder = find_drafts_folder(imap, cfg.get("drafts_folder_override"))
        drafted_count = 0

        for entity_key in REVIEWABLE_ENTITIES:
            try:
                _schema, pending = _scan_pending(tracker_ref, entity_key)
            except Exception:
                logger.exception("FAILED scanning '%s' tab for reviewed rows", entity_key)
                continue

            if pending:
                logger.info("Found %d reviewed row(s) ready to draft in '%s' tab", len(pending), entity_key)

            for row in pending:
                invoice_no = row.get("invoice_no")
                try:
                    process_row(entity_key, row, imap, from_addr, drafts_folder, tracker_ref)
                    drafted_count += 1
                except Exception:
                    logger.exception("FAILED drafting email for %s (%s)", invoice_no, entity_key)
                    continue  # one bad row must not kill the cycle or skip the others

        logger.info("Draft cycle: %d row(s) drafted", drafted_count)
    finally:
        try:
            imap.logout()
        except Exception:
            pass
