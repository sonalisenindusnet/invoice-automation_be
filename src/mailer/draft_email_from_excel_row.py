"""
draft_email_from_excel_row.py

The second stage of the flow: read a specific row back OUT of whichever
tab it landed in (USA / UK / Poland 2026 — see entity_resolver.py for how
that's decided) and draft the client-facing email + invoice PDF from it.
This is the step you called "the LLM reads the excel and drafts those
emails" — the Excel row, not the original CP email, is the source of
truth from here on.

Foreign clients only, no tax: none of these three entities charge GST —
USA has no tax at all, UK/Poland use VAT (see config/tabs/*.json for each
entity's rate). India isn't wired up right now (see entity_resolver.py).

Usage:
    python draft_email_from_excel_row.py <tracker.xlsx> --latest --entity usa
    python draft_email_from_excel_row.py <tracker.xlsx> --invoice-no INT/UK/26-27/001 --entity uk

Caveat for a re-draft run via THIS CLI on the UK or Poland 2026 tabs: those
tabs have no "Invoice Description" column, so once a row is on the sheet
the original description text isn't stored anywhere — a re-read here falls
back to "Services" in the PDF and a blank spot in the email body. The live
server (email_server.py) and demo/run_demo.py both avoid this by drafting
from the row still in memory right after appending it, before anything
would need to be re-read from the sheet.
"""
import argparse
import json
import logging
import re
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent.parent.parent
SRC_DIR = BASE / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from excel.append_invoice_to_excel import header_row_index, existing_rows
from excel.entity_resolver import load_schema
from pdf.generate_invoice_pdf_intl import render_international_invoice
from utils.xlsx_io import load_workbook_with_retry
from mailer.llm_email_drafter import draft_invoice_email_body_llm

OUTPUT_DIR = BASE / "output"
logger = logging.getLogger("email_server.draft_email")


def find_row(xlsx_path, entity_key, invoice_no=None, latest=False):
    schema = load_schema(entity_key)
    wb = load_workbook_with_retry(xlsx_path, data_only=True)
    ws = wb[schema["sheet_name"]]
    rows = list(existing_rows(ws, schema))
    if not rows:
        raise ValueError(f"No data rows found in '{schema['sheet_name']}' tab")

    if latest:
        return schema, max(rows, key=lambda r: (r.get("sl_no") or 0, str(r.get("invoice_no") or "")))

    for r in rows:
        if (r.get("invoice_no") or "").strip() == invoice_no.strip():
            return schema, r
    raise ValueError(f"Invoice No {invoice_no!r} not found in '{schema['sheet_name']}' tab")


def _base_amount_for_intl(row):
    """USA/UK/Poland tabs don't have a separate base_amount column (only
    India/Kar Ventures do) — they store 'total' as the tax-INCLUSIVE amount
    plus a separate 'vat' column for the tax portion (0 for USA, which has
    no tax at all). Back out the pre-tax base so the PDF renderer — which
    re-derives tax from the base itself — doesn't double-count it."""
    if row.get("base_amount") is not None:
        return float(row["base_amount"])
    total = float(row.get("total") or 0)
    vat = float(row.get("vat") or 0)
    return total - vat


_MONTH_OF_RE = re.compile(r"for\s+the\s+month\s+of\s+([A-Za-z]+'?\d{2,4})", re.IGNORECASE)


def _address_lines(raw):
    """'client_address' arrives as one free-text field from the CP email (if
    present at all). Only split into multiple PDF lines on an explicit line
    break — a plain comma-separated address (e.g. the real Poland sample,
    "SUITE 204,MISSISSAUGA,L4W 4Y1") stays on one line, matching how real
    CP-supplied addresses have actually looked so far; a CP who pastes a
    multi-line address (one line per \\n, as UK/Singapore's real invoices
    show) gets that broken out the same way."""
    if not raw:
        return []
    raw = str(raw).strip()
    if not raw:
        return []
    return [p.strip() for p in raw.split("\n") if p.strip()]


def _derive_month_label(description):
    """The CP's Invoice Description conventionally already says something
    like '... for the month of July'26' — reuse that instead of guessing
    from the invoice date (which is often a few days into the NEXT month
    relative to the service period, so it can't be assumed to match)."""
    if not description:
        return None
    m = _MONTH_OF_RE.search(description)
    return m.group(1) if m else None


def _row_to_intl_row(entity_key, row):
    """Adapts an Excel row dict (India-style canonical keys) into the shape
    generate_invoice_pdf_intl.render_international_invoice() expects.

    Honest limitation: the CP's 'Invoice Summary' email doesn't carry a
    per-resource cost breakdown — so for USA (which shows one row per
    resource on the real invoice) this falls back to a single line item for
    the whole amount. Good enough to get a correct, correctly-taxed PDF out
    the door; not a pixel-perfect match to a multi-resource USA invoice
    until that data is available somewhere.

    client_address / po_no / po_date are OPTIONAL fields (see
    config/field_schema.json) — most CP emails won't include them, in which
    case these just come back empty and the PDF quietly omits those lines,
    exactly like it already does for due_date."""
    base_amount = _base_amount_for_intl(row)
    description = row.get("invoice_description") or ""
    return {
        "invoice_no": row.get("invoice_no"),
        "invoice_date": row.get("invoice_date"),
        "due_date": row.get("due_date") or None,
        "po_no": row.get("po_no") or None,
        "po_date": row.get("po_date") or None,
        "client_name": row.get("client_company") or "",
        "client_address_lines": _address_lines(row.get("client_address")),
        "month_label": _derive_month_label(description),
        "description": description,
        "line_items": [{"label": description or "Services", "amount": base_amount}],
        "subtotal": base_amount,
        # NEW 2026-08-13: the actual currency saved in the Excel row's
        # Currency column (from the CP email's "Currency" field), so the
        # PDF shows what was really billed instead of always falling back
        # to the entity's static default (see CURRENCY_SYMBOLS in
        # generate_invoice_pdf_intl.py -- this was the bug: PDF always
        # showed USD/$ regardless of what the email said).
        "currency": row.get("currency") or None,
    }


def render_invoice_for_entity(entity_key, row, out_path):
    intl_row = _row_to_intl_row(entity_key, row)
    return render_international_invoice(entity_key, intl_row, out_path)


def compose_email(row, entity_key="usa"):
    contact = row.get("client_contact_person") or "Team"
    to_list = [e.strip() for e in (row.get("client_mail_to") or "").split(",") if e.strip()]
    cc_list = [e.strip() for e in (row.get("int_cc_mail") or "").split(",") if e.strip()]
    description = row.get("invoice_description") or "Services"

    subject = (
        f"Invoice {row.get('invoice_no')} - {row.get('client_company')} - {description}"
    ).strip()

    currency = row.get("currency") or ""
    total = row.get("total")
    total_str = f"{currency} {float(total):,.2f}" if total not in (None, "") else ""

    # Foreign clients only: USA has no tax at all; UK/Poland show VAT if
    # their rate is non-zero. No GST/CGST/SGST wording — that's India-only
    # and India isn't wired up right now.
    tax_lines = ""
    if row.get("vat"):
        base = (float(row.get("total") or 0) - float(row.get("vat") or 0))
        tax_lines = (
            f"Base Amount: {currency} {base:,.2f}\n"
            f"VAT: {currency} {float(row.get('vat') or 0):,.2f}\n"
        )

    # 2026-08-14, per explicit instruction ("use the LLM to generate the
    # message [body] only"): the body is now drafted by Gemini
    # (llm_email_drafter.py) instead of this fixed template. The template
    # below is kept as-is and used ONLY as a fallback if the LLM call fails
    # for any reason (missing API key, rate limit, network, bad JSON reply)
    # -- a live client email should degrade to a plain-but-correct template
    # rather than block on a flaky API call.
    try:
        body = draft_invoice_email_body_llm(row)
    except Exception as exc:
        logger.warning(
            "  LLM email draft failed for %s (%s) -- falling back to the standard template",
            row.get("invoice_no"), exc,
        )
        body = f"""Dear {contact},

Please find attached the invoice {row.get('invoice_no')} for {description}.

{tax_lines}Total Payable: {total_str}

Kindly process the payment at your earliest convenience. Please let us know if any further information or supporting documents are required.

Best regards,
Accounts Team
INTglobal
"""
    return {"to": to_list, "cc": cc_list, "subject": subject, "body": body}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("--entity", required=True, choices=["usa", "uk", "poland"])
    ap.add_argument("--invoice-no", default=None)
    ap.add_argument("--latest", action="store_true")
    ap.add_argument("--out-dir", default=str(OUTPUT_DIR))
    args = ap.parse_args()

    if not args.invoice_no and not args.latest:
        raise SystemExit("Provide --invoice-no or --latest")

    schema, row = find_row(args.xlsx_path, args.entity, invoice_no=args.invoice_no, latest=args.latest)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(exist_ok=True, parents=True)
    safe_inv = (row.get("invoice_no") or "unknown").replace("/", "-")
    pdf_path = out_dir / f"invoice_{safe_inv}.pdf"

    pdf_meta = render_invoice_for_entity(args.entity, row, pdf_path)
    draft = compose_email(row, args.entity)

    result = {"row": row, "pdf": pdf_meta, "draft_email": draft}
    (out_dir / f"draft_{safe_inv}.json").write_text(
        json.dumps(result, indent=2, default=str, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))


if __name__ == "__main__":
    main()
