"""
append_invoice_to_excel.py

The "normal process automation" step: take the parsed CP email data
(from parse_invoice_summary.py) and append one row into the CORRECT tab
of the Invoice Tracker workbook — USA, UK, or Poland 2026 — no LLM
involved, just deterministic rules: figure out which tab this client
belongs to, compute tax (none for USA, VAT for UK/Poland — no GST is
ever applied here), assign the next invoice number, and append.

Duplicate check: DISABLED as of 2026-08-13 (see the commented-out block in
append_invoice() for why) — a project's PF ID and invoice amount are
usually the same every month, so the old (pf_id, base_amount) check
flagged legitimate monthly resubmissions as duplicates. Every matching
email now always gets appended as a new row.

Which tab a client belongs to is decided by entity_resolver.py — see that
file's docstring. Every tab's schema (columns, tax rule, invoice-number
series) lives in config/tabs/*.json, editable without touching this file.

Usage:
    python append_invoice_to_excel.py <tracker.xlsx> <parsed_data.json> \
        [--entity usa|uk|poland] \
        [--requested-by "Name"] [--invoice-date YYYY-MM-DD] [--dry-run]

    (--entity is optional — omit it and the company name in the parsed
    data is used to look up the right tab automatically.)
"""
import argparse
import json
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl

SRC_DIR = Path(__file__).resolve().parent.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from excel.entity_resolver import load_schema, resolve_entity_for_company, ENTITY_SCHEMA_PATHS
from mis.mis_api import fetch_project_info
# 2026-08-14: swapped from utils.xlsx_io's openpyxl-only loader to
# tracker_io's loader, which transparently opens either a local .xlsx file
# (a plain string, unchanged behavior) or the live Google Sheet (a
# {"type": "google_sheets", ...} dict, per config/email_server_config.json's
# "tracker_source") -- see tracker_io.py's module docstring for why. Every
# `xlsx_path` parameter below is really "tracker_ref" now; kept the name
# `xlsx_path` to minimize the diff across this already-carefully-tested file.
from utils.tracker_io import load_tracker_with_retry as load_workbook_with_retry

# Email Drafted is owned by the draft server. New rows start as False; once
# their PDF and Gmail draft are created, it becomes True. Review Status is
# retained as tracker data but is not part of the draft-server selection rule.
REVIEW_STATUS_PENDING = "Pending Review"

# MIS project-info check, added 2026-08-12 (real endpoint now exists --
# see mis/mis_api.py). Called once per append, right after the duplicate
# check passes, using the CP email's own pf_id. "MIS Verified" is a
# pipeline-owned boolean, same pattern as "Email Drafted": True only when
# the API call actually succeeded for that pf_id. Any failure (API down,
# timeout, expired token, project not found) does NOT block the append --
# per explicit instruction, the row still goes in with this column False,
# "so accountant manually check the project ID."
MIS_VERIFIED_KEY = "mis_verification_done"

# Added 2026-08-14, per explicit instruction ("add the invoice creation
# date, and payment due date today + 7 day logic"). Invoice creation date
# defaults to date.today().isoformat() whenever the caller does not supply
# one. Payment
# due date was always blank before this -- now computed as invoice_date + 7
# calendar days for every new row, every tab (USA/UK/Poland all use the
# same rule; no tab-specific override requested). Never blocks an append:
# if invoice_date isn't a clean YYYY-MM-DD string for some reason, due_date
# is just left blank, same as its old default.
DEFAULT_DUE_DATE_OFFSET_DAYS = 7


def _compute_due_date(invoice_date_str, days=DEFAULT_DUE_DATE_OFFSET_DAYS):
    if not invoice_date_str:
        return ""
    try:
        d = date.fromisoformat(str(invoice_date_str)[:10])
    except (TypeError, ValueError):
        return ""
    return (d + timedelta(days=days)).isoformat()


def _col_letter(idx):
    return openpyxl.utils.get_column_letter(idx)


def ensure_sheet(wb, schema):
    """Only ever creates a sheet for a tab that doesn't exist yet — USA/UK/
    Poland 2026 already exist with real data in your tracker, so this is a
    no-op for them in normal use."""
    name = schema["sheet_name"]
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    if schema.get("legend_row"):
        ws.append(schema["legend_row"])
    ws.append([c["header"] for c in schema["columns"]])
    return ws


def header_row_index(schema):
    """Row the column headers sit on for this tab — USA/UK/Poland 2026 all
    have the header on row 1 (header_row_index=1), no legend row above it."""
    return schema.get("header_row_index", 1)


def existing_rows(ws, schema):
    """Yield dicts of {key: value} for each existing data row."""
    cols = schema["columns"]
    start = header_row_index(schema) + 1
    for row in ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        yield {cols[i]["key"]: row[i] if i < len(row) else None for i in range(len(cols))}


def next_data_row(ws, schema):
    """The row index to write a NEW row at — right after the last row that
    actually HAS data, never ws.max_row on its own. openpyxl's ws.max_row
    counts any row that has ever had a cell touched, including rows that
    are only pre-formatted (borders/number formats applied hundreds of rows
    ahead of the real data, a common thing in hand-built Excel templates —
    exactly what your real USA/UK/Poland tabs do, pre-styled to row 1000).
    ws.append() blindly writes at max_row + 1, which is how 088-092 ended up
    stranded at rows 1001-1005 with an ~900-row blank gap above them
    instead of continuing at row 92. This walks the sheet the same way
    existing_rows() does (skipping pure-formatting rows) to find where data
    actually ends, so every future append lands immediately after it."""
    cols = schema["columns"]
    start = header_row_index(schema) + 1
    last_data_row = start - 1
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True), start=start
    ):
        if row is not None and not all(v is None for v in row[:len(cols)]):
            last_data_row = row_idx
    return last_data_row + 1


def _has_column(schema, key):
    return any(c["key"] == key for c in schema["columns"])


def next_sl_no(ws, schema):
    """None of USA/UK/Poland 2026 have a Sl No column — they number by
    Invoice No only. Kept generic in case a future tab needs it."""
    if not _has_column(schema, "sl_no"):
        return None
    n = 0
    for r in existing_rows(ws, schema):
        try:
            n = max(n, int(r.get("sl_no") or 0))
        except (TypeError, ValueError):
            pass
    return n + 1


def next_invoice_no(ws, schema):
    numbering = schema["invoice_numbering"]
    prefix = f"{numbering['prefix']}/{numbering['financial_year']}/"
    max_seq = 0
    rows = list(existing_rows(ws, schema))
    for r in rows:
        inv = r.get("invoice_no")
        if isinstance(inv, str) and inv.startswith(prefix):
            tail = inv[len(prefix):]
            if tail.isdigit():
                max_seq = max(max_seq, int(tail))

    # CRITICAL BUG found 2026-08-14 via a real append test on Poland 2026:
    # that tab's real Invoice No. column is a FORMULA
    # (="INT/PL/2026/"&TEXT(ROW(I{n-1}),"000")) with no cached value stored
    # in the file (confirmed at the raw XML level -- an empty <v/>), so
    # openpyxl's .value returns the literal formula TEXT, which never
    # starts with `prefix` -- max_seq silently stayed 0 for a tab with 76
    # real existing invoices, and this would have handed out
    # "INT/PL/2026/001" as the "next" number, DUPLICATING an
    # already-issued real invoice, the moment a genuine new Poland row
    # went through the live pipeline.
    #
    # Fix does NOT touch any existing row/formula/value in the sheet --
    # it only changes how the code computes the next number to hand out.
    # Safe fallback: if literal-text parsing found nothing (max_seq == 0)
    # despite real existing rows being present, use the COUNT of existing
    # rows as max_seq instead. This is exactly correct for Poland's formula
    # (which ties sequence directly to row position -- row R always means
    # seq R-1, so no gap is structurally possible), and is a safe no-op
    # for USA/UK (their Invoice No. is literal text, so max_seq is already
    # found correctly above and this branch is never reached unless a tab
    # is genuinely brand new/empty, where seq=1 is correct anyway).
    if max_seq == 0 and rows:
        max_seq = len(rows)

    seq = max_seq + 1
    return numbering["format"].format(
        prefix=numbering["prefix"], financial_year=numbering["financial_year"],
        seq=seq, padding=numbering["padding"],
    )


def _existing_base_amount(row_dict):
    """The row's pre-tax base amount, for comparing like-with-like against
    an incoming CP email's invoice_value.amount (which is ALWAYS pre-tax).

    USA/UK/Poland/India tabs don't have their own 'base_amount' column
    (only India/Kar Ventures-style tabs do) — they store 'total' as the
    tax-INCLUSIVE amount, plus separate tax columns. A real bug lived here:
    falling back to comparing the incoming pre-tax amount directly against
    the stored tax-INCLUSIVE total meant a duplicate submission on any
    entity with a non-zero tax rate (confirmed live on UK, VAT 20%) was
    silently never caught — 5000 base + 1000 VAT = 6000 total never equals
    the resubmitted email's base of 5000, so find_duplicate() never
    matched, and a second row/invoice number got created for the same
    request. Fixed by backing out every tax column (vat/cgst/sgst/igst)
    from the stored total before comparing, matching what
    draft_email_from_excel_row.py's _base_amount_for_intl() already does
    for the same reason on the read side."""
    if row_dict.get("base_amount") is not None:
        try:
            return float(row_dict["base_amount"])
        except (TypeError, ValueError):
            return -1
    try:
        total = float(row_dict.get("total") or 0)
    except (TypeError, ValueError):
        return -1
    tax_total = 0.0
    for tax_key in ("vat", "cgst", "sgst", "igst"):
        try:
            tax_total += float(row_dict.get(tax_key) or 0)
        except (TypeError, ValueError):
            pass
    return total - tax_total


def find_duplicate(ws, schema, data):
    """Match on the configured fields (default: pf_id + base_amount) —
    same rule for every tab, just scoped to that tab's own existing rows."""
    match_fields = schema["duplicate_check"]["match_on"]
    candidate = {}
    if "pf_id" in match_fields:
        candidate["pf_id"] = data.get("pf_id")
    if "base_amount" in match_fields:
        candidate["base_amount"] = (data.get("invoice_value") or {}).get("amount")

    for r in existing_rows(ws, schema):
        matched = True
        for field in match_fields:
            if field == "pf_id":
                if (r.get("pf_id") or "").strip() != (candidate.get("pf_id") or "").strip():
                    matched = False
                    break
            elif field == "base_amount":
                existing_amt = _existing_base_amount(r)
                try:
                    candidate_amt = float(candidate["base_amount"]) if candidate.get("base_amount") is not None else -2
                except (TypeError, ValueError):
                    candidate_amt = -2
                if existing_amt != candidate_amt:
                    matched = False
                    break
        if matched:
            return r
    return None


def compute_tax(base_amount, schema):
    """Generalized tax calc — dispatches on schema['tax']['type']:
      "vat"  — UK (20%), Poland (0% by default)
      "none" — USA
    ("cgst_sgst"/"igst" also supported here for a future India tab, but no
    active schema uses them right now — foreign clients only, no GST.)
    Always returns the full set of components; a tab's build_row() only
    writes whichever of these it actually has a column for."""
    tax = schema.get("tax", {"type": "none"})
    ttype = tax.get("type", "none")
    cgst = sgst = igst = vat = 0.0

    if ttype == "cgst_sgst":
        cgst = round(base_amount * tax["cgst_rate"], 2)
        sgst = round(base_amount * tax["sgst_rate"], 2)
    elif ttype == "igst":
        igst = round(base_amount * tax["igst_rate"], 2)
    elif ttype == "vat":
        vat = round(base_amount * tax["rate"], 2)
    # "none": all stay 0.0

    raw_total = base_amount + cgst + sgst + igst + vat
    total = round(raw_total, 2)
    round_off = round(total - raw_total, 2)
    return {"cgst": cgst, "sgst": sgst, "igst": igst, "vat": vat, "round_off": round_off, "total": total}


def build_row(data, schema, requested_by, invoice_date, sl_no, invoice_no, mis_result=None):
    base_amount = (data.get("invoice_value") or {}).get("amount")
    if base_amount is None:
        raise ValueError("invoice_value.amount is missing — cannot compute the row without a numeric amount")
    tax = compute_tax(base_amount, schema)

    # Canonical pool of every field ANY tab might have a column for — each
    # tab's schema only pulls out the keys it actually has, via
    # `values.get(c["key"], "")` below, so unused keys are simply dropped.
    values = {
        "sl_no": sl_no if sl_no is not None else "",
        "requested_by": requested_by or "",
        "client_type": "",
        "hsn": "",
        # Was "business_mode" (missing the 'l') -- every tab's schema column
        # key is "business_model", so this never matched and the column
        # was always written blank by the pipeline regardless of what a
        # tab's real Business Model data held (accountants have been
        # filling it in by hand directly in Excel instead). Found
        # 2026-08-14 while unifying UK/Poland's columns to match USA's
        # layout. Fixed the key; still blank by default since nothing
        # upstream (email parse or MIS) actually supplies this value yet.
        "business_model": "",
        "gstin": "",
        "client_company": data.get("client_company") or "",
        "client_contact_person": data.get("client_contact_person") or "",
        "client_address": data.get("client_address") or "",
        "po_no": data.get("po_no") or "",
        "po_date": data.get("po_date") or "",
        "pf_id": data.get("pf_id") or "",
        "master_project_id": data.get("master_project_id") or "",
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "invoice_description": data.get("invoice_description") or "",
        "work_order": data.get("work_order") or "",
        "pos": "",
        "state": "",
        "department": "",
        "new_old": "",
        "country": "",
        # Was reading (data.get("invoice_value") or {}).get("currency") -- that
        # nested key never existed (_parse_currency() only ever returns
        # {"raw", "amount"}), so this silently always fell through to the tab
        # default before 2026-08-13. Now reads the new top-level "currency"
        # field (parse_invoice_summary.py's "Currency" row, e.g. "INR"/"Dollar"
        # -> normalized to a 3-letter code) if the email included it, still
        # falling back to the tab's own default_currency if it didn't.
        "currency": data.get("currency") or schema.get("default_currency", "INR"),
        "base_amount": base_amount,
        "cgst": tax["cgst"],
        "sgst": tax["sgst"],
        "igst": tax["igst"],
        "vat": tax["vat"],
        "round_off": tax["round_off"],
        "total": tax["total"],
        "amount_to_transfer": "",
        "total_order_value": (data.get("total_order_value") or {}).get("amount") or "",
        "client_mail_to": ", ".join(data.get("client_mail_to") or []),
        "int_cc_mail": ", ".join(data.get("int_cc_mail") or []),
        "review_status": REVIEW_STATUS_PENDING,
        "email_drafted": False,
        MIS_VERIFIED_KEY: bool(mis_result and mis_result.get("verified")),
        "email_status": "Not Sent",
        "email_sent_date": "",
        "email_sent_to": "",
        "email_sent_cc": "",
        "payment_status": "Not Paid",
        "transaction_id": "",
        "transaction_id_2": "",
        "received_amount": "",
        "payment_date": "",
        "tds": "",
        "tds_on_igst": "",
        "due_date": _compute_due_date(invoice_date),
        "remarks": "",
    }
    # Any real fields the MIS API returned (per MIS_FIELD_MAP in mis_api.py)
    # get merged in here -- only ones a tab actually has a column for end up
    # written, same rule as every other key in `values` above.
    if mis_result and mis_result.get("fields"):
        values.update(mis_result["fields"])
    return [values.get(c["key"], "") for c in schema["columns"]], values


def append_invoice(xlsx_path, data, requested_by=None, invoice_date=None, dry_run=False,
                    out_path=None, entity_key=None, debug_dir=None):
    resolution = "given" if entity_key else None
    if entity_key is None:
        entity_key, resolution = resolve_entity_for_company(data.get("client_company"), xlsx_path)

    if entity_key is None:
        return {
            "status": "entity_unresolved",
            "message": (
                f"Could not determine which tab '{data.get('client_company')!r}' belongs to. "
                "Not appended anywhere — add this company to config/company_tab_map.json "
                "(or to an existing tab's Company Name column) and re-run."
            ),
            "client_company": data.get("client_company"),
        }

    schema = load_schema(entity_key)
    wb = load_workbook_with_retry(xlsx_path)
    ws = ensure_sheet(wb, schema)

    # Duplicate check DISABLED 2026-08-13, per explicit instruction: the PF ID
    # is the same for a project every month, and the invoice amount is also
    # usually the same each month for a given project -- so matching on
    # (pf_id, base_amount) flagged every legitimate month-2/month-3/etc.
    # resubmission for the same project as a "duplicate" and refused to
    # append it. Every incoming matching email now always gets appended as
    # a new row + new invoice number, with no automatic duplicate check at
    # all -- including a true accidental resend (e.g. the same CP email
    # forwarded twice) that the old check *did* correctly catch live on
    # 2026-08-13. That safety net is intentionally gone now; catching an
    # actual accidental double-send is the accountant's job during manual
    # review (Review Status), not the pipeline's. find_duplicate() itself is
    # left in place, just unused, in case this ever needs to come back.
    #
    # dup = find_duplicate(ws, schema, data)
    # if dup:
    #     return {
    #         "status": "duplicate_flagged",
    #         "message": (
    #             f"Matches existing row in '{schema['sheet_name']}': PF ID={dup.get('pf_id')!r}, "
    #             f"Invoice No={dup.get('invoice_no')!r}. Not appended."
    #         ),
    #         "existing_row": dup,
    #         "entity_key": entity_key,
    #     }

    # MIS project-info check -- happens right here, AFTER the duplicate
    # check (no point calling the API for a row we're not going to append)
    # and BEFORE the row is built, so the whole row (email fields + any
    # MIS-sourced fields) goes into Excel in one shot. Never blocks the
    # append -- see mis_api.py's fetch_project_info docstring.
    debug_path = (Path(debug_dir) / "debug_last_mis_response.json") if debug_dir else None
    mis_result = fetch_project_info(data.get("pf_id"), debug_path=debug_path)

    sl_no = next_sl_no(ws, schema)
    invoice_no = next_invoice_no(ws, schema)
    invoice_date = invoice_date or date.today().isoformat()

    row_values, row_dict = build_row(
        data, schema, requested_by, invoice_date, sl_no, invoice_no, mis_result=mis_result,
    )

    if not dry_run:
        target_row = next_data_row(ws, schema)
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=target_row, column=col_idx, value=value)
        wb.save(out_path or xlsx_path)

    return {
        "status": "appended", "row": row_dict, "sheet": schema["sheet_name"],
        "entity_key": entity_key, "resolution": resolution,
        "mis_verification": {"verified": mis_result.get("verified"), "error": mis_result.get("error")},
    }


def find_row_index(ws, schema, invoice_no):
    """Return the actual worksheet row number (not a dict) matching invoice_no, or None."""
    cols = schema["columns"]
    inv_col = next(i for i, c in enumerate(cols) if c["key"] == "invoice_no") + 1
    start = header_row_index(schema) + 1
    for row_idx in range(start, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=inv_col).value
        if (val or "").strip() == invoice_no.strip():
            return row_idx
    return None


def _is_truthy_flag(value):
    """Excel booleans round-trip cleanly through openpyxl as Python True/False
    when written natively (which mark_email_drafted below does) — but treat
    a stray "TRUE"/"1" text value as truthy too, in case someone types it by
    hand instead of using a real checkbox/boolean cell."""
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in ("true", "1", "yes")


def _is_false_flag(value):
    """Return True only for an explicitly false checkbox/boolean value.

    A blank cell is deliberately *not* eligible: historical tracker rows
    can be blank because the Email Drafted column was added later, and they
    must never create drafts merely because the automation is started.
    """
    if isinstance(value, bool):
        return not value
    return str(value).strip().lower() in ("false", "0", "no")


def find_rows_ready_for_invoicing(xlsx_path, entity_key):
    """Return rows whose Email Drafted flag is not True.

    The invoice-generation pass is driven exclusively by Email Drafted:
    False, "FALSE", "0", or "No" means the row is ready; blank cells and
    True, "TRUE", "1", or "Yes" are skipped. Review Status is intentionally
    not used as a gate, so the workflow matches the Google Sheet automation
    requirement exactly.

    Returns a list of row dicts (same shape as append_invoice's row_dict),
    read straight back out of the sheet — this is the ONLY way the
    invoice-generation pass gets its data once it's decoupled from the
    email-read step that originally appended the row.
    """
    schema = load_schema(entity_key)
    if not _has_column(schema, "email_drafted"):
        return []

    wb = load_workbook_with_retry(xlsx_path, data_only=True)
    sheet_name = schema["sheet_name"]
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]

    ready = []
    for r in existing_rows(ws, schema):
        if not _is_false_flag(r.get("email_drafted")):
            continue
        ready.append(r)
    return ready


def mark_email_drafted(xlsx_path, invoice_no, out_path=None, entity_key="usa"):
    """Writes True to the Email Drafted cell for one row by Invoice No —
    called once the invoice PDF + Gmail draft for that row have actually
    been created, so it's never picked up twice. Deliberately does NOT
    touch Review Status — that column stays exactly as the human set it,
    a permanent record separate from the pipeline's own progress."""
    schema = load_schema(entity_key)
    if not _has_column(schema, "email_drafted"):
        return {
            "status": "not_supported",
            "message": f"Tab '{schema['sheet_name']}' has no Email Drafted column — skipping.",
            "invoice_no": invoice_no,
        }

    wb = load_workbook_with_retry(xlsx_path)
    ws = wb[schema["sheet_name"]]

    row_idx = find_row_index(ws, schema, invoice_no)
    if row_idx is None:
        return {"status": "not_found", "invoice_no": invoice_no}

    col_idx = {c["key"]: i + 1 for i, c in enumerate(schema["columns"])}
    ws.cell(row=row_idx, column=col_idx["email_drafted"], value=True)
    wb.save(out_path or xlsx_path)
    return {"status": "updated", "invoice_no": invoice_no, "email_drafted": True,
            "row": row_idx, "sheet": schema["sheet_name"]}


def update_email_status(xlsx_path, invoice_no, status, sent_date=None, sent_to=None, sent_cc=None,
                         out_path=None, entity_key="usa"):
    """
    Called AFTER the email draft (or send) happens. None of USA/UK/Poland
    2026 have EMAIL STATUS / SENT DATE / SENT TO / SENT CC columns in your
    real tracker, so this is always a safe no-op that says so, rather than
    an error — kept generic in case a tracked tab is added later.
    """
    schema = load_schema(entity_key)
    if not _has_column(schema, "email_status"):
        return {
            "status": "not_supported",
            "message": f"Tab '{schema['sheet_name']}' has no EMAIL STATUS column — skipping status write-back.",
            "invoice_no": invoice_no,
        }

    wb = load_workbook_with_retry(xlsx_path)
    ws = wb[schema["sheet_name"]]

    row_idx = find_row_index(ws, schema, invoice_no)
    if row_idx is None:
        return {"status": "not_found", "invoice_no": invoice_no}

    col_idx = {c["key"]: i + 1 for i, c in enumerate(schema["columns"])}

    def _join(v):
        return ", ".join(v) if isinstance(v, (list, tuple)) else (v or "")

    ws.cell(row=row_idx, column=col_idx["email_status"], value=status)
    if sent_date is not None and "email_sent_date" in col_idx:
        ws.cell(row=row_idx, column=col_idx["email_sent_date"], value=sent_date)
    if sent_to is not None and "email_sent_to" in col_idx:
        ws.cell(row=row_idx, column=col_idx["email_sent_to"], value=_join(sent_to))
    if sent_cc is not None and "email_sent_cc" in col_idx:
        ws.cell(row=row_idx, column=col_idx["email_sent_cc"], value=_join(sent_cc))

    wb.save(out_path or xlsx_path)
    return {
        "status": "updated", "invoice_no": invoice_no, "email_status": status,
        "row": row_idx, "sheet": schema["sheet_name"],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("parsed_data_json")
    ap.add_argument("--entity", default=None, choices=list(ENTITY_SCHEMA_PATHS.keys()),
                     help="Skip auto-detection and force a specific tab")
    ap.add_argument("--requested-by", default="")
    ap.add_argument("--invoice-date", default=None)
    ap.add_argument("--out", default=None, help="Write to a different file instead of overwriting the input")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    with open(args.parsed_data_json, "r", encoding="utf-8") as f:
        payload = json.load(f)
    data = payload.get("data", payload)

    result = append_invoice(
        args.xlsx_path, data,
        requested_by=args.requested_by,
        invoice_date=args.invoice_date,
        dry_run=args.dry_run,
        out_path=args.out,
        entity_key=args.entity,
    )
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))


if __name__ == "__main__":
    main()
