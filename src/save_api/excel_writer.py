"""
excel_writer.py

Maps a validated invoice request into a new row in the real tracker
workbook, in whichever tab the request's "entity" (usa/uk/poland) belongs
to. Assigns the next invoice number in that tab's own numbering series and
stamps a "Created At" timestamp. A handful of other columns are also
auto-filled on every new row (Payment Status, Payment Due Date, Review
Status, Email Drafted, MIS Verified -- see build_row()); everything else
the tab has a column for is left blank.

Each entity's real tab layout (sheet name, columns, invoice-number series)
is described in config/tabs/<entity>.json.

The real tracker lives on a live Google Sheet (not a local file) --
reached via utils/tracker_io.py's gspread-backed adapter, which duck-types
enough of openpyxl's Workbook/Worksheet interface that the row/column logic
below doesn't need to know it's talking to Sheets. Local-.xlsx tracker
support has been removed entirely.

The sheet is also edited by hand by a human accountant, so two things are
handled carefully:
  - The next invoice number is computed from the workbook's own state at
    save time (never cached across requests), by scanning for the highest
    existing number in that tab's series -- so a number a human has since
    typed in is respected, not overwritten or skipped.
  - `TRACKER_LOCK` (shared from utils/xlsx_io.py with the draft-mailer
    poller, which also reads-then-writes this same tracker) serializes
    save_invoice() calls within this process, so two writers arriving at
    nearly the same moment can't both compute state from a stale read
    before either has saved. Google Sheets itself already supports true
    concurrent multi-writer access with no file locks -- this lock is only
    about this process's own two features never racing each other.
"""
import json
import os
from datetime import date, datetime, timedelta
from pathlib import Path

from utils.xlsx_io import TRACKER_LOCK
from utils.tracker_io import load_tracker_with_retry, save_tracker

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
TABS_DIR = PROJECT_ROOT / "config" / "tabs"

CREATED_AT_HEADER = "Created At"

# Auto-filled on every new row -- not derived from the request payload.
PAYMENT_STATUS_DEFAULT = "Not Paid"
REVIEW_STATUS_DEFAULT = "Pending Review"
PAYMENT_DUE_DAYS_ENV = "PAYMENT_DUE_DAYS"
DEFAULT_PAYMENT_DUE_DAYS = 7


def _payment_due_days():
    """Number of days after the invoice date that payment is due, read from
    the PAYMENT_DUE_DAYS environment variable. Falls back to
    DEFAULT_PAYMENT_DUE_DAYS if unset or not a valid integer."""
    raw = os.environ.get(PAYMENT_DUE_DAYS_ENV)
    if raw is None or not raw.strip():
        return DEFAULT_PAYMENT_DUE_DAYS
    try:
        return int(raw.strip())
    except ValueError:
        return DEFAULT_PAYMENT_DUE_DAYS


class UnknownEntityError(ValueError):
    """Raised when there's no schema for the requested entity key."""


class RowNotFoundError(ValueError):
    """Raised when no row in the tab matches the given invoice number."""


class PfIdMismatchError(ValueError):
    """Raised when the row found by invoice number belongs to a different
    PF ID than the one in the request -- a safety check against flipping
    the wrong project's row on a typo'd invoice number."""


def load_schema(entity_key):
    path = TABS_DIR / f"{entity_key}.json"
    if not path.exists():
        raise UnknownEntityError(f"No schema configured for entity '{entity_key}'")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _existing_rows(ws, schema):
    """Yields {key: value} for each row that actually has data."""
    cols = schema["columns"]
    start = schema.get("header_row_index", 1) + 1
    for row in ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        yield {cols[i]["key"]: row[i] if i < len(row) else None for i in range(len(cols))}


def _next_data_row(ws, schema):
    """Row index to write a new row at -- right after the last row that
    actually has data, not just ws.max_row (which also counts rows that
    are pre-formatted but otherwise empty, common in this hand-built
    tracker)."""
    cols = schema["columns"]
    start = schema.get("header_row_index", 1) + 1
    last_data_row = start - 1
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True), start=start
    ):
        if row is not None and not all(v is None for v in row[:len(cols)]):
            last_data_row = row_idx
    return last_data_row + 1


def _next_invoice_no(ws, schema):
    """Next number in this tab's own series: highest existing sequence
    number for the configured prefix/financial-year, plus one. Only rows
    whose Invoice No. actually matches the series format count -- a
    human's typo or a blank cell is simply ignored, never treated as a
    conflict."""
    numbering = schema["invoice_numbering"]
    prefix = f"{numbering['prefix']}/{numbering['financial_year']}/"
    max_seq = 0
    rows = list(_existing_rows(ws, schema))
    for r in rows:
        inv = r.get("invoice_no")
        if isinstance(inv, str) and inv.startswith(prefix):
            tail = inv[len(prefix):]
            if tail.isdigit():
                max_seq = max(max_seq, int(tail))
    if max_seq == 0 and rows:
        # A formula-driven Invoice No. column (no cached value) would parse
        # as 0 above despite real existing rows -- fall back to the row
        # count so a new number is never handed out on top of one already
        # in use.
        max_seq = len(rows)

    seq = max_seq + 1
    return numbering["format"].format(
        prefix=numbering["prefix"], financial_year=numbering["financial_year"],
        seq=seq, padding=numbering["padding"],
    )


def _created_at_column(ws, header_row_idx):
    """Returns the column index of the "Created At" header, adding it
    right after the last non-empty header cell if it isn't there yet.
    Never touches or moves any existing column."""
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=header_row_idx, column=col_idx).value
        if isinstance(value, str) and value.strip().lower() == CREATED_AT_HEADER.lower():
            return col_idx

    last_used_col = 0
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=header_row_idx, column=col_idx).value not in (None, ""):
            last_used_col = col_idx
    new_col = last_used_col + 1
    ws.cell(row=header_row_idx, column=new_col, value=CREATED_AT_HEADER)
    return new_col


def _find_row_by_invoice_no(ws, schema, invoice_no):
    """Returns (row_idx, row_dict) for the row whose Invoice No. matches
    `invoice_no` exactly (after stripping whitespace), or (None, None) if
    no row matches."""
    target = (invoice_no or "").strip()
    for row_idx, row in enumerate(
        ws.iter_rows(
            min_row=schema.get("header_row_index", 1) + 1, max_row=ws.max_row, values_only=True,
        ),
        start=schema.get("header_row_index", 1) + 1,
    ):
        if row is None:
            continue
        row_dict = {schema["columns"][i]["key"]: row[i] for i in range(len(schema["columns"])) if i < len(row)}
        inv = row_dict.get("invoice_no")
        if isinstance(inv, str) and inv.strip() == target:
            return row_idx, row_dict
    return None, None


def update_mis_verified(entity_key, invoice_no, pf_id, mis_verified, tracker_ref):
    """Finds `invoice_no` in `entity_key`'s real tab and sets its MIS
    Verified column to `mis_verified` (True/False). Raises RowNotFoundError
    if no row matches, or PfIdMismatchError if the row's own PF ID doesn't
    match `pf_id` -- a safety check before flipping anything. Returns
    {"sheet": ..., "invoice_no": ..., "mis_verification_done": ...}."""
    with TRACKER_LOCK:
        schema = load_schema(entity_key)
        wb = load_tracker_with_retry(tracker_ref)
        ws = wb[schema["sheet_name"]]

        row_idx, row_dict = _find_row_by_invoice_no(ws, schema, invoice_no)
        if row_idx is None:
            raise RowNotFoundError(f"Invoice No {invoice_no!r} not found in '{schema['sheet_name']}' tab")

        existing_pf_id = (row_dict.get("pf_id") or "")
        if pf_id and str(existing_pf_id).strip() != str(pf_id).strip():
            raise PfIdMismatchError(
                f"PF ID mismatch for {invoice_no!r}: request said {pf_id!r}, row has {existing_pf_id!r}"
            )

        col_idx = next(i for i, c in enumerate(schema["columns"]) if c["key"] == "mis_verification_done") + 1
        ws.cell(row=row_idx, column=col_idx, value=bool(mis_verified))
        save_tracker(wb, tracker_ref)

    return {
        "sheet": schema["sheet_name"],
        "invoice_no": invoice_no,
        "mis_verification_done": bool(mis_verified),
    }


def build_row(data, schema, invoice_no, invoice_date, requested_by, due_date):
    """Fields that came in on the request (plus the assigned invoice
    number) get their value from the request; five more columns are
    auto-filled on every new row regardless of what the request contains --
    Payment Status ("Not Paid"), Payment Due Date (`due_date`, computed by
    the caller as invoice_date + PAYMENT_DUE_DAYS), Review Status
    ("Pending Review", so the draft-mailer poller never picks up a row
    until a human changes it to "Reviewed"), Email Drafted (False, so the
    poller doesn't mistake a fresh row for one it already drafted), and
    MIS Verified (False, until the MIS-verification API flips it). Every
    other column this tab has is left blank."""
    values = {
        "invoice_date": invoice_date,
        "invoice_no": invoice_no,
        "requested_by": requested_by or "",
        "client_company": data.get("client_company") or "",
        "client_contact_person": data.get("client_contact_person") or "",
        "invoice_description": data.get("invoice_description") or "",
        "work_order": data.get("work_order") or "",
        "pf_id": data.get("pf_id") or "",
        "master_project_id": data.get("master_project_id") or "",
        "currency": data.get("currency") or "",
        "total": (data.get("invoice_value") or {}).get("amount"),
        "client_mail_to": ", ".join(data.get("client_mail_to") or []),
        "int_cc_mail": ", ".join(data.get("int_cc_mail") or []),
        "payment_status": PAYMENT_STATUS_DEFAULT,
        "due_date": due_date,
        "review_status": REVIEW_STATUS_DEFAULT,
        "email_drafted": False,
        "mis_verification_done": False,
    }
    row_values = [values.get(c["key"], "") for c in schema["columns"]]
    return row_values, values


def save_invoice(data, entity_key, tracker_ref, requested_by=None):
    """Appends one row for `data` into `entity_key`'s real tab of the live
    tracker at `tracker_ref` (see utils.tracker_io.tracker_ref_from_config),
    assigning the next invoice number in that tab's series, a Created At
    timestamp, and the auto-filled Payment Status/Payment Due Date/Review
    Status/Email Drafted/MIS Verified fields (see build_row()). Returns
    {"sheet": ..., "row": <written field values, as a dict>}."""
    with TRACKER_LOCK:
        schema = load_schema(entity_key)
        wb = load_tracker_with_retry(tracker_ref)
        ws = wb[schema["sheet_name"]]
        header_row_idx = schema.get("header_row_index", 1)

        invoice_no = _next_invoice_no(ws, schema)
        invoice_date = date.today().isoformat()
        due_date = (date.today() + timedelta(days=_payment_due_days())).isoformat()
        created_at = datetime.now().isoformat(timespec="seconds")

        row_values, row_dict = build_row(data, schema, invoice_no, invoice_date, requested_by, due_date)
        row_dict["created_at"] = created_at

        target_row = _next_data_row(ws, schema)
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=target_row, column=col_idx, value=value)

        created_at_col = _created_at_column(ws, header_row_idx)
        ws.cell(row=target_row, column=created_at_col, value=created_at)

        save_tracker(wb, tracker_ref)

    return {"sheet": schema["sheet_name"], "row": row_dict}
