"""
excel_writer.py

Maps a validated invoice request into a new row in the real tracker
workbook, in whichever tab the request's "entity" (usa/uk/poland) belongs
to. Assigns the next invoice number in that tab's own numbering series and
stamps a "Created At" timestamp. No tax calculation or other derived
columns -- everything else the tab has a column for is left blank.

Each entity's real tab layout (sheet name, columns, invoice-number series)
is described in config/tabs/<entity>.json.

The sheet is also edited by hand by a human accountant, so two things are
handled carefully:
  - The next invoice number is computed from the workbook's own state at
    save time (never cached across requests), by scanning for the highest
    existing number in that tab's series -- so a number a human has since
    typed in is respected, not overwritten or skipped.
  - `TRACKER_LOCK` (shared from utils/xlsx_io.py with the draft-mailer
    poller, which also reads-then-writes this same file) serializes
    save_invoice() calls within this process, so two writers arriving at
    nearly the same moment can't both compute state from a stale read
    before either has saved. This does not protect against a human
    editing the file in Excel at the exact same instant --
    load_workbook_with_retry already retries through the brief file-lock
    window that causes, but a true simultaneous write from Excel itself is
    outside what this process can arbitrate.
"""
import json
from datetime import date, datetime
from pathlib import Path

from utils.xlsx_io import load_workbook_with_retry, TRACKER_LOCK

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
TABS_DIR = PROJECT_ROOT / "config" / "tabs"

CREATED_AT_HEADER = "Created At"


class UnknownEntityError(ValueError):
    """Raised when there's no schema for the requested entity key."""


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


def build_row(data, schema, invoice_no, invoice_date, requested_by):
    """Only fields that actually came in on the request (plus the assigned
    invoice number) get a value; every other column this tab has is left
    blank."""
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
    }
    row_values = [values.get(c["key"], "") for c in schema["columns"]]
    return row_values, values


def save_invoice(data, entity_key, tracker_path, requested_by=None):
    """Appends one row for `data` into `entity_key`'s real tab of the
    tracker at `tracker_path`, assigning the next invoice number in that
    tab's series and a Created At timestamp. Returns {"sheet": ...,
    "row": <written field values, as a dict>}."""
    with TRACKER_LOCK:
        schema = load_schema(entity_key)
        wb = load_workbook_with_retry(str(tracker_path))
        ws = wb[schema["sheet_name"]]
        header_row_idx = schema.get("header_row_index", 1)

        invoice_no = _next_invoice_no(ws, schema)
        invoice_date = date.today().isoformat()
        created_at = datetime.now().isoformat(timespec="seconds")

        row_values, row_dict = build_row(data, schema, invoice_no, invoice_date, requested_by)
        row_dict["created_at"] = created_at

        target_row = _next_data_row(ws, schema)
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=target_row, column=col_idx, value=value)

        created_at_col = _created_at_column(ws, header_row_idx)
        ws.cell(row=target_row, column=created_at_col, value=created_at)

        wb.save(str(tracker_path))

    return {"sheet": schema["sheet_name"], "row": row_dict}
