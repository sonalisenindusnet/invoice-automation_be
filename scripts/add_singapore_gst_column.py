"""
add_singapore_gst_column.py

One-off, SAFE column-insert migration for the live "Singapore" tab.

Why this exists: config/tabs/singapore.json briefly included a "GST (SGD)"
tax-amount column on the assumption the live "Singapore" tab hadn't been
created yet. That assumption was wrong -- create_singapore_tab.py reported
"tab already exists" -- so the live header row still has the ORIGINAL
24-column layout with no GST column. The schema was immediately reverted
to match (see config/tabs/singapore.json's own _tax_note) to stop any risk
of a save writing the tax amount into whatever real column happens to sit
where "vat" would have landed (originally "Client Mail To"), which would
have silently shifted every column after it for every future row.

This script performs the actual physical insert on the LIVE sheet instead:
it inserts a new "GST (SGD)" column right after "Total Amount (SGD)",
shifting every column from that point rightward by one -- for the header
row AND every real data row (if any exist), so nothing is lost, blank
rows are skipped safely, and a row's own genuine trailing columns (e.g. a
dynamically-added "Created At") shift correctly too since the whole row's
actual width is read and shifted, not just the 24 columns the schema
currently lists.

Run this ONCE, from your machine (this can't be run from the cloud sandbox
-- it has no network path to Google's APIs):

    cd invoice-automation_be
    python scripts/add_singapore_gst_column.py

Idempotent: if "GST (SGD)" already exists as a header, it does nothing and
reports where. Prints the header row before and after, and how many rows
it touched, so you can see exactly what happened before trusting it with
real invoices.

AFTER this succeeds: send me the printed output (or just say it worked) --
I'll update config/tabs/singapore.json to add the "vat" column back at the
matching position and re-verify before any Singapore invoice is saved
against the new layout.
"""
import json
import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from utils.tracker_io import load_tracker_with_retry, tracker_ref_from_config  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "save_api_config.json"

SHEET_NAME = "Singapore"
INSERT_AFTER_HEADER = "Total Amount (SGD)"
NEW_HEADER = "GST (SGD)"


def main():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    tracker_ref = tracker_ref_from_config(cfg, CONFIG_PATH.parent)
    wb = load_tracker_with_retry(tracker_ref)

    if SHEET_NAME not in wb.sheetnames:
        print(f"'{SHEET_NAME}' tab doesn't exist -- nothing to migrate.")
        return

    ws = wb[SHEET_NAME]
    max_row = ws.max_row
    if max_row == 0:
        print(f"'{SHEET_NAME}' tab is completely empty -- nothing to migrate.")
        return

    rows = list(ws.iter_rows(min_row=1, max_row=max_row, values_only=True))
    header_row = list(rows[0]) if rows[0] is not None else []

    if NEW_HEADER in header_row:
        print(f"'{NEW_HEADER}' already exists at column {header_row.index(NEW_HEADER) + 1} "
              f"-- nothing to do.")
        return

    if INSERT_AFTER_HEADER not in header_row:
        print(f"ERROR: couldn't find '{INSERT_AFTER_HEADER}' in the header row -- "
              f"aborting without changing anything.")
        print(f"Header row was: {header_row}")
        return

    insert_at = header_row.index(INSERT_AFTER_HEADER) + 2  # 1-based column, right after it

    print(f"Header row (before, {len(header_row)} columns): {header_row}")
    print(f"Inserting '{NEW_HEADER}' at column {insert_at} across {max_row} row(s) "
          f"(including header)...")

    touched = 0
    for row_idx, row in enumerate(rows, start=1):
        if row is None:
            continue  # a genuinely blank row -- nothing to shift, leave it alone
        row = list(row)
        # Shift every existing column from insert_at onward one column to
        # the right. All reads already happened above (`rows` is a plain
        # list in memory), so the order of these .cell() writes doesn't
        # matter -- each just stages one cell independently.
        for col in range(len(row), insert_at - 1, -1):
            value = row[col - 1] if col - 1 < len(row) else ""
            ws.cell(row=row_idx, column=col + 1, value=value)
        new_value = NEW_HEADER if row_idx == 1 else ""
        ws.cell(row=row_idx, column=insert_at, value=new_value)
        touched += 1

    wb.save()

    new_header_row = list(list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0])
    print(f"Header row (after,  {len(new_header_row)} columns): {new_header_row}")
    print(f"Done. Touched {touched} row(s). '{NEW_HEADER}' now at column {insert_at}.")
    print()
    print("NEXT STEP: let me know this succeeded (paste the output above if you like) "
          "so I can update config/tabs/singapore.json to match and re-verify before "
          "any Singapore invoice is saved.")


if __name__ == "__main__":
    main()
