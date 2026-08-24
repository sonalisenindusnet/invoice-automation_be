"""
create_singapore_tab.py

One-off script: creates the "Singapore" worksheet in the live Invoice
Tracker Google Sheet and writes its header row, using the same
config/tabs/singapore.json schema the rest of the app already reads.

Run this ONCE, from your machine (this can't be run from the cloud sandbox
-- it has no network path to Google's APIs):

    cd invoice-automation_be/src
    python scripts/create_singapore_tab.py

Idempotent: if a "Singapore" tab already exists, it does nothing and just
reports that. Safe to re-run.
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
SCHEMA_PATH = PROJECT_ROOT / "config" / "tabs" / "singapore.json"


def main():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        schema = json.load(f)

    tracker_ref = tracker_ref_from_config(cfg, CONFIG_PATH.parent)
    wb = load_tracker_with_retry(tracker_ref)

    sheet_name = schema["sheet_name"]
    if sheet_name in wb.sheetnames:
        print(f"'{sheet_name}' tab already exists -- nothing to do.")
        return

    print(f"Creating '{sheet_name}' tab...")
    ws = wb.create_sheet(sheet_name)

    header_row_idx = schema.get("header_row_index", 1)
    for col_idx, column in enumerate(schema["columns"], start=1):
        ws.cell(row=header_row_idx, column=col_idx, value=column["header"])

    wb.save()
    print(f"Done. '{sheet_name}' tab created with {len(schema['columns'])} header columns.")


if __name__ == "__main__":
    main()
