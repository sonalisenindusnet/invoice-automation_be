"""
check_singapore_header.py

Read-only sanity check: prints the LIVE "Singapore" tab's actual header row,
and compares it column-by-column against config/tabs/singapore.json. Makes
NO changes to the sheet -- safe to run any time.

Run this from your machine (the cloud sandbox has no network path to
Google's APIs):

    cd invoice-automation_be
    python scripts/check_singapore_header.py

Paste the output back so the schema can be corrected if anything doesn't
line up.
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
SHEET_NAME = "Singapore"


def main():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        schema = json.load(f)

    tracker_ref = tracker_ref_from_config(cfg, CONFIG_PATH.parent)
    wb = load_tracker_with_retry(tracker_ref)

    if SHEET_NAME not in wb.sheetnames:
        print(f"'{SHEET_NAME}' tab doesn't exist.")
        return

    ws = wb[SHEET_NAME]
    header_row_idx = schema.get("header_row_index", 1)
    live_header = list(
        list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=True))[0]
    )
    # Trim trailing Nones/blanks for a clean printout.
    while live_header and live_header[-1] in (None, ""):
        live_header.pop()

    expected_header = [c["header"] for c in schema["columns"]]

    print(f"Live header  ({len(live_header)} columns): {live_header}")
    print(f"Schema header ({len(expected_header)} columns): {expected_header}")
    print()

    mismatches = []
    max_len = max(len(live_header), len(expected_header))
    for i in range(max_len):
        live_val = live_header[i] if i < len(live_header) else "<missing>"
        exp_val = expected_header[i] if i < len(expected_header) else "<missing>"
        if live_val != exp_val:
            mismatches.append((i + 1, live_val, exp_val))

    if not mismatches:
        print("MATCH -- schema lines up with the live header exactly, column for column.")
    else:
        print(f"MISMATCH at {len(mismatches)} column(s):")
        for col_num, live_val, exp_val in mismatches:
            print(f"  column {col_num}: live={live_val!r}  schema={exp_val!r}")
        print()
        print("Paste this output back so config/tabs/singapore.json can be corrected "
              "before any Singapore invoice is saved.")


if __name__ == "__main__":
    main()
