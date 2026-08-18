"""
mark_email_sent.py

Writes the outcome of the email step back into the tracker, so it stays
the single source of truth for "has this invoice actually gone out, and
to whom" — not just "was a row created for it".

Usage:
    python mark_email_sent.py <tracker.xlsx> --invoice-no INT/USA/26-27/001 \
        --entity usa --status "Draft Created" --to "accounts@client.com" \
        --cc "accountsint@intglobal.com" --date 2026-08-07

Note: none of USA/UK/Poland 2026 have EMAIL STATUS / SENT DATE / TO / CC
columns in your real tracker — this always returns status "not_supported"
rather than an error (those tabs simply don't track this today).
"""
import argparse
import json
import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from excel.append_invoice_to_excel import update_email_status


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("--entity", default="usa", choices=["usa", "uk", "poland"])
    ap.add_argument("--invoice-no", required=True)
    ap.add_argument("--status", required=True, help='e.g. "Draft Created", "Sent", "Send Failed"')
    ap.add_argument("--date", default=None, help="Defaults to leaving the date column untouched if omitted")
    ap.add_argument("--to", default=None, help="Comma-separated addresses actually used")
    ap.add_argument("--cc", default=None, help="Comma-separated addresses actually used")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    to_list = [e.strip() for e in args.to.split(",")] if args.to else None
    cc_list = [e.strip() for e in args.cc.split(",")] if args.cc else None

    result = update_email_status(
        args.xlsx_path, args.invoice_no, args.status,
        sent_date=args.date, sent_to=to_list, sent_cc=cc_list, out_path=args.out,
        entity_key=args.entity,
    )
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))


if __name__ == "__main__":
    main()
