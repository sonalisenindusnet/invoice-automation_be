"""
run_demo.py

Full demo on a COPY of your uploaded tracker (never touches the original
upload, and never touches your real tracker either):

  Stage 1 (deterministic, "normal process automation"):
    parse_invoice_summary.py  ->  entity_resolver.py (which tab?)  ->  append_invoice_to_excel.py
    (writes a new row into the correct tab — USA in this sample; foreign
     clients only, no GST/India right now)

  Stage 2 (reads the row just appended, drafts the email):
    draft_email_from_excel_row.py
    (invoice PDF, using that entity's own template + To/CC/Subject/Body;
     no tax lines for USA, VAT for UK/Poland)

  Stage 3 (writes the outcome back into the same row, where supported):
    mark_email_sent.py / update_email_status()
    (none of USA/UK/Poland 2026 have an EMAIL STATUS column today, so this
     always reports "not_supported" — flagged honestly, not hidden)

Run twice in a row to see the duplicate-check reject the second attempt.

Also runs a bonus scenario at the end: a completely unknown company name,
proving the resolver flags it instead of guessing a tab.
"""
import json
import sys
from datetime import date
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
SRC_DIR = BASE / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from parsing.parse_invoice_summary import parse_invoice_summary
from excel.append_invoice_to_excel import append_invoice, update_email_status
from mailer.draft_email_from_excel_row import compose_email, render_invoice_for_entity

OUTPUT_DIR = BASE / "output"
TRACKER_COPY = OUTPUT_DIR / "tracker_demo.xlsx"

# 'Somax Inc' is already listed in config/company_tab_map.json as a USA
# client — this is the main demo scenario.
SAMPLE_INVOICE_SUMMARY_TEXT = """\
1\tPf Id\t2505/Retail-Diverse/9999
2\tAccount Name\tDE_COC_FC
3\tClient Name\tSomax Inc
4\tInvoice Value\t14000
5\tInvoice Description\tMonthly retainer - August 2026
6\tClient Name\tJordan Lee
7\tClient Mail id (To and CC)\taccounts@somax-demo.com
8\tINT CC mail id\taccountsint@intglobal.com
9\tWork Order\tYes
10\tMaster Project ID\t
11\tTotal Order Value\t168000
"""

# A company that isn't in config/company_tab_map.json and doesn't already
# exist in any tab — proves the resolver flags it instead of guessing.
SAMPLE_UNKNOWN_TEXT = """\
1\tPf Id\t9999/Unknown/0001
2\tAccount Name\tDE_COC_FC
3\tClient Name\tTotally New Client Ltd
4\tInvoice Value\t5000
5\tInvoice Description\tFirst engagement
6\tClient Name\tSomeone
7\tClient Mail id (To and CC)\taccounts@totallynewclient-demo.com
8\tINT CC mail id\taccountsint@intglobal.com
9\tWork Order\tYes
10\tMaster Project ID\t
11\tTotal Order Value\t5000
"""


def _unresolved_bonus_scenario():
    print("\n\n=== Bonus: an unmapped company is flagged, never guessed ===")
    data, _ = parse_invoice_summary(SAMPLE_UNKNOWN_TEXT, source_type="text")
    result = append_invoice(TRACKER_COPY, data, requested_by="Demo CP", invoice_date=date(2026, 8, 7).isoformat())
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))
    if result["status"] == "entity_unresolved":
        print("-> Correctly flagged, NOT appended anywhere. This is what happens for a real "
              "unmapped client — add it to config/company_tab_map.json and it resolves next time.")


def main():
    print(f"=== Stage 1: parse email ===")
    data, warnings = parse_invoice_summary(SAMPLE_INVOICE_SUMMARY_TEXT, source_type="text")
    print(json.dumps(data, indent=2, ensure_ascii=False))
    for w in warnings:
        print(f"  WARNING: {w}")

    print(f"\n=== Stage 1: figure out which tab this client belongs to, then append ===")
    result = append_invoice(
        TRACKER_COPY, data,
        requested_by="Demo CP (from email sender)",
        invoice_date=date(2026, 8, 7).isoformat(),
    )
    print(json.dumps(result, indent=2, default=str, ensure_ascii=False))

    if result["status"] == "entity_unresolved":
        print("\nCould not determine which tab this client belongs to — nothing was appended. "
              "See config/company_tab_map.json.")
        return

    if result["status"] == "duplicate_flagged":
        print("\nRow already existed — stopping here (this is what a repeat email would do).")
        return

    entity_key = result["entity_key"]
    row = result["row"]
    invoice_no = row["invoice_no"]

    print(f"\n=== Stage 2: draft the email from the row just appended to the '{result['sheet']}' tab ===")
    pdf_path = OUTPUT_DIR / f"invoice_{invoice_no.replace('/', '-')}.pdf"
    pdf_meta = render_invoice_for_entity(entity_key, row, pdf_path)
    draft = compose_email(row, entity_key)

    print(json.dumps({"pdf": pdf_meta, "draft_email": draft}, indent=2, ensure_ascii=False))

    eml_preview = (
        f"To: {', '.join(draft['to'])}\n"
        f"Cc: {', '.join(draft['cc'])}\n"
        f"Subject: {draft['subject']}\n"
        f"Attachment: {Path(pdf_meta['path']).name}\n\n"
        f"{draft['body']}"
    )
    (OUTPUT_DIR / f"draft_preview_{invoice_no.replace('/', '-')}.eml").write_text(eml_preview, encoding="utf-8")
    print("\n=== Draft email preview ===")
    print(eml_preview)

    print(f"\n=== Stage 3: try to write the outcome back into row {invoice_no} ===")
    status_result = update_email_status(
        TRACKER_COPY, invoice_no, "Draft Created",
        sent_date=date(2026, 8, 7).isoformat(),
        sent_to=draft["to"], sent_cc=draft["cc"],
        entity_key=entity_key,
    )
    print(json.dumps(status_result, indent=2, default=str, ensure_ascii=False))
    if status_result["status"] == "not_supported":
        print(f"\n{status_result['message']} (this is expected for USA/UK/Poland 2026 — "
              "none of them track email status in your real tracker.)")

    _unresolved_bonus_scenario()


if __name__ == "__main__":
    main()
