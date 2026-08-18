"""
generate_invoice_pdf.py

GST-aware invoice PDF, built from a row dict as produced by
append_invoice_to_excel.py (i.e. what's now sitting in the India tab).
Shows Base Amount / CGST / SGST / IGST / Round Off / Total as separate
line items, which the earlier flat-total demo (generate_invoice_pdf.py)
did not.

Same swap-later design as generate_invoice_pdf.py: once you send a real
invoice template, only render_invoice_pdf()'s body changes.
"""
import json
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "config"
COMPANY_PROFILE_PATH = CONFIG_DIR / "company_profile.json"


def load_company_profile():
    with open(COMPANY_PROFILE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _fmt_money(v, currency="INR"):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v) if v else "-"
    symbol = "Rs" if currency == "INR" else currency
    return f"{symbol} {v:,.2f}"


def render_invoice_pdf(row, out_path, company=None):
    """
    row: dict with keys matching config/india_tab_schema.json column keys
         (client_company, client_contact_person, pf_id, invoice_no,
          invoice_date, invoice_description, base_amount, cgst, sgst,
          igst, round_off, total, gstin, hsn, currency, ...)
    out_path: str/Path to write the PDF to
    """
    company = company or load_company_profile()
    currency = row.get("currency") or "INR"

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("InvoiceTitle", parent=styles["Title"], fontSize=20, spaceAfter=2)
    label_style = ParagraphStyle("Label", parent=styles["Normal"], textColor=colors.HexColor("#555555"), fontSize=9)
    normal = styles["Normal"]

    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        topMargin=20 * mm, bottomMargin=20 * mm, leftMargin=20 * mm, rightMargin=20 * mm,
    )
    story = []

    story.append(Paragraph(company.get("company_name", ""), title_style))
    for line in company.get("company_address_lines", []):
        story.append(Paragraph(line, label_style))
    story.append(Paragraph(f"GSTIN: {company.get('gstin', '')}", label_style))
    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph("TAX INVOICE", ParagraphStyle(
        "H1", parent=styles["Heading1"], fontSize=16, textColor=colors.HexColor("#1a3c6e")
    )))
    story.append(HRFlowable(width="100%", color=colors.HexColor("#1a3c6e"), thickness=1))
    story.append(Spacer(1, 4 * mm))

    bill_to_lines = [f"<b>{row.get('client_company') or ''}</b>"]
    if row.get("client_contact_person"):
        bill_to_lines.append(f"Attn: {row['client_contact_person']}")
    if row.get("gstin"):
        bill_to_lines.append(f"GSTIN: {row['gstin']}")
    if row.get("client_mail_to"):
        bill_to_lines.append(row["client_mail_to"])

    meta_table_data = [
        [Paragraph("Bill To", label_style), Paragraph("Invoice Details", label_style)],
        [
            Paragraph("<br/>".join(bill_to_lines), normal),
            Paragraph(
                f"Invoice No: <b>{row.get('invoice_no') or ''}</b><br/>"
                f"Invoice Date: {row.get('invoice_date') or ''}<br/>"
                f"Project ID: {row.get('pf_id') or ''}<br/>"
                f"Master Project ID: {row.get('master_project_id') or '—'}<br/>"
                f"HSN: {row.get('hsn') or '—'}<br/>"
                f"Work Order: {row.get('work_order') or ''}",
                normal,
            ),
        ],
    ]
    meta_table = Table(meta_table_data, colWidths=[85 * mm, 85 * mm])
    meta_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, 0), 4)]))
    story.append(meta_table)
    story.append(Spacer(1, 8 * mm))

    # Description line item (base amount only)
    desc_table = Table(
        [["Description", "Amount"], [row.get("invoice_description") or "", _fmt_money(row.get("base_amount"), currency)]],
        colWidths=[130 * mm, 40 * mm],
    )
    desc_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3c6e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, 1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(desc_table)
    story.append(Spacer(1, 4 * mm))

    # GST breakdown
    gst_rows = [["", "Amount"], ["Base Amount", _fmt_money(row.get("base_amount"), currency)]]
    if row.get("cgst"):
        gst_rows.append(["CGST @ 9%", _fmt_money(row.get("cgst"), currency)])
    if row.get("sgst"):
        gst_rows.append(["SGST @ 9%", _fmt_money(row.get("sgst"), currency)])
    if row.get("igst"):
        gst_rows.append(["IGST @ 18%", _fmt_money(row.get("igst"), currency)])
    if row.get("round_off"):
        gst_rows.append(["Round Off", _fmt_money(row.get("round_off"), currency)])
    gst_rows.append(["Total Payable", _fmt_money(row.get("total"), currency)])

    gst_table = Table(gst_rows, colWidths=[130 * mm, 40 * mm])
    gst_table.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LINEABOVE", (0, -1), (-1, -1), 0.75, colors.HexColor("#1a3c6e")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(gst_table)
    story.append(Spacer(1, 4 * mm))

    if row.get("total_order_value"):
        story.append(Paragraph(
            f"<font size=8 color='#777777'>Total Order Value (reference, full engagement): "
            f"{_fmt_money(row.get('total_order_value'), currency)}</font>", normal
        ))
    story.append(Spacer(1, 8 * mm))

    story.append(Paragraph("Payment Terms", ParagraphStyle("H3", parent=styles["Heading3"], fontSize=11)))
    story.append(Paragraph(company.get("default_payment_terms", ""), normal))
    story.append(Spacer(1, 4 * mm))

    bank = company.get("bank_details", {})
    bank_lines = "<br/>".join([
        f"Account Name: {bank.get('account_name', '')}",
        f"Account Number: {bank.get('account_number', '')}",
        f"IFSC: {bank.get('ifsc', '')}",
        f"Bank: {bank.get('bank_name', '')}",
    ])
    story.append(Paragraph("Bank Details", ParagraphStyle("H3b", parent=styles["Heading3"], fontSize=11)))
    story.append(Paragraph(bank_lines, normal))
    story.append(Spacer(1, 10 * mm))

    for line in company.get("signature_block", []):
        story.append(Paragraph(line, normal))

    doc.build(story)
    return {"invoice_number": row.get("invoice_no"), "path": str(out_path)}


def main():
    if len(sys.argv) != 3:
        print("Usage: python generate_invoice_pdf.py <row.json> <output.pdf>")
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        row = json.load(f)
    meta = render_invoice_pdf(row, sys.argv[2])
    print(json.dumps(meta, indent=2))


if __name__ == "__main__":
    main()
