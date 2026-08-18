"""
generate_invoice_pdf_intl.py

Per-entity invoice PDF rendering for the four non-India formats the user
supplied as real reference invoices: USA, UK, Singapore, Poland. Each of
these is a genuinely different layout (see config/entities/*.json for the
legal name / address / tax rule / bank details / table-shape differences
extracted from the reference PDFs) — this module does NOT touch or import
generate_invoice_pdf.py, so the working India/GST flow used by
draft_email_from_excel_row.py and email_server.py is unaffected.

Entry point:
    render_international_invoice(entity_key, row, out_path)

    entity_key: "usa" | "uk" | "singapore" | "poland"
    row: a dict describing one invoice — see the docstring on
         render_international_invoice() below for the exact shape.
    out_path: str/Path to write the PDF to

This is intentionally decoupled from india_tab_schema.json: the existing
Excel tabs for USA/UK/Poland (in the user's real tracker) only carry
summary/tracking columns (dates, amounts, payment status), not the
line-item/description/bank-detail data an invoice PDF needs, so for now
this module is driven by an explicit `row` dict built by the caller
(e.g. a future per-entity draft script, or manual test data). Wiring this
to a live per-country email pipeline is a follow-up step once it's clear
how an incoming CP email indicates which entity/country an invoice is for.
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

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
ENTITIES_DIR = CONFIG_DIR / "entities"

_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
         "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
         "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def load_entity_config(entity_key):
    path = ENTITIES_DIR / f"{entity_key}.json"
    if not path.exists():
        raise ValueError(f"No entity config for '{entity_key}' (expected {path})")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _three_digit_words(n):
    if n == 0:
        return ""
    words = []
    if n >= 100:
        words.append(_ONES[n // 100] + " Hundred")
        n %= 100
    if n >= 20:
        words.append(_TENS[n // 10])
        if n % 10:
            words.append(_ONES[n % 10])
    elif n > 0:
        words.append(_ONES[n])
    return " ".join(words)


def _int_to_words(n):
    """Integer -> English words, Title Case, e.g. 26560 -> 'Twenty Six Thousand
    Five Hundred Sixty'. No currency name/decimals — matches how the real USA
    and Poland invoices phrase their amount-in-words line."""
    if n == 0:
        return "Zero"
    scale_names = [(1_000_000_000, "Billion"), (1_000_000, "Million"), (1_000, "Thousand")]
    parts = []
    remaining = n
    for scale, name in scale_names:
        if remaining >= scale:
            count = remaining // scale
            parts.append(f"{_three_digit_words(count)} {name}")
            remaining %= scale
    if remaining:
        parts.append(_three_digit_words(remaining))
    return " ".join(p for p in parts if p)


def amount_in_words(amount):
    try:
        n = int(round(float(amount)))
    except (TypeError, ValueError):
        return ""
    return _int_to_words(n)


def _fmt_money(v, symbol="", decimals=2):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v) if v else "-"
    return f"{symbol} {v:,.2f}" if symbol else f"{v:,.2f}"


def _styles():
    styles = getSampleStyleSheet()
    return {
        "styles": styles,
        "title": ParagraphStyle("IntlTitle", parent=styles["Heading1"], fontSize=18,
                                 textColor=colors.HexColor("#1a3c6e"), spaceAfter=2),
        "label": ParagraphStyle("IntlLabel", parent=styles["Normal"],
                                 textColor=colors.HexColor("#555555"), fontSize=8.5),
        "normal": styles["Normal"],
        "small_center": ParagraphStyle("SmallCenter", parent=styles["Normal"], fontSize=8,
                                        textColor=colors.HexColor("#666666"), alignment=1),
        "h3": ParagraphStyle("H3", parent=styles["Heading3"], fontSize=11),
    }


def _bill_to_and_meta_table(entity, row, sty):
    normal = sty["normal"]
    label = sty["label"]

    client_lines = [f"<b>{row.get('client_name') or ''}</b>"]
    client_lines.extend(row.get("client_address_lines") or [])
    bill_to = Paragraph("<br/>".join(client_lines), normal)

    meta_lines = [
        f"Invoice No: <b>{row.get('invoice_no') or ''}</b>",
        f"Invoice Date: {row.get('invoice_date') or ''}",
    ]
    if row.get("due_date"):
        meta_lines.append(f"Due Date: {row['due_date']}")
    if entity.get("show_po_fields"):
        if row.get("po_no"):
            meta_lines.append(f"PO No.: {row['po_no']}")
        if row.get("po_date"):
            meta_lines.append(f"PO Dt.: {row['po_date']}")
    if entity.get("gst_reg_no"):
        meta_lines.append(f"{entity.get('gst_reg_label', 'GST Reg No')}: {entity['gst_reg_no']}")
    if entity.get("vat_no") and entity["entity_key"] == "uk":
        meta_lines.append(f"{entity.get('vat_label', 'VAT NO')}: {entity['vat_no']}")

    meta = Paragraph("<br/>".join(meta_lines), normal)

    t = Table([[Paragraph("Bill To", label), Paragraph("Invoice Details", label)], [bill_to, meta]],
               colWidths=[95 * mm, 75 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, 0), 4)]))
    return t


def _line_items_table_per_resource(entity, row, sty, currency_symbol):
    """USA-style: one row per resource, then a bold monthly-billing subtotal
    row, then a Sub-Total row (the real invoice repeats the same figure on
    both rows — kept as-is rather than assumed to be a typo)."""
    cols = entity["table_columns"]
    normal = sty["normal"]
    header_row = [Paragraph(f"<b>{c}</b>", ParagraphStyle(
        "HeadWhite2", parent=normal, textColor=colors.white, fontSize=9
    )) for c in cols]
    data = [header_row]
    line_items = row.get("line_items") or []
    for i, item in enumerate(line_items, start=1):
        data.append([str(i), item.get("label", ""), _fmt_money(item.get("amount"), "")])

    subtotal = row.get("subtotal")
    if subtotal is None:
        subtotal = sum((item.get("amount") or 0) for item in line_items)

    bold_row_start = len(data)
    data.append(["", entity.get("monthly_billing_label", "Total Monthly Billing"), _fmt_money(subtotal, "")])
    data.append(["", entity.get("subtotal_label", "Sub-Total"), _fmt_money(subtotal, "")])

    table = Table(data, colWidths=[16 * mm, 116 * mm, 38 * mm])
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3c6e")),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -3), "CENTER"),
        ("GRID", (0, 0), (-1, -3), 0.5, colors.HexColor("#cccccc")),
        ("FONTNAME", (0, bold_row_start), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE", (0, bold_row_start), (-1, bold_row_start), 0.75, colors.HexColor("#1a3c6e")),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    table.setStyle(TableStyle(style))
    return table, subtotal


def _line_items_table_single_line(entity, row, sty, currency_symbol):
    """UK/Singapore/Poland-style: one row, a (often multi-line) description,
    with a single total amount."""
    cols = entity["table_columns"]
    normal = sty["normal"]
    description = row.get("description") or (row.get("line_items") or [{}])[0].get("label", "")
    description_html = description.replace("\n", "<br/>")
    amount = row.get("subtotal")
    if amount is None:
        line_items = row.get("line_items") or []
        amount = sum((item.get("amount") or 0) for item in line_items)

    header_row = [Paragraph(f"<b>{c}</b>", ParagraphStyle(
        "HeadWhite", parent=normal, textColor=colors.white, fontSize=9
    )) for c in cols]
    data = [header_row, ["1", Paragraph(description_html, normal), _fmt_money(amount, "")]]
    col_widths = [16 * mm, 116 * mm, 38 * mm]

    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3c6e")),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return table, amount


def _totals_block(entity, row, sty, subtotal, currency_symbol):
    """Builds the subtotal/tax/total rows below the line-items table,
    following each entity's own tax position/labelling — this is the part
    that differs the most across the four real invoices."""
    normal = sty["normal"]
    tax_cfg = entity.get("tax")
    story = []

    if tax_cfg is None:
        # USA: no tax line at all — Sub-Total (already rendered in the
        # line-items table) IS the total.
        total = subtotal
        return story, total

    rate = tax_cfg["rate"]
    tax_amount = round(subtotal * rate, 2)

    if tax_cfg["position"] == "before_subtotal":
        # Poland: "Add : VAT 0%" line shown, then Sub-Total = subtotal + tax
        rows = [[tax_cfg["label"], _fmt_money(tax_amount, "")],
                [tax_cfg.get("subtotal_label", "Sub-Total"), _fmt_money(subtotal + tax_amount, "")]]
        table = Table(rows, colWidths=[130 * mm, 40 * mm])
        table.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LINEABOVE", (0, -1), (-1, -1), 0.75, colors.HexColor("#1a3c6e")),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        total = subtotal + tax_amount

    elif tax_cfg["position"] == "after_subtotal":
        # UK: Sub-Total already shown in line-items table; add "Vat 20%" then
        # a final unlabelled total row.
        rows = [[tax_cfg["label"], _fmt_money(tax_amount, "")],
                ["", _fmt_money(subtotal + tax_amount, "")]]
        table = Table(rows, colWidths=[130 * mm, 40 * mm])
        table.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LINEABOVE", (0, -1), (-1, -1), 0.75, colors.HexColor("#1a3c6e")),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        total = subtotal + tax_amount

    elif tax_cfg["position"] == "verbose_after_subtotal":
        # Singapore: fully spelled-out labels, bold final line.
        rows = [
            [tax_cfg.get("subtotal_label", "Total amount payable excluding GST"), _fmt_money(subtotal, "")],
            [tax_cfg["label"], _fmt_money(tax_amount, "")],
            [tax_cfg.get("total_label", "Total amount payable including GST"), _fmt_money(subtotal + tax_amount, "")],
        ]
        table = Table(rows, colWidths=[130 * mm, 40 * mm])
        table.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, -1), (-1, -1), 11),
            ("LINEABOVE", (0, -1), (-1, -1), 0.75, colors.HexColor("#1a3c6e")),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        total = subtotal + tax_amount

    else:
        total = subtotal

    return story, total


def _bank_details_block(entity, sty):
    normal = sty["normal"]
    bank = entity.get("bank_details", {})
    label_map = [
        ("bank_name", "Bank Name"), ("bank_address", "Bank Address"),
        ("beneficiary_name", "Beneficiary Name"), ("account_name", "A/C Name"),
        ("account_number", "Account Number" if "beneficiary_name" in bank else "A/C No"),
        ("aba_routing_number", "ABA Routing Number"), ("swift_code", "SWIFT Code"),
        ("sort_code", "Sort Code"), ("type_of_account", "Type of Account"),
        ("account_type", "Account Type"), ("beneficiary_address", "Beneficiary Address"),
    ]
    lines = []
    seen = set()
    for key, display in label_map:
        if key in bank and key not in seen:
            lines.append(f"{display}: {bank[key]}")
            seen.add(key)
    story = [Paragraph("Bank Details", sty["h3"]), Paragraph("<br/>".join(lines), normal)]
    payment_options = entity.get("payment_options")
    if payment_options and payment_options.get("paypal"):
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"Or PayPal: {payment_options['paypal']}", normal))
    return story


def render_international_invoice(entity_key, row, out_path):
    """
    entity_key: "usa" | "uk" | "singapore" | "poland"
    row: {
        "invoice_no": str, "invoice_date": str, "due_date": str (optional),
        "po_no": str (Poland only), "po_date": str (Poland only),
        "client_name": str, "client_address_lines": [str, ...],
        "month_label": str, e.g. "Aug'26" (used when the entity's table
            shows a "For the month of ..." sub-header),
        # USA (table_mode = "per_resource"):
        "line_items": [{"label": "Goutam Barai (PM)", "amount": 2240.0}, ...],
        # UK / Singapore / Poland (table_mode = "single_line"):
        "description": str (can be multi-line, use "\\n" for line breaks),
        "subtotal": number (optional — computed from line_items/description
            amount if omitted),
    }
    out_path: str/Path to write the PDF to
    """
    entity = load_entity_config(entity_key)
    sty = _styles()
    currency_symbol = entity.get("currency_symbol", "")

    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        topMargin=18 * mm, bottomMargin=18 * mm, leftMargin=18 * mm, rightMargin=18 * mm,
    )
    story = []

    story.append(Paragraph(entity["legal_name"], sty["title"]))
    story.append(Paragraph(entity["title"].upper(), ParagraphStyle(
        "IntlH1", parent=sty["styles"]["Heading2"], fontSize=13,
        textColor=colors.HexColor("#555555"), spaceAfter=4,
    )))
    story.append(HRFlowable(width="100%", color=colors.HexColor("#1a3c6e"), thickness=1))
    story.append(Spacer(1, 4 * mm))

    story.append(_bill_to_and_meta_table(entity, row, sty))
    story.append(Spacer(1, 6 * mm))

    if entity.get("show_month_subheader") and row.get("month_label"):
        story.append(Paragraph(f"For the month of {row['month_label']}",
                                ParagraphStyle("Month", parent=sty["normal"], fontName="Helvetica-Bold", fontSize=9)))
        story.append(Spacer(1, 2 * mm))

    if entity["table_mode"] == "per_resource":
        table, subtotal = _line_items_table_per_resource(entity, row, sty, currency_symbol)
    else:
        table, subtotal = _line_items_table_single_line(entity, row, sty, currency_symbol)
    story.append(table)
    story.append(Spacer(1, 2 * mm))

    totals_story, total = _totals_block(entity, row, sty, subtotal, currency_symbol)
    story.extend(totals_story)

    if entity.get("amount_in_words"):
        words_amount = total if not entity.get("amount_in_words_replaces_total_label") else total
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"<i>Amount in words: {amount_in_words(words_amount)} {entity['currency']}</i>",
                                sty["normal"]))

    story.append(Spacer(1, 8 * mm))
    story.extend(_bank_details_block(entity, sty))
    story.append(Spacer(1, 10 * mm))

    footer_lines = [entity["legal_name"]] + entity.get("footer_address_lines", [])
    reg_bits = []
    if entity.get("registration_no"):
        reg_bits.append(f"{entity.get('registration_label', 'CIN')}: {entity['registration_no']}")
    if entity.get("vat_no"):
        reg_bits.append(f"{entity.get('vat_label', 'VAT NO')}: {entity['vat_no']}")
    if entity.get("nip_no"):
        reg_bits.append(f"{entity.get('nip_label', 'NIP NO')}: {entity['nip_no']}")
    if reg_bits:
        footer_lines.append(" | ".join(reg_bits))
    if entity.get("query_line"):
        footer_lines.append(entity["query_line"])

    story.append(HRFlowable(width="100%", color=colors.HexColor("#cccccc"), thickness=0.5))
    story.append(Spacer(1, 2 * mm))
    for line in footer_lines:
        story.append(Paragraph(line, sty["small_center"]))

    doc.build(story)
    return {"entity": entity_key, "invoice_number": row.get("invoice_no"), "path": str(out_path), "total": total}


def main():
    if len(sys.argv) != 4:
        print("Usage: python generate_invoice_pdf_intl.py <usa|uk|singapore|poland> <row.json> <output.pdf>")
        sys.exit(1)
    entity_key = sys.argv[1]
    with open(sys.argv[2], "r", encoding="utf-8") as f:
        row = json.load(f)
    meta = render_international_invoice(entity_key, row, sys.argv[3])
    print(json.dumps(meta, indent=2, default=str))


if __name__ == "__main__":
    main()
