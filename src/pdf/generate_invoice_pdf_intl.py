"""
generate_invoice_pdf_intl.py

Per-entity invoice PDF rendering for real client-facing formats: USA, UK,
Singapore, Poland. Each of these is a genuinely different layout (see
config/entities/*.json for the legal name / address / tax rule / bank
details / table-shape differences extracted from the reference PDFs).

Entry point:
    render_international_invoice(entity_key, row, out_path)

    entity_key: "usa" | "uk" | "singapore" | "poland"
    row: a dict describing one invoice — see the docstring on
         render_international_invoice() below for the exact shape.
    out_path: str/Path to write the PDF to

USA/UK/Poland are wired into the live pipeline (see entity_resolver.py and
draft_email_from_excel_row.py); Singapore isn't yet, since no Singapore
tab exists in the real tracker.
"""
import calendar
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.graphics.shapes import Drawing, Circle, String

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "config"
ENTITIES_DIR = CONFIG_DIR / "entities"

# Currency SYMBOL lookup, keyed by the same standard 3-letter codes that
# parse_invoice_summary.py's CURRENCY_ALIASES normalizes the email's
# "Currency" field to. Added 2026-08-13: this renderer used to ALWAYS use
# the entity's own static default currency/currency_symbol
# (config/entities/*.json) no matter what currency was actually parsed out
# of the CP's email and saved in the Excel row's Currency column -- so a
# USA-tab invoice always showed "$"/"USD" even after the accounts team
# wrote "Currency: INR" in the email and it was correctly saved to the
# sheet. Fixed by preferring the row's OWN currency (real, already-saved
# value) whenever present; only falling back to the entity's static default
# when the row has none at all (e.g. an older row saved before the
# Currency column/field existed, or an unrecognized code with no symbol).
CURRENCY_SYMBOLS = {
    "USD": "$", "GBP": "£", "INR": "Rs.", "EUR": "€", "PLN": "PLN", "SGD": "S$",
}

# config/entities/*.json's "table_columns" bakes the entity's DEFAULT
# currency symbol as static text into the amount column header itself
# (e.g. "AMOUNT($)" for USA, "AMOUNT(£)" for UK) -- found 2026-08-13 as a
# SECOND place (besides the totals/amount-in-words line, fixed above) this
# renderer was showing the wrong currency: even after currency_symbol was
# correctly resolved from the row's real Currency value, this header text
# was untouched static config, so a USA-tab invoice billed in INR still
# showed the column header "AMOUNT($)". Fixed by re-writing whatever
# symbol is inside the parens to match the ACTUAL resolved currency_symbol
# for this invoice, rather than trusting the static config text.
_AMOUNT_HEADER_RE = re.compile(r"^(AMOUNT)\(.*\)$", re.IGNORECASE)


def _localize_column_headers(cols, currency_symbol):
    sym = currency_symbol.strip()
    return [_AMOUNT_HEADER_RE.sub(lambda m: f"{m.group(1)}({sym})", c) for c in cols]

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


def _int_to_words_western(n):
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


def _int_to_words_indian(n):
    """Integer -> English words using the Indian digit-grouping convention
    (Crore / Lakh / Thousand, i.e. 2-2-3 grouping) rather than the Western
    Thousand/Million/Billion (3-3-3) grouping — e.g. 776440 -> 'Seven Lakh
    Seventy Six Thousand Four Hundred Forty', not 'Seven Hundred Seventy Six
    Thousand...'. This is the convention Indian GST invoices use; getting it
    wrong here would misstate the amount for an Indian reader even though
    the digits are correct."""
    if n == 0:
        return "Zero"
    crore, n = divmod(n, 10_000_000)
    lakh, n = divmod(n, 100_000)
    thousand, hundred_rest = divmod(n, 1_000)
    parts = []
    if crore:
        parts.append(f"{_three_digit_words(crore)} Crore")
    if lakh:
        parts.append(f"{_three_digit_words(lakh)} Lakh")
    if thousand:
        parts.append(f"{_three_digit_words(thousand)} Thousand")
    if hundred_rest:
        parts.append(_three_digit_words(hundred_rest))
    return " ".join(p for p in parts if p)


def amount_in_words(amount, system="western"):
    try:
        n = int(round(float(amount)))
    except (TypeError, ValueError):
        return ""
    return _int_to_words_indian(n) if system == "indian" else _int_to_words_western(n)


def _fmt_money(v, symbol="", decimals=2):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v) if v else "-"
    return f"{symbol} {v:,.2f}" if symbol else f"{v:,.2f}"


def _fmt_money_or_dash(v, symbol=""):
    """Poland's real invoice shows a bare '-' for a zero-rate tax line
    ('Add : VAT 0%    -') rather than '$ 0.00' — matches that convention
    generally, for any entity/amount that comes out to exactly zero."""
    try:
        if float(v) == 0:
            return "-"
    except (TypeError, ValueError):
        pass
    return _fmt_money(v, symbol)


_ORDINAL_SUFFIXES = {1: "st", 2: "nd", 3: "rd"}


def _ordinal_suffix(n):
    if 11 <= (n % 100) <= 13:
        return "th"
    return _ORDINAL_SUFFIXES.get(n % 10, "th")


def _parse_iso_date(raw):
    """Best-effort parse of 'YYYY-MM-DD' / 'YYYY/MM/DD' -> datetime, or None
    if it isn't in one of those shapes (e.g. blank, or already hand-
    formatted some other way)."""
    if not raw:
        return None
    raw = str(raw).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw[:10], fmt)
        except ValueError:
            continue
    return None


def _format_display_date(raw, month_format="abbr", year_digits=2):
    """'2026-08-05' -> \"05th Aug'26\" (Singapore/Poland style) or
    \"05th August'2026\" (UK style, month_format='full'/year_digits=4) —
    both zero-pad the day, per the real invoices. Falls back to returning
    the input unchanged if it isn't a parseable ISO date (e.g. already
    hand-formatted, or blank)."""
    if not raw:
        return ""
    dt = _parse_iso_date(raw)
    if dt is None:
        return str(raw).strip()  # already formatted some other way — show as-is rather than guess
    year = dt.strftime("%Y") if year_digits == 4 else dt.strftime("%y")
    month = calendar.month_name[dt.month] if month_format == "full" else calendar.month_abbr[dt.month]
    return f"{dt.day:02d}{_ordinal_suffix(dt.day)} {month}'{year}"


def _resolve_due_date(entity, row):
    """Row-supplied due_date wins if present (e.g. a genuine tracker value).
    Otherwise, if this entity has a due_days rule (USA: 30, UK: 7) and we
    have an invoice_date to count from, compute it — due_days was already
    being configured per entity but never actually used anywhere."""
    if row.get("due_date"):
        return row["due_date"]
    due_days = entity.get("due_days")
    inv_dt = _parse_iso_date(row.get("invoice_date"))
    if due_days and inv_dt:
        return (inv_dt + timedelta(days=due_days)).strftime("%Y-%m-%d")
    return None


def _logo_drawing(diameter_mm=15):
    """Draws the round 'INT.' logo seen on every real invoice — a filled
    blue circle with the wordmark in white. No image asset needed."""
    d_pt = diameter_mm * mm
    d = Drawing(d_pt, d_pt)
    d.add(Circle(d_pt / 2, d_pt / 2, d_pt / 2, fillColor=colors.HexColor("#1961ac"), strokeColor=None))
    d.add(String(d_pt / 2, d_pt / 2 - d_pt * 0.12, "INT.", textAnchor="middle",
                  fontName="Helvetica-Bold", fontSize=d_pt * 0.26, fillColor=colors.white))
    return d


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
    """Client block on the left, invoice meta on the right — directly, with
    no 'Bill To' / 'Invoice Details' captions above them. None of the real
    invoices (UK/Singapore/Poland) show those captions; they just show the
    client name (bold) + address, and the meta lines, right away."""
    normal = sty["normal"]

    client_lines = [f"<b>{row.get('client_name') or ''}</b>"]
    client_lines.extend(row.get("client_address_lines") or [])
    bill_to = Paragraph("<br/>".join(client_lines), normal)

    month_fmt = entity.get("date_month_format", "abbr")
    year_digits = entity.get("date_year_digits", 2)
    meta_lines = [
        f"Invoice No: <b>{row.get('invoice_no') or ''}</b>",
        f"Invoice Date: {_format_display_date(row.get('invoice_date'), month_fmt, year_digits)}",
    ]
    due_date = _resolve_due_date(entity, row)
    if due_date:
        meta_lines.append(f"Due Date: {_format_display_date(due_date, month_fmt, year_digits)}")
    if entity.get("show_po_fields"):
        if row.get("po_no"):
            meta_lines.append(f"PO No. {row['po_no']}")
        if row.get("po_date"):
            meta_lines.append(f"PO Dt. {row['po_date']}")
    if entity.get("gst_reg_no"):
        meta_lines.append(f"{entity.get('gst_reg_label', 'GST Reg No')}: {entity['gst_reg_no']}")
    if entity.get("vat_no") and entity["entity_key"] == "uk":
        meta_lines.append(f"{entity.get('vat_label', 'VAT NO')}: {entity['vat_no']}")

    # India: the CLIENT's own GSTIN and Place-of-Supply are legally required
    # on the invoice, alongside our GSTIN in the footer — confirmed present
    # on every real row in your tracker's Kar Ventures/Bad Debt tabs.
    if entity.get("show_client_gst_fields"):
        if row.get("client_gstin"):
            meta_lines.append(f"Client GSTIN: {row['client_gstin']}")
        if row.get("pos"):
            meta_lines.append(f"Place of Supply: {row['pos']}")

    meta = Paragraph("<br/>".join(meta_lines), normal)

    t = Table([[bill_to, meta]], colWidths=[95 * mm, 75 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return t


def _line_items_table_per_resource(entity, row, sty, currency_symbol):
    """USA-style: one row per resource, then a bold monthly-billing subtotal
    row, then a Sub-Total row (the real invoice repeats the same figure on
    both rows — kept as-is rather than assumed to be a typo)."""
    cols = _localize_column_headers(entity["table_columns"], currency_symbol)
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
    """UK/Singapore/Poland/India-style: one row, a (often multi-line)
    description, with a single total amount. India's real GST tabs
    (Kar Ventures/Bad Debt) also carry an HSN/SAC code per line — a 4th
    'table_columns' entry (checked by count, not by entity_key) triggers
    that extra column so this stays generic rather than India-specific."""
    cols = _localize_column_headers(entity["table_columns"], currency_symbol)
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

    has_hsn_column = len(cols) == 4
    if has_hsn_column:
        data_row = ["1", Paragraph(description_html, normal), row.get("hsn") or "", _fmt_money(amount, "")]
        col_widths = [14 * mm, 94 * mm, 28 * mm, 34 * mm]
    else:
        data_row = ["1", Paragraph(description_html, normal), _fmt_money(amount, "")]
        col_widths = [16 * mm, 116 * mm, 38 * mm]

    data = [header_row, data_row]
    table = Table(data, colWidths=col_widths)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3c6e")),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    if has_hsn_column:
        style.append(("ALIGN", (2, 1), (2, -1), "CENTER"))
    table.setStyle(TableStyle(style))
    return table, amount


def _finalize_totals_table(rows, entity, total, csym, extra_style=None):
    """Shared table-building step for all tax positions below. If this
    entity's amount-in-words is meant to REPLACE the final 'Total' label
    (Poland: the real invoice's last row reads 'Three Thousand $ 3,000.00'
    instead of 'Total $ 3,000.00'), append that as the actual final row here
    — rather than as a separate paragraph after the table — so the bold +
    line-above styling (which always targets the last row) lands on it."""
    if entity.get("amount_in_words") and entity.get("amount_in_words_replaces_total_label"):
        words_system = entity.get("amount_in_words_system", "western")
        rows = rows + [[amount_in_words(total, words_system), _fmt_money(total, csym)]]

    style = [
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE", (0, -1), (-1, -1), 0.75, colors.HexColor("#1a3c6e")),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if extra_style:
        style.extend(extra_style)
    table = Table(rows, colWidths=[130 * mm, 40 * mm])
    table.setStyle(TableStyle(style))
    return table


def _totals_block(entity, row, sty, subtotal, currency_symbol):
    """Builds the subtotal/tax/total rows below the line-items table,
    following each entity's own tax position/labelling — this is the part
    that differs the most across the four real invoices.

    Whether a currency symbol appears inline on THESE rows (as opposed to
    just in the "AMOUNT(...)" column header) varies by entity: UK's real
    invoice shows plain numbers here, Singapore and Poland show '$' inline.
    Controlled by entity["totals_show_currency_symbol"] (default True)."""
    tax_cfg = entity.get("tax")
    story = []
    csym = currency_symbol if entity.get("totals_show_currency_symbol", True) else ""

    if tax_cfg is None:
        # USA: no tax line at all — Sub-Total (already rendered in the
        # line-items table) IS the total. Nothing to add here; if this
        # entity's amount-in-words doesn't replace a label (USA doesn't),
        # render_international_invoice() adds it as its own line instead.
        return story, subtotal

    if tax_cfg["position"] == "cgst_sgst_after_subtotal":
        # India: two separate tax lines (CGST + SGST), each its own row,
        # THEN the total — per the CGST/SGST split confirmed in your real
        # Kar Ventures/Bad Debt tabs (9% + 9%, always intra-state for now).
        cgst = round(subtotal * tax_cfg["cgst_rate"], 2)
        sgst = round(subtotal * tax_cfg["sgst_rate"], 2)
        total = subtotal + cgst + sgst
        rows = [
            [tax_cfg.get("subtotal_label", "Sub-Total"), _fmt_money(subtotal, csym)],
            [tax_cfg.get("cgst_label", "CGST"), _fmt_money_or_dash(cgst, csym)],
            [tax_cfg.get("sgst_label", "SGST"), _fmt_money_or_dash(sgst, csym)],
        ]
        if not (entity.get("amount_in_words") and entity.get("amount_in_words_replaces_total_label")):
            rows.append(["Total", _fmt_money(total, csym)])
        story.append(_finalize_totals_table(rows, entity, total, csym))
        return story, total

    rate = tax_cfg["rate"]
    tax_amount = round(subtotal * rate, 2)
    total = subtotal + tax_amount

    if tax_cfg["position"] == "before_subtotal":
        # Poland: "Add : VAT 0%" line shown (dash if zero), then Sub-Total
        rows = [[tax_cfg["label"], _fmt_money_or_dash(tax_amount, csym)],
                [tax_cfg.get("subtotal_label", "Sub-Total"), _fmt_money(total, csym)]]
        story.append(_finalize_totals_table(rows, entity, total, csym))

    elif tax_cfg["position"] == "after_subtotal":
        # UK: explicit Sub-Total row (NOT already shown anywhere else —
        # the line-items table only shows the single item amount), then
        # "Vat 20%", then a final unlabelled total row (unless amount-in-
        # words replaces it).
        rows = [
            [tax_cfg.get("subtotal_label", "Sub-Total"), _fmt_money(subtotal, csym)],
            [tax_cfg["label"], _fmt_money_or_dash(tax_amount, csym)],
        ]
        if not (entity.get("amount_in_words") and entity.get("amount_in_words_replaces_total_label")):
            rows.append(["", _fmt_money(total, csym)])
        story.append(_finalize_totals_table(rows, entity, total, csym))

    elif tax_cfg["position"] == "verbose_after_subtotal":
        # Singapore: fully spelled-out labels, bold final line.
        rows = [
            [tax_cfg.get("subtotal_label", "Total amount payable excluding GST"), _fmt_money(subtotal, csym)],
            [tax_cfg["label"], _fmt_money_or_dash(tax_amount, csym)],
            [tax_cfg.get("total_label", "Total amount payable including GST"), _fmt_money(total, csym)],
        ]
        story.append(_finalize_totals_table(
            rows, entity, total, csym,
            extra_style=[("FONTSIZE", (0, -1), (-1, -1), 11)],
        ))

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


def _queries_line_html(query_line):
    """'accountsint@indusnet.co.in | +91-33-2357 6070' -> the email segment
    bold + blue (matches the real invoices' link styling), '|' kept plain,
    the rest bold. Works fine for a query_line with no '|' too (Poland)."""
    parts = [p.strip() for p in query_line.split("|")]
    styled = [f'<font color="#1961ac"><b>{p}</b></font>' if "@" in p else f"<b>{p}</b>" for p in parts]
    return "QUERIES : " + " | ".join(styled)


def _header_block(entity, sty):
    """Logo top-left, title + QUERIES line top-right — matches every real
    invoice seen so far (UK/Singapore/Poland). No legal name is shown up
    here at all in the real ones; it only appears in the footer."""
    title_style = ParagraphStyle(
        "IntlTitleReal", parent=sty["styles"]["Normal"], fontName="Helvetica",
        fontSize=24, leading=28, textColor=colors.HexColor("#333333"), alignment=2,  # 2 = right
    )
    queries_style = ParagraphStyle(
        "IntlQueries", parent=sty["styles"]["Normal"], fontSize=8.5, leading=11,
        textColor=colors.HexColor("#333333"), alignment=2,
    )
    right_cell = [Paragraph(entity["title"], title_style)]
    if entity.get("query_line"):
        right_cell.append(Spacer(1, 1.5 * mm))
        right_cell.append(Paragraph(_queries_line_html(entity["query_line"]), queries_style))

    header = Table([[_logo_drawing(), right_cell]], colWidths=[22 * mm, 152 * mm])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return header


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
        "currency": str, optional -- a standard 3-letter code (e.g. "USD",
            "INR") as saved in the row's Excel Currency column. When
            present, this OVERRIDES the entity's static default currency/
            symbol for this one invoice (see CURRENCY_SYMBOLS above) --
            added 2026-08-13 because the entity's default alone can't
            reflect a CP asking for a different billing currency on a
            given invoice. Falls back to the entity's own default currency
            if omitted or not in CURRENCY_SYMBOLS.
    }
    out_path: str/Path to write the PDF to
    """
    entity = load_entity_config(entity_key)
    sty = _styles()
    row_currency = (row.get("currency") or "").strip().upper()
    if row_currency and row_currency in CURRENCY_SYMBOLS:
        currency_code = row_currency
        currency_symbol = CURRENCY_SYMBOLS[row_currency]
    elif row_currency:
        # Recognized-as-present but not in our symbol table (e.g. a code
        # CURRENCY_ALIASES doesn't know about either) -- show the code
        # itself rather than silently falling back to the entity's
        # unrelated default, since that would misrepresent what was
        # actually billed.
        currency_code = row_currency
        currency_symbol = row_currency + " "
    else:
        currency_code = entity.get("currency", "")
        currency_symbol = entity.get("currency_symbol", "")

    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        topMargin=18 * mm, bottomMargin=18 * mm, leftMargin=18 * mm, rightMargin=18 * mm,
    )
    story = []

    story.append(_header_block(entity, sty))
    story.append(Spacer(1, 8 * mm))

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

    # USA/Poland/India use "amount_in_words"; Poland's (and India's, if
    # configured that way) version is already merged into the totals table
    # above as its final row's label — a separate line here is only needed
    # when it ISN'T replacing that label, e.g. USA.
    if entity.get("amount_in_words") and not entity.get("amount_in_words_replaces_total_label"):
        words_system = entity.get("amount_in_words_system", "western")
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"<i>Amount in words: {amount_in_words(total, words_system)} {currency_code}</i>",
                                sty["normal"]))

    story.append(Spacer(1, 8 * mm))

    default_intro = ("Please make the payment by PayPal or Bank Transfer, the details are given below:"
                      if (entity.get("payment_options") or {}).get("paypal")
                      else "Please make the payment with the following details :")
    intro_line = entity.get("payment_intro_line", default_intro)
    story.append(Paragraph(intro_line, ParagraphStyle(
        "PaymentIntro", parent=sty["normal"], fontSize=9,
    )))
    story.append(Spacer(1, 2 * mm))
    story.extend(_bank_details_block(entity, sty))
    story.append(Spacer(1, 4 * mm))

    note_line = entity.get(
        "payment_note_line", "*Note: Please mention the Invoice Number while making the payment"
    )
    story.append(Paragraph(note_line, ParagraphStyle(
        "PaymentNote", parent=sty["normal"], fontSize=8.5, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1961ac"),
    )))
    story.append(Spacer(1, 10 * mm))

    # Legal name + address + registration numbers ONLY — none of the real
    # invoices repeat the QUERIES line down here, it's a header-only thing.
    footer_lines = [entity["legal_name"]] + entity.get("footer_address_lines", [])
    reg_bits = []
    if entity.get("registration_no"):
        reg_bits.append(f"{entity.get('registration_label', 'CIN')}: {entity['registration_no']}")
    if entity.get("vat_no"):
        reg_bits.append(f"{entity.get('vat_label', 'VAT NO')}: {entity['vat_no']}")
    if entity.get("nip_no"):
        reg_bits.append(f"{entity.get('nip_label', 'NIP NO')}: {entity['nip_no']}")
    if entity.get("gstin"):
        reg_bits.append(f"{entity.get('gstin_label', 'GSTIN')}: {entity['gstin']}")
    if entity.get("pan"):
        reg_bits.append(f"{entity.get('pan_label', 'PAN')}: {entity['pan']}")
    if reg_bits:
        footer_lines.append(" | ".join(reg_bits))

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
