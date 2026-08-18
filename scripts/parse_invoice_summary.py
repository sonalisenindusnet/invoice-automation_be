"""
parse_invoice_summary.py

Parses a CP "Invoice Summary" email into a structured dict, using
config/field_schema.json as the field map. Handles three input shapes,
since Gmail can hand back the body as HTML, as plain text, or you might
already have it as a dict from an upstream step:

    1. HTML containing a <table> (Sl No | Particulars | Details)
    2. Plain text with one "field  value" pair per line
    3. A pre-built dict (e.g. {"Pf Id": "...", "Account Name": "...", ...})

Usage:
    python parse_invoice_summary.py path/to/email_body.html
    python parse_invoice_summary.py path/to/email_body.txt
    (or import parse_invoice_summary() and call it directly)
"""
import json
import logging
import re
import sys
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
FIELD_SCHEMA_PATH = CONFIG_DIR / "field_schema.json"

logger = logging.getLogger("parse_invoice_summary")

# Common ways a human marks a field as intentionally empty — treated as
# equivalent to a truly blank cell, not stored as literal text.
BLANK_PLACEHOLDERS = {"blank", "n/a", "na", "none", "-", "--", "nil", "nothing"}


def load_field_schema():
    with open(FIELD_SCHEMA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)["fields"]


def _clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def _split_emails(raw):
    """Split a comma/semicolon/newline separated email blob into a clean list."""
    if not raw:
        return []
    parts = re.split(r"[,;\n]+", raw)
    emails = []
    for p in parts:
        p = _clean(p)
        m = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", p)
        if m:
            emails.append(m.group(0))
    return emails


def _parse_currency(raw):
    """'Rs 78600' / 'Rs. 78,600' / '78600' -> 78600 (int) plus the original string."""
    raw = _clean(raw)
    digits = re.sub(r"[^\d.]", "", raw)
    value = None
    if digits:
        try:
            value = float(digits)
            if value == int(value):
                value = int(value)
        except ValueError:
            value = None
    return {"raw": raw, "amount": value}


def _extract_rows_from_html(html):
    """Return list of (sl_no, particulars, details) from the first table found."""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        raise RuntimeError(
            "beautifulsoup4 is required for HTML parsing. "
            "Install with: pip install beautifulsoup4 --break-system-packages"
        )
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return []
    rows = []
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if len(cells) < 2:
            continue
        # Skip header row ("Sl No" / "Particulars" / "Details")
        if cells[0].strip().lower() in ("sl no", "sl. no", "sl no."):
            continue
        if len(cells) >= 3:
            sl_no_raw, particulars, details = cells[0], cells[1], " ".join(cells[2:])
        else:
            sl_no_raw, particulars, details = "", cells[0], cells[1]
        sl_no = None
        m = re.search(r"\d+", sl_no_raw)
        if m:
            sl_no = int(m.group(0))
        rows.append((sl_no, _clean(particulars), _clean(details)))
    return rows


def _known_labels(schema):
    """All label aliases from the schema, longest first, so e.g. 'Client Mail
    id (To and CC)' is tried before any shorter alias that happens to be a
    prefix of it."""
    labels = []
    for field in schema:
        labels.extend(field.get("label_aliases", []))
    return sorted(set(labels), key=len, reverse=True)


def _extract_rows_from_text(text, schema=None):
    """
    Handles plain-text tables, one row per line, in ANY of these real-world
    shapes — by matching against the schema's known labels directly, rather
    than guessing column boundaries from whitespace (which breaks the moment
    an email client collapses everything to single spaces):

      "1\tPf Id\t2603/BFS/5084"          (pasted from Excel/Outlook — tabs)
      "Pf Id    2603/BFS/5084"           (2+ spaces between columns)
      "1. Pf Id: 2603/BFS/5084"          (hand-typed, colon-separated)
      "1 Pf Id 2603/BFS/5084"            (single spaces — typical of a real
                                           pasted/forwarded email body)

    A line that doesn't start with a recognized label is treated as a
    continuation of the previous row's value (handles a long CC list that
    wraps onto a second line). Parsing stops once the highest Sl No in the
    schema has been seen, so trailing signature/disclaimer text is ignored
    rather than getting appended to the last field.
    """
    schema = schema or load_field_schema()

    logger.info("Extracting rows from text - schema : %s, schema has %d fields", schema, len(schema))

    labels = _known_labels(schema)

    logger.info("labels : %s", labels)

    max_sl_no = max((f.get("sl_no") or 0) for f in schema) if schema else None

    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if re.match(r"^sl\.?\s*no\.?\s+particulars\s+details$", line, re.IGNORECASE):
            continue  # header row

        sl_no = None
        rest = line
        m = re.match(r"^(\d+)[.\)]?\s+(.*)$", line)
        if m:
            sl_no = int(m.group(1))
            rest = m.group(2)

        matched_label = next((lbl for lbl in labels if rest.lower().startswith(lbl.lower())), None)

        if matched_label:
            details = rest[len(matched_label):].strip(" \t:-")
            rows.append([sl_no, matched_label, details])
            if max_sl_no and sl_no == max_sl_no:
                break
        elif rows:
            rows[-1][2] = (rows[-1][2] + " " + rest).strip()
        # else: stray line before any row has matched yet (greeting, etc.) — ignore

    return [(sl_no, _clean(label), _clean(details)) for sl_no, label, details in rows]


def _rows_from_dict(d):
    rows = []
    for i, (k, v) in enumerate(d.items(), start=1):
        rows.append((i, _clean(str(k)), _clean(str(v))))
    return rows


def parse_invoice_summary(source, source_type="auto"):
    """
    source: raw HTML string, raw plain-text string, or a dict of {label: value}
    source_type: "html" | "text" | "dict" | "auto"
    Returns: (data: dict, warnings: list[str])
    """
    schema = load_field_schema()
    warnings = []

    logger.info("Parsing invoice summary, source=%s, source_type=%s", source, source_type)

    if source_type == "auto":
        if isinstance(source, dict):
            source_type = "dict"
        elif "<table" in source.lower() or "<td" in source.lower():
            source_type = "html"
        else:
            source_type = "text"

    if source_type == "dict":
        rows = _rows_from_dict(source)
    elif source_type == "html":
        rows = _extract_rows_from_html(source)
    else:
        rows = _extract_rows_from_text(source, schema)

    if not rows:
        warnings.append("No rows could be extracted from the input — check the source format.")

    # Build lookup by sl_no, and a fallback queue by label for rows with no sl_no
    by_sl_no = {sl: (label, details) for sl, label, details in rows if sl is not None}
    label_queue = {}
    for sl, label, details in rows:
        label_queue.setdefault(label.lower(), []).append(details)

    data = {}
    for field in schema:
        key = field["key"]
        sl_no = field.get("sl_no")
        aliases = [a.lower() for a in field.get("label_aliases", [])]
        raw_value = None

        if sl_no is not None and sl_no in by_sl_no:
            label, details = by_sl_no[sl_no]
            raw_value = details
        else:
            # fallback: consume the next matching label from the queue
            for alias in aliases:
                if alias in label_queue and label_queue[alias]:
                    raw_value = label_queue[alias].pop(0)
                    break

        if raw_value is not None and raw_value.strip("() ").lower() in BLANK_PLACEHOLDERS:
            raw_value = None  # e.g. "(Blank)", "N/A", "-" — a human explicitly marking it empty

        if raw_value is None:
            if field.get("required"):
                warnings.append(f"Missing required field: {key} (Sl No {sl_no})")
            data[key] = None
            continue

        ftype = field.get("type", "text")
        if ftype == "email_list":
            data[key] = _split_emails(raw_value)
            if field.get("required") and not data[key]:
                warnings.append(f"Field '{key}' had no valid email addresses in: {raw_value!r}")
        elif ftype == "currency":
            data[key] = _parse_currency(raw_value)
        else:
            data[key] = _clean(raw_value)

    if data.get("master_project_id") in (None, ""):
        warnings.append(
            "Master Project ID is blank — MIS validation (Project ID lookup) is out of scope "
            "for this phase, but flagging so it isn't silently lost."
        )

    data["tab"] = extract_tab(source).lower() if "tab" not in data or not data["tab"] else data["tab"].lower() 

    return data, warnings

def extract_tab(body: str) -> str:
    match = re.search(
        r"(?im)^\s*Tab\s*:\s*(.+?)\s*$",
        body
    )

    if not match:
        raise ValueError("Tab is missing from invoice email")

    return match.group(1).strip()


def main():
    if len(sys.argv) != 2:
        print("Usage: python parse_invoice_summary.py <path-to-email-body.html-or-.txt>")
        sys.exit(1)
    path = Path(sys.argv[1])
    text = path.read_text(encoding="utf-8")
    source_type = "html" if path.suffix.lower() in (".html", ".htm") else "text"
    data, warnings = parse_invoice_summary(text, source_type=source_type)
    print(json.dumps({"data": data, "warnings": warnings}, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
