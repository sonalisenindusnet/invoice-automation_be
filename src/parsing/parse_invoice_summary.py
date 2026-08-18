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
import re
import sys
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "config"
FIELD_SCHEMA_PATH = CONFIG_DIR / "field_schema.json"

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


# NEW 2026-08-13: a "Currency" row was added to the CP table (e.g.
# "12  Currency  INR" or "... Dollar"), per explicit instruction ("here
# will be 12. currency: inr/dollar will be added in the body and this
# value will be added into the excel"). Only a couple of example words
# were given -- normalize the common real-world spellings we know about to
# a standard 3-letter code; anything NOT recognized is never silently
# guessed into some other code -- it's kept verbatim (just uppercased) so
# it's visibly wrong/reviewable in the sheet, with a warning raised so it
# doesn't get missed. Add more aliases here as real emails reveal them.
CURRENCY_ALIASES = {
    "inr": "INR", "rupee": "INR", "rupees": "INR", "rs": "INR", "rs.": "INR",
    "usd": "USD", "dollar": "USD", "dollars": "USD", "us dollar": "USD", "us dollars": "USD",
    "gbp": "GBP", "pound": "GBP", "pounds": "GBP", "sterling": "GBP",
    "eur": "EUR", "euro": "EUR", "euros": "EUR",
    "pln": "PLN", "zloty": "PLN", "zlotys": "PLN",
    "sgd": "SGD", "singapore dollar": "SGD", "singapore dollars": "SGD",
}


def _normalize_currency(raw_value):
    """Returns (normalized_code, warning_message_or_None)."""
    key = raw_value.strip().lower()
    normalized = CURRENCY_ALIASES.get(key)
    if normalized:
        return normalized, None
    upper = raw_value.strip().upper()
    return upper, (
        f"Currency value {raw_value!r} wasn't recognized (expected something like "
        f"INR/Dollar/USD/GBP/Euro/etc) — kept as {upper!r} verbatim; check the Currency "
        f"cell for this row and consider adding it to CURRENCY_ALIASES if it's valid."
    )


def _rows_from_table(table):
    """Extract (sl_no, particulars, details) rows from ONE <table> element.
    Also reports whether this table had the literal "Sl No" header row, as
    a strong positive signal for _extract_rows_from_html's table-selection
    logic below."""
    rows = []
    has_header_signature = False
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if len(cells) < 2:
            continue
        # Skip header row ("Sl No" / "Particulars" / "Details")
        if cells[0].strip().lower() in ("sl no", "sl. no", "sl no."):
            has_header_signature = True
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
    return rows, has_header_signature


def _extract_rows_from_html(html, schema=None):
    """Return list of (sl_no, particulars, details) from the CORRECT table
    in the HTML -- not necessarily the first one found.

    Added 2026-08-12: the real flow is now CP -> account team member ->
    account team member forwards into the shared inbox this server reads.
    A forwarded message can carry a SECOND table before the real one -- for
    example some mail clients (Outlook in particular) render their own
    "From / Sent / To / Subject" forward-header block as an actual HTML
    <table>. Blindly taking the first table (the old behavior) would grab
    that instead of the CP's real Invoice Summary table.

    Selection rule: extract candidate rows from EVERY table in the
    document, score each one by (a) whether it has the literal "Sl No"
    header row (counts for a lot -- a forward-header table never has this),
    and (b) how many of its rows' first-cell text matches a KNOWN field
    label from config/field_schema.json (e.g. "Pf Id", "Account Name") --
    a forward-header table's "From"/"Sent"/"To"/"Subject" labels won't
    match any of these. Pick the highest-scoring table.

    If NO table shows any real signal at all (score 0 everywhere -- e.g. a
    single plain table with no recognizable structure), fall back to the
    first table, exactly like the pre-2026-08-12 behavior -- so every email
    that already parses correctly today keeps parsing exactly the same way.
    """
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        raise RuntimeError(
            "beautifulsoup4 is required for HTML parsing. "
            "Install with: pip install beautifulsoup4 --break-system-packages"
        )
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        return []

    schema = schema or load_field_schema()
    known_labels = [lbl.lower() for lbl in _known_labels(schema)]

    best_rows, best_score = None, -1
    for table in tables:
        rows, has_header_signature = _rows_from_table(table)
        if not rows:
            continue
        label_matches = sum(
            1 for _, particulars, _ in rows
            if any(particulars.lower().startswith(lbl) for lbl in known_labels)
        )
        score = label_matches + (100 if has_header_signature else 0)
        if score > best_score:
            best_score, best_rows = score, rows

    if best_rows is not None and best_score > 0:
        return best_rows
    # No table showed any real signal -- fall back to the first table,
    # matching the old (pre-multi-table-aware) behavior exactly.
    fallback_rows, _ = _rows_from_table(tables[0])
    return fallback_rows


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
    wraps onto a second line).

    Stops at a BLANK line once at least one row has already matched -- BUT
    ONLY if a lookahead shows what follows genuinely ISN'T another table
    row. Three iterations to get here, each fixing a real failure the
    previous one caused (all found/fixed 2026-08-13):
      1. Originally stopped once the highest Sl No IN THE SCHEMA had been
         seen. Broke the moment that highest number belonged to a rarely-
         used OPTIONAL field (e.g. 'Client Address', Sl No 15): a real
         email that never reaches that number never stops, so a trailing
         "Thanks,\nAccounts Team" signature got silently glued onto the
         last matched field's value.
      2. Changed to "stop at the FIRST blank line, period." Too aggressive
         the other way: a REAL forwarded email confirmed live the same day
         puts a blank line between literally EVERY row of the table (an
         artifact of the mail client's plain-text rendering) -- that broke
         parsing after row 1, every time.
      3. Also confirmed live: a field number appearing OUT OF ITS EXPECTED
         SEQUENCE (e.g. the rare 'Client Address' Sl No 15 sent in the
         MIDDLE of the table instead of at the end) made the OLD "highest
         Sl No" rule stop right there, silently discarding every real field
         physically after it in the text -- so relying on Sl No order at
         all for the stop condition is unsafe.
    Current rule, content-based rather than number-based: on a blank line,
    peek ahead past any further blank lines to the next real line. If THAT
    line looks like a table row (starts with a recognized label, optionally
    after a leading "N."/"N)" number), the blank line was just a
    within-table separator -- skip it and keep going. Only if the next real
    line does NOT look like a table row (e.g. a signature's "Thanks," or
    "Regards,") is this treated as the genuine end of the table. A
    signature glued directly onto the last row with NO blank line at all is
    the one pattern this doesn't catch -- accepted as the safer trade-off,
    since it fails by visibly appending extra text to a field (noticeable)
    rather than silently discarding real fields (invisible and worse for
    money-bearing data).
    """
    schema = schema or load_field_schema()
    labels = _known_labels(schema)

    def _looks_like_table_row(candidate_line):
        m = re.match(r"^(\d+)[.\)]?\s+(.*)$", candidate_line)
        rest = m.group(2) if m else candidate_line
        return any(rest.lower().startswith(lbl.lower()) for lbl in labels)

    lines = text.splitlines()
    n = len(lines)
    rows = []
    i = 0
    while i < n:
        line = lines[i].strip()

        if not line:
            if rows:
                j = i + 1
                while j < n and not lines[j].strip():
                    j += 1
                if j < n and _looks_like_table_row(lines[j].strip()):
                    i += 1
                    continue  # just a separator between two real rows -- keep going
                break  # genuinely the end of the table
            i += 1
            continue

        if re.match(r"^sl\.?\s*no\.?\s+particulars\s+details$", line, re.IGNORECASE):
            i += 1
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
        elif rows:
            rows[-1][2] = (rows[-1][2] + " " + rest).strip()
        # else: stray line before any row has matched yet (greeting, etc.) — ignore

        i += 1

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
        rows = _extract_rows_from_html(source, schema)
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

        # Trust the Sl No match ONLY if the label actually recorded at that
        # number looks like one of THIS field's own aliases. Real bug found
        # 2026-08-13: a test email numbered "Currency" as row 11 and "Total
        # Order Value" as row 12 -- reversed from the schema (Currency=12,
        # Total Order Value=11). Blindly trusting the number alone (the old
        # behavior) silently swapped the two fields' values -- Currency
        # ended up with the amount ("150000"), Total Order Value ended up
        # with "USD". Cross-checking the label text against this field's
        # aliases before trusting the number means a mis-numbered-but-
        # correctly-labeled row still resolves correctly by falling through
        # to the label-based lookup below instead of trusting a wrong number.
        sl_no_label_matches = (
            sl_no is not None and sl_no in by_sl_no
            and any(by_sl_no[sl_no][0].lower().startswith(alias) for alias in aliases)
        )
        if sl_no_label_matches:
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
        elif ftype == "currency_code":
            data[key], currency_warning = _normalize_currency(raw_value)
            if currency_warning:
                warnings.append(currency_warning)
        else:
            data[key] = _clean(raw_value)

    if data.get("master_project_id") in (None, ""):
        warnings.append(
            "Master Project ID is blank — MIS validation (Project ID lookup) is out of scope "
            "for this phase, but flagging so it isn't silently lost."
        )

    return data, warnings


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
