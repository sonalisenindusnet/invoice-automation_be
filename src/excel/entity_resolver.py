"""
entity_resolver.py

Decides WHICH tab (USA / UK / Poland 2026 — the only tabs this is wired up
to use right now) a new invoice request belongs in, based on the client
company name.

CURRENT LIVE ROUTING (as of the 2026-08-14 API/Google-Sheets migration):
the frontend passes `entity` explicitly on every
`POST /invoice/api/v1/invoice-generation` call (see `scripts/api_server.py`),
so `append_invoice()` is normally called with `entity_key` already set and
none of the lookup logic in this file runs at all for live traffic.

`resolve_entity_for_company()` below still exists as the fallback used by
`append_invoice()` when no `entity_key` is supplied — in practice today
that's only the `append_invoice_to_excel.py` CLI run manually (`--entity`
omitted). Two-step lookup, in order:
  1. config/company_tab_map.json — the file you maintain by hand. Exact
     match (case/spacing/punctuation-insensitive) against "company_name".
  2. Fallback: scan every known tab's company-name column in the actual
     tracker for an existing row with that exact company name. If found,
     the match is written BACK into company_tab_map.json automatically
     (with a note that it was auto-detected), so the next lookup for that
     same client resolves instantly from step 1.

If neither step finds a match, this returns None — the caller must NOT
guess a tab for money. See append_invoice_to_excel.py's "entity_unresolved"
status.

Note: India and Kar Ventures are intentionally NOT in ENTITY_SCHEMA_PATHS
below — those are GST/India-specific and out of scope for now (foreign
clients only). A client that resolves to one of those would come back
unresolved rather than routed anywhere. Tell me if/when you want India
support added back — it's a small change (add the tab back to this dict).

(Removed 2026-08-19: an earlier `resolve_entity_from_email_text()` /
"Tab: <name>" first-line-of-email convention that briefly sat between the
company-name approach above and the current explicit-`entity`-field API.
Confirmed unreferenced anywhere in the codebase before removal — the live
server no longer reads emails at all, and the API takes `entity` directly.)
"""
import json
import re
from datetime import date
from pathlib import Path

# 2026-08-14: see tracker_io.py's module docstring -- transparently opens
# either a local .xlsx path (unchanged) or the live Google Sheet, depending
# on what's passed in.
from utils.tracker_io import load_tracker_with_retry as load_workbook_with_retry

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "config"
COMPANY_MAP_PATH = CONFIG_DIR / "company_tab_map.json"
TABS_DIR = CONFIG_DIR / "tabs"

ENTITY_SCHEMA_PATHS = {
    "usa": TABS_DIR / "usa.json",
    "uk": TABS_DIR / "uk.json",
    "poland": TABS_DIR / "poland.json",
    # "india" / "kar_ventures" / "singapore" deliberately absent — foreign
    # clients only, for now.
}

_LEGAL_SUFFIXES = [
    "private limited", "pvt ltd", "pte ltd", "sp z o o", "sarl", "llc",
    "limited", "ltd", "inc", "corp", "corporation", "co",
]


def load_schema(entity_key):
    path = ENTITY_SCHEMA_PATHS.get(entity_key)
    if path is None or not path.exists():
        raise ValueError(f"No Excel tab schema configured for entity '{entity_key}'")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_company_map():
    if not COMPANY_MAP_PATH.exists():
        return {"companies": [], "entity_tabs": {}}
    with open(COMPANY_MAP_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_company_map(company_map):
    with open(COMPANY_MAP_PATH, "w", encoding="utf-8") as f:
        json.dump(company_map, f, indent=2, ensure_ascii=False)
        f.write("\n")


def _normalize(name):
    """Lowercase, strip punctuation, collapse whitespace, and drop a trailing
    legal-entity suffix — so 'Accuride International Ltd.' and
    'accuride international ltd' resolve to the same key, without being so
    loose it risks matching the wrong client."""
    if not name:
        return ""
    n = name.lower().strip()
    n = re.sub(r"[.,()]", "", n)
    n = re.sub(r"\s+", " ", n).strip()
    for suffix in _LEGAL_SUFFIXES:
        pattern = r"\s+" + re.escape(suffix) + r"$"
        n = re.sub(pattern, "", n)
    return n.strip()


def resolve_entity_for_company(company_name, xlsx_path=None):
    """
    Returns (entity_key, resolution) where resolution is one of:
      "mapped"      — found in company_tab_map.json
      "auto_detected" — found by scanning an existing tab; also just got
                        added to company_tab_map.json for next time
      None entity_key, resolution "unresolved" — not found anywhere;
                        caller must flag this, never guess.
    """
    if not company_name or not company_name.strip():
        return None, "unresolved"

    target = _normalize(company_name)
    company_map = _load_company_map()

    for entry in company_map.get("companies", []):
        if _normalize(entry.get("company_name", "")) == target:
            return entry["entity_key"], "mapped"

    if xlsx_path is None:
        return None, "unresolved"

    try:
        wb = load_workbook_with_retry(xlsx_path, data_only=True)
    except FileNotFoundError:
        return None, "unresolved"

    for entity_key, schema_path in ENTITY_SCHEMA_PATHS.items():
        if not schema_path.exists():
            continue
        with open(schema_path, "r", encoding="utf-8") as f:
            schema = json.load(f)
        sheet_name = schema["sheet_name"]
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        cols = schema["columns"]
        company_key = schema.get("company_column_key", "client_company")
        col_idx = next((i for i, c in enumerate(cols) if c["key"] == company_key), None)
        if col_idx is None:
            continue
        start = schema.get("header_row_index", 1) + 1
        for row in ws.iter_rows(min_row=start, max_row=ws.max_row, values_only=True):
            if row is None or col_idx >= len(row):
                continue
            existing_name = row[col_idx]
            if existing_name and _normalize(str(existing_name)) == target:
                company_map.setdefault("companies", []).append({
                    "company_name": company_name.strip(),
                    "entity_key": entity_key,
                    "note": f"auto-detected from an existing '{sheet_name}' row on {date.today().isoformat()}",
                })
                _save_company_map(company_map)
                return entity_key, "auto_detected"

    return None, "unresolved"
