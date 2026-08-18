"""
mis_api.py

Calls the real MIS "project info" lookup API — added 2026-08-12, replacing
the earlier "bypass MIS validation" stance now that a real endpoint exists:

    POST {base_url}{endpoint_path}
    x-api-key: <MIS_API_KEY, read from .env>
    Content-Type: application/json
    body: {"project_id": "<pf_id from the CP's email>"}

Auth changed 2026-08-12: the API moved from a Bearer JWT to a static
x-api-key header. MIS_API_TOKEN / the old "Authorization: Bearer" header
are gone -- replaced by MIS_API_KEY and the "x-api-key" header below.

This is called from append_invoice_to_excel.py's append_invoice(), right
after the row is confirmed not a duplicate and pf_id is known — BEFORE the
row is built and written, so the whole row (email-parsed fields + whatever
the MIS API adds) goes into Excel in one shot, per explicit instruction:
"After call this api receive some more fields which is also append in the
excel, thus we can append the row in one go."

Key refresh: per the same standing instruction as when this was a Bearer
token ("token each time I will update in env manually, you dont need to
work here"), there is deliberately NO auto-refresh logic here. You update
MIS_API_KEY in .env by hand whenever it changes/rotates, and restart the
server. A call made with a missing/wrong key is treated exactly like any
other MIS API failure below — it never blocks the append.

Field mapping (MIS_FIELD_MAP below) is deliberately EMPTY right now — the
real /api/projectinfo response shape hasn't been captured yet (this
sandbox can't reach that internal IP to test it directly). Every call's
raw response is written to <output_dir>/debug_last_mis_response.json
(same pattern as email_server.py's debug_last_email_body dump) so it can
be read directly and used to fill in MIS_FIELD_MAP once the endpoint has
actually returned something real — at that point, add the matching column
to config/tabs/*.json at the same time and it'll start getting written.
"""
import json
import os
import urllib.error
import urllib.request
from pathlib import Path

CONFIG_PATH = Path(__file__).resolve().parent.parent.parent / "config" / "mis_api_config.json"

# Maps a key in the MIS API's JSON response -> our internal row key (the
# same keys used everywhere else, e.g. build_row()'s canonical `values`
# dict in append_invoice_to_excel.py). Empty until a real response is
# captured via the debug file described above — filling this in is the
# ONLY code change needed to start writing extra MIS-sourced columns; add
# the matching column (same key) to config/tabs/*.json at the same time.
MIS_FIELD_MAP = {
    # "client_name": "mis_client_name",
    # "approved_order_value": "mis_approved_order_value",
    # "status": "mis_project_status",
}


def load_mis_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def fetch_project_info(pf_id, debug_path=None, cfg=None):
    """Looks up one project by pf_id. NEVER raises — any failure (disabled
    in config, missing API key, no pf_id, timeout, connection error, non-200,
    bad JSON) comes back as verified=False with `error` set, so the caller
    can append the row anyway with mis_verification_done=False, per
    explicit instruction (2026-08-12): "if the mis api failed or timeout or
    any error happen then append the row but this column should be false.
    so accountant manually check the project ID."

    Returns:
        {"verified": bool, "fields": {<mapped row keys>: value, ...},
         "raw": <parsed JSON body, or raw text if not JSON, or None>,
         "error": str or None}
    """
    cfg = cfg or load_mis_config()
    if not cfg.get("enabled", True):
        result = {"verified": False, "fields": {}, "raw": None, "error": "MIS check disabled in config"}
        _write_debug(debug_path, pf_id, result)
        return result

    if not pf_id:
        result = {"verified": False, "fields": {}, "raw": None, "error": "No PF ID to look up"}
        _write_debug(debug_path, pf_id, result)
        return result

    api_key = os.environ.get("MIS_API_KEY")
    if not api_key:
        result = {"verified": False, "fields": {}, "raw": None,
                   "error": "MIS_API_KEY not set in .env -- skipping MIS check for this row"}
        _write_debug(debug_path, pf_id, result)
        return result

    url = cfg["base_url"].rstrip("/") + cfg["endpoint_path"]
    body = json.dumps({"project_id": pf_id}).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"x-api-key": api_key, "Content-Type": "application/json"},
    )

    try:
        with urllib.request.urlopen(req, timeout=cfg.get("timeout_seconds", 15)) as resp:
            raw_text = resp.read().decode("utf-8", errors="replace")
            status = getattr(resp, "status", None) or resp.getcode()
    except urllib.error.HTTPError as e:
        raw_text = e.read().decode("utf-8", errors="replace") if e.fp else ""
        result = {"verified": False, "fields": {}, "raw": raw_text,
                   "error": f"HTTP {e.code} from MIS API: {raw_text or e.reason}"}
        _write_debug(debug_path, pf_id, result)
        return result
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        result = {"verified": False, "fields": {}, "raw": None,
                   "error": f"MIS API unreachable or timed out: {e}"}
        _write_debug(debug_path, pf_id, result)
        return result

    try:
        parsed = json.loads(raw_text) if raw_text else None
    except json.JSONDecodeError:
        result = {"verified": False, "fields": {}, "raw": raw_text,
                   "error": "MIS API returned a non-JSON response"}
        _write_debug(debug_path, pf_id, result)
        return result

    if status != 200 or parsed is None:
        result = {"verified": False, "fields": {}, "raw": parsed,
                   "error": f"MIS API returned status {status}"}
        _write_debug(debug_path, pf_id, result)
        return result

    mapped = {}
    source = parsed.get("data", parsed) if isinstance(parsed, dict) else {}
    if isinstance(source, dict):
        for api_key, row_key in MIS_FIELD_MAP.items():
            if api_key in source:
                mapped[row_key] = source[api_key]

    result = {"verified": True, "fields": mapped, "raw": parsed, "error": None}
    _write_debug(debug_path, pf_id, result)
    return result


def _write_debug(debug_path, pf_id, result):
    """Best-effort only -- a debug-file write failure must never break the
    real append. Overwrites the same file each time (mirrors
    debug_last_email_body.* in email_server.py) so the LATEST call's raw
    response is always what you find there."""
    if not debug_path:
        return
    try:
        Path(debug_path).write_text(
            json.dumps({"pf_id": pf_id, **result}, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
    except OSError:
        pass
