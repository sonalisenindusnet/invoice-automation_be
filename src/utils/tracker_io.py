"""
tracker_io.py

Single entry point every read/write in save_api/excel_writer.py and
draft_mailer/poller.py goes through to reach the actual Invoice Tracker
storage -- the live Google Sheet (sheet id + service-account credentials
come from config; see tracker_ref_from_config below). Local-.xlsx support
has been removed entirely: `tracker_ref` is always the
{"type": "google_sheets", "sheet_id": ..., "service_account_json": ...}
dict form now, and load_tracker_with_retry() raises immediately if it's
handed anything else.

WHAT THIS MODULE DOES: opens the live Google Sheet via the Sheets API
(gspread) and returns a WorkbookHandle that duck-types the small subset of
openpyxl's Workbook interface the rest of the codebase actually uses
(.sheetnames, wb[name], .create_sheet(name), .save()) -- so none of the
actual business logic in excel_writer.py/poller.py (invoice numbering,
MIS-verify lookup, Created-At stamping, Reviewed-row scanning) needs to
know it's talking to Sheets rather than a local file.

Google Sheets has no separate "cached formula value" concept the way a
local .xlsx can -- it always live-evaluates formulas on read, so
`data_only` is accepted here for call-site compatibility but has no effect.
"""
import logging
import time
from pathlib import Path

logger = logging.getLogger("email_server.tracker_io")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

_client_cache = {}


def _get_gspread_client(service_account_json):
    if service_account_json not in _client_cache:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_file(service_account_json, scopes=SCOPES)
        _client_cache[service_account_json] = gspread.authorize(creds)
    return _client_cache[service_account_json]


def load_tracker_with_retry(tracker_ref, data_only=False, retries=4, delay_seconds=1.5):
    """Opens the live Google Sheet described by `tracker_ref`
    ({"type": "google_sheets", "sheet_id": ..., "service_account_json": ...})
    and returns a WorkbookHandle. Transient errors (network blips, Sheets
    API rate limiting -- HTTP 429/5xx) are retried a few times before the
    original exception is re-raised. `data_only` is accepted for call-site
    compatibility but has no effect (Sheets always live-evaluates formulas).

    Raises ValueError immediately (no retry) if `tracker_ref` isn't the
    expected dict shape -- local-.xlsx tracker support has been removed."""
    if not isinstance(tracker_ref, dict) or tracker_ref.get("type") != "google_sheets":
        raise ValueError(
            f"tracker_ref must be a google_sheets dict, got {tracker_ref!r} -- "
            "local .xlsx tracker support has been removed."
        )

    import gspread

    last_err = None
    for attempt in range(1, retries + 1):
        try:
            gc = _get_gspread_client(tracker_ref["service_account_json"])
            sh = gc.open_by_key(tracker_ref["sheet_id"])
            return WorkbookHandle(sh)
        except (gspread.exceptions.APIError, TimeoutError, ConnectionError) as e:
            last_err = e
            if attempt < retries:
                logger.warning(
                    "  Google Sheets open attempt %d/%d failed (%s) -- retrying in %.2fs",
                    attempt, retries, e, delay_seconds,
                )
                time.sleep(delay_seconds)
            else:
                logger.error("  Google Sheets open FAILED after %d attempts: %s", retries, e)
    raise last_err


def tracker_ref_from_config(cfg, config_dir):
    """Resolves a config dict into the `tracker_ref` shape
    load_tracker_with_retry() expects -- always the google_sheets dict now.
    Both save_api and draft_mailer resolve their (separately-loaded,
    deliberately duplicated) configs through this one shared function so
    the sheet-id/service-account resolution logic never drifts between the
    two. Raises KeyError with a clear message if the config is missing
    either required key."""
    if "google_sheet_id" not in cfg or "google_service_account_json" not in cfg:
        raise KeyError(
            "Config is missing 'google_sheet_id' and/or 'google_service_account_json' -- "
            "the tracker is always the live Google Sheet now, local .xlsx support has been removed."
        )
    sa_path = Path(cfg["google_service_account_json"])
    if not sa_path.is_absolute():
        sa_path = (config_dir / sa_path).resolve()
    return {
        "type": "google_sheets",
        "sheet_id": cfg["google_sheet_id"],
        "service_account_json": str(sa_path),
    }


def save_tracker(wb, tracker_ref):
    """Flushes every pending write on `wb` to the live Sheet. `tracker_ref`
    isn't actually needed by WorkbookHandle.save() (which takes no
    arguments), but is accepted here for call-site symmetry."""
    wb.save()


class _CellRef:
    """Mimics openpyxl's Cell just enough for `.value` reads (the only
    attribute any caller in this codebase touches)."""
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class WorksheetHandle:
    """Duck-types the small subset of openpyxl's Worksheet interface used
    across this codebase (.max_row, .iter_rows(), .cell(), .append()),
    backed by a live gspread Worksheet.

    Reads: the whole sheet's values are pulled ONCE per handle (lazily, on
    first access) via get_all_values() and cached for the handle's
    lifetime -- exactly matching openpyxl's "load whole sheet into memory,
    then read from memory" model, so next_invoice_no()'s scan followed by
    next_data_row()'s scan (both inside one append_invoice() call) costs
    exactly one API read, not two.

    Writes: `.cell(row, col, value=...)` stages the write in memory (also
    updating this handle's own read cache, so a write followed by a read
    within the same call sees the pending value immediately -- matching
    openpyxl's in-memory-until-.save() model). Nothing actually reaches
    Google until `.flush()` is called (by WorkbookHandle.save()), at which
    point every pending cell across this worksheet goes out in ONE batched
    API call, regardless of how many separate .cell() calls staged them."""

    def __init__(self, worksheet):
        self._ws = worksheet
        self._values = None
        self._pending = {}  # (row, col) -> value

    def _load(self):
        if self._values is None:
            self._values = self._ws.get_all_values()
        return self._values

    @property
    def max_row(self):
        vals = self._load()
        pending_max = max((r for (r, _c) in self._pending), default=0)
        return max(len(vals), pending_max)

    @property
    def max_column(self):
        vals = self._load()
        existing_max = max((len(row) for row in vals), default=0)
        pending_max = max((c for (_r, c) in self._pending), default=0)
        return max(existing_max, pending_max)

    def _row_width(self, row):
        vals = self._load()
        existing = len(vals[row - 1]) if 0 <= row - 1 < len(vals) else 0
        pending = max((c for (r, c) in self._pending if r == row), default=0)
        return max(existing, pending)

    def _cell_value(self, row, col):
        if (row, col) in self._pending:
            return self._pending[(row, col)]
        vals = self._load()
        if 0 <= row - 1 < len(vals) and 0 <= col - 1 < len(vals[row - 1]):
            v = vals[row - 1][col - 1]
            return v if v != "" else None
        return None

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        upper = max_row if max_row is not None else self.max_row
        for r in range(min_row, upper + 1):
            width = self._row_width(r)
            if width == 0:
                yield None
                continue
            yield tuple(self._cell_value(r, c) for c in range(1, width + 1))

    def cell(self, row, column, value=None):
        if value is not None:
            self._pending[(row, column)] = value
            return _CellRef(value)
        return _CellRef(self._cell_value(row, column))

    def append(self, values):
        """Only ever hit by ensure_sheet() for a brand-new tab -- never
        triggered for USA/UK/Poland 2026 in real use since they already
        exist on the live Sheet."""
        next_row = self.max_row + 1
        for i, v in enumerate(values, start=1):
            self._pending[(next_row, i)] = v

    def flush(self):
        if not self._pending:
            return
        from gspread.utils import rowcol_to_a1

        data = [
            {"range": rowcol_to_a1(r, c), "values": [[v]]}
            for (r, c), v in self._pending.items()
        ]
        self._ws.batch_update(data, value_input_option="USER_ENTERED")
        self._pending.clear()
        self._values = None  # force a fresh read next time this handle is used


class WorkbookHandle:
    """Duck-types the small subset of openpyxl's Workbook interface used
    across this codebase (.sheetnames, wb[name], .create_sheet(name),
    .save(path)), backed by a live gspread Spreadsheet.

    Returns the SAME WorksheetHandle on repeated wb[name] access within one
    call, so a function that does e.g. `ws = ensure_sheet(wb, schema)` and
    then reads/writes `ws` several times keeps one consistent read cache +
    pending-write buffer, exactly like holding one openpyxl worksheet
    object throughout a call."""

    def __init__(self, spreadsheet):
        self._sh = spreadsheet
        self._open = {}

    @property
    def sheetnames(self):
        return [w.title for w in self._sh.worksheets()]

    def __getitem__(self, name):
        if name not in self._open:
            self._open[name] = WorksheetHandle(self._sh.worksheet(name))
        return self._open[name]

    def create_sheet(self, name):
        ws = self._sh.add_worksheet(title=name, rows=1000, cols=30)
        handle = WorksheetHandle(ws)
        self._open[name] = handle
        return handle

    def save(self, path=None):
        """No separate "save" step exists for Sheets the way it does for a
        local .xlsx -- every write already lives on Google's servers the
        moment .flush() runs. `path` is accepted (and ignored) purely so
        call sites written as `wb.save(out_path or xlsx_path)` don't need
        to change."""
        for handle in self._open.values():
            handle.flush()
