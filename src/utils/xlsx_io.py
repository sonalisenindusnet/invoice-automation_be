"""
xlsx_io.py

Wraps openpyxl.load_workbook() with a short retry loop. Added 2026-08-12
after a real live failure during testing:

    xml.etree.ElementTree.ParseError: not well-formed (invalid token):
    line 85, column 1342297

...thrown while Pass 2 was scanning the Poland tab. Root cause, confirmed
right after by pulling the file fresh from the device: Excel had
"Invoice Traker.xlsx" open (its own autosave/recalculation) at the exact
moment this server's poll cycle tried to read the same file, and caught it
mid-write -- a "torn read". The bytes on disk were briefly inconsistent
(not actually corrupted data, just an incomplete write), which a strict XML
parser reports as invalid. Loading the SAME file again immediately after
worked perfectly -- no data was lost, no sheet was actually damaged, and
Excel could open the file normally throughout.

This is a DIFFERENT issue from the older, separate, known concern that
openpyxl round-trip saves silently drop cached formula VALUES (not the
formulas) on formula-driven tabs (Kar Ventures / SBI Yet to raise / India)
-- that's why schema-only edits to this file are done via surgical XML
patches instead of openpyxl. This retry wrapper does not change that
policy; it only smooths over a transient read-vs-write race on the exact
same file, which becomes more likely the more often Excel is open on it
while the server is also running.

Root fix for the user: don't keep Excel open on the tracker while the
server is running, if avoidable. This retry is a safety net for exactly
the moments that still happen anyway (autosave, background recalculation),
NOT a substitute for that.
"""
import logging
import threading
import time
import xml.etree.ElementTree as ET
import zipfile

import openpyxl

logger = logging.getLogger("email_server.xlsx_io")

# Only retry on errors that look like "the file was mid-write", never on
# genuine problems (FileNotFoundError, permission errors, a truly missing
# sheet, etc.) -- those should surface immediately, not get masked by a
# retry loop.
TRANSIENT_ERRORS = (zipfile.BadZipFile, ET.ParseError)

# Shared by every part of this process that does a read-modify-write cycle
# against the tracker file (save_api's append, the draft poller's row
# scan-and-update) -- serializes them so two such cycles running at nearly
# the same moment can't both read the "before" state and then clobber each
# other's write. Does not protect against a human editing the file in Excel
# at the same instant; load_workbook_with_retry above is what smooths over
# that.
TRACKER_LOCK = threading.Lock()


def load_workbook_with_retry(path, data_only=False, retries=4, delay_seconds=0.75):
    """Same contract as openpyxl.load_workbook(path, data_only=...), except
    a transient "file was mid-write" error is retried a few times (with a
    short pause) before giving up. After all retries are exhausted, the
    ORIGINAL exception is re-raised -- this never silently swallows a real,
    persistent problem; it only smooths over a brief race."""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            return openpyxl.load_workbook(path, data_only=data_only)
        except TRANSIENT_ERRORS as e:
            last_err = e
            if attempt < retries:
                logger.warning(
                    "  load_workbook attempt %d/%d hit a transient error (%s) -- "
                    "likely Excel mid-save on %r; retrying in %.2fs",
                    attempt, retries, e, str(path), delay_seconds,
                )
                time.sleep(delay_seconds)
            else:
                logger.error(
                    "  load_workbook FAILED after %d attempts (%s) -- giving up for this cycle: %r",
                    retries, e, str(path),
                )
    raise last_err
