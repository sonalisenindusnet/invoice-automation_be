"""
xlsx_io.py

Local-.xlsx support has been removed -- the real tracker is always the
live Google Sheet now (see utils/tracker_io.py). All that's left here is
`TRACKER_LOCK`, a bare threading.Lock() with no storage-backend coupling,
shared by every part of this process that does a read-modify-write cycle
against the tracker (save_api's append/MIS-verify update, the draft
poller's row scan-and-update) -- it serializes them so two such cycles
running at nearly the same moment can't both read the "before" state and
then clobber each other's write. (Google Sheets itself already supports
true concurrent multi-writer access with no file locks -- this lock is
only about this one process's own features never racing each other.)
"""
import threading

TRACKER_LOCK = threading.Lock()
