"""
main.py

Single always-on entry point for the whole backend. Runs BOTH halves of the
pipeline inside ONE process, as two independent background tasks:

  1. The Invoice API (scripts/api_server.py's FastAPI app) -- serves
     POST /invoice/api/v1/invoice-generation for the frontend. Appends a row
     to the configured tracker (Google Sheet or local Excel, per
     config/email_server_config.json's "tracker_source").
  2. The draft/poll loop (src/server/email_server.py's run_once()) -- scans
     the tracker every `poll_interval_seconds` for rows whose "Email
     Drafted" flag isn't True, renders the PDF, drafts the Gmail email, and
     flips the flag.

Why one process instead of two terminals: these two jobs are still fully
independent (they only ever talk to each other through the tracker, exactly
as designed in email_server.py/api_server.py's own docstrings) -- this file
just gives them one shared lifecycle. Start `main.py`, both are live; stop
it, both stop. A crash in one task's cycle is caught and logged without
touching the other (the draft loop already does this per-tab; the API
task's own request handlers already catch their own errors into HTTP 4xx/5xx
responses -- see api_server.py).

Run:
    python main.py                  # both, forever (default host 0.0.0.0:5000)
    python main.py --api-only       # just the API
    python main.py --draft-only     # just the draft/poll loop
    python main.py --interval 30    # override poll_interval_seconds
    python main.py --port 8080      # override the API port

This does NOT replace running scripts/api_server.py or
src/server/email_server.py standalone -- both still work exactly as before
for isolated testing/debugging. This file is the "run everything, always"
entry point for actual day-to-day use.
"""
import argparse
import asyncio
import logging
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
SRC_DIR = BASE / "src"
for _p in (str(BASE), str(SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import uvicorn

from utils.env_loader import load_env_file
import server.email_server as email_server  # noqa: E402  (module import only -- no logging.basicConfig side effect at import time, safe to import before setup_logging())

logger = logging.getLogger("main")


async def draft_loop(cfg, interval, stop_event):
    """Runs email_server.run_once() forever on `interval`-second spacing.
    Each cycle runs in a worker thread (run_once does blocking IMAP/Sheets
    I/O) so it never blocks the API's event loop. Mirrors email_server.py's
    own while-loop exactly: an ImapAuthError just logs and retries next
    interval; any other exception is logged and swallowed so one bad cycle
    never kills the loop or the process."""
    logger.info("Draft loop starting (interval=%ss)", interval)
    while not stop_event.is_set():
        try:
            await asyncio.to_thread(email_server.run_once, cfg)
        except email_server.ImapAuthError as e:
            logger.error("Draft loop: auth error, will retry next interval: %s", e)
        except Exception as e:
            logger.exception("Draft loop: unexpected error: %s", e)
        try:
            await asyncio.wait_for(stop_event.wait(), timeout=interval)
        except asyncio.TimeoutError:
            pass  # normal case: interval elapsed, loop again
    logger.info("Draft loop stopped.")


async def run_api(host, port):
    """Imports the FastAPI app lazily, AFTER setup_logging() has already
    attached our handlers to the root logger -- api_server.py calls
    logging.basicConfig() at import time, which is a harmless no-op once a
    handler already exists, so this ordering keeps every log line (API
    and draft loop alike) flowing into the same configured log file instead
    of api_server.py's import silently winning the race and leaving the
    draft loop's file handler never attached."""
    from scripts.api_server import app as api_app

    config = uvicorn.Config(api_app, host=host, port=port, log_level="info")
    server_obj = uvicorn.Server(config)
    await server_obj.serve()


async def main_async(args):
    load_env_file()
    cfg = email_server.load_config()
    email_server.setup_logging(cfg["log_file"])
    interval = args.interval or cfg["poll_interval_seconds"]

    logger.info(
        "main.py starting -- api=%s draft=%s interval=%ss",
        not args.draft_only, not args.api_only, interval,
    )

    stop_event = asyncio.Event()
    tasks = {}
    if not args.draft_only:
        tasks["api"] = asyncio.create_task(run_api(args.host, args.port), name="api")
    if not args.api_only:
        tasks["draft"] = asyncio.create_task(draft_loop(cfg, interval, stop_event), name="draft")

    if not tasks:
        raise SystemExit("Nothing to run -- pass at most one of --api-only/--draft-only, not both.")

    # Joint lifecycle: if EITHER task ends (e.g. uvicorn caught Ctrl+C/SIGTERM
    # and shut down cleanly, or a task died unexpectedly), stop the other one
    # too and exit -- rather than leaving an orphaned half-running process.
    done, pending = await asyncio.wait(tasks.values(), return_when=asyncio.FIRST_COMPLETED)
    stop_event.set()
    for t in pending:
        t.cancel()
    if pending:
        await asyncio.gather(*pending, return_exceptions=True)
    for t in done:
        exc = t.exception()
        if exc:
            logger.error("Task %r ended with an error: %s", t.get_name(), exc)
    logger.info("main.py stopped.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--api-only", action="store_true", help="Run only the Invoice API")
    ap.add_argument("--draft-only", action="store_true", help="Run only the draft/poll loop")
    ap.add_argument("--interval", type=int, default=None, help="Override poll_interval_seconds")
    ap.add_argument("--host", default="0.0.0.0", help="API bind host (default 0.0.0.0)")
    ap.add_argument("--port", type=int, default=5000, help="API bind port (default 5000)")
    args = ap.parse_args()
    if args.api_only and args.draft_only:
        raise SystemExit("--api-only and --draft-only are mutually exclusive.")

    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        print("Stopped by user (Ctrl+C).")


if __name__ == "__main__":
    main()
