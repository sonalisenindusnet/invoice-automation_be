"""
main.py

Entry point: runs both background jobs in one process.
  1. The save API (save_api/app.py) -- receives an invoice from the
     frontend and saves it into the Excel tracker.
  2. The draft poll loop (draft_mailer/poller.py) -- scans the tracker
     for rows marked "Reviewed" and drafts a Gmail email (with the invoice
     PDF attached) for each one that hasn't been drafted yet.

These stay independent -- they only ever interact through the tracker file
itself (and its shared lock, see utils/xlsx_io.TRACKER_LOCK) -- this file
just gives them one shared lifecycle: start main.py, both are live; stop
it (Ctrl+C), both stop.

Run:
    python src/main.py                # both, forever
    python src/main.py --api-only     # just the save API
    python src/main.py --draft-only   # just the draft poll loop
    python src/main.py --interval 30  # override DRAFT_POLL_INTERVAL_SECONDS
"""
import argparse
import asyncio
import logging
import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from utils.env_loader import load_env_file

# Load .env into os.environ before anything else runs -- EMAIL_ADDRESS/
# EMAIL_APP_PASSWORD (gmail_imap.py), GEMINI_API_KEY (llm_drafter.py),
# PAYMENT_DUE_DAYS (excel_writer.py), and DRAFT_POLL_INTERVAL_SECONDS
# (poller.py) are all read straight from os.environ with no dotenv
# machinery of their own -- without this call, a real .env file on disk is
# silently never actually loaded into the process.
load_env_file()

import uvicorn

from save_api.app import app as save_api_app, load_config as load_save_api_config
from draft_mailer import poller
from draft_mailer.gmail_imap import ImapAuthError

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("main")


async def run_api(cfg):
    config = uvicorn.Config(save_api_app, host=cfg.get("host", "0.0.0.0"), port=cfg.get("port", 5000), log_level="info")
    server = uvicorn.Server(config)
    await server.serve()


async def draft_loop(cfg, interval, stop_event):
    """Runs poller.run_once() forever on `interval`-second spacing. Each
    cycle runs in a worker thread (blocking IMAP/Excel I/O) so it never
    blocks the API's event loop. An IMAP auth error just logs and retries
    next interval; any other exception is logged and swallowed so one bad
    cycle never kills the loop or the process."""
    logger.info("Draft loop starting (interval=%ss)", interval)
    while not stop_event.is_set():
        try:
            await asyncio.to_thread(poller.run_once, cfg)
        except ImapAuthError as e:
            logger.error("Draft loop: auth error, will retry next interval: %s", e)
        except Exception as e:
            logger.exception("Draft loop: unexpected error: %s", e)
        try:
            await asyncio.wait_for(stop_event.wait(), timeout=interval)
        except asyncio.TimeoutError:
            pass  # normal case: interval elapsed, loop again
    logger.info("Draft loop stopped.")


async def main_async(args):
    stop_event = asyncio.Event()
    tasks = {}

    if not args.draft_only:
        api_cfg = load_save_api_config()
        tasks["api"] = asyncio.create_task(run_api(api_cfg), name="api")

    if not args.api_only:
        draft_cfg = poller.load_config()
        interval = args.interval or poller.poll_interval_seconds()
        tasks["draft"] = asyncio.create_task(draft_loop(draft_cfg, interval, stop_event), name="draft")

    if not tasks:
        raise SystemExit("Nothing to run -- pass at most one of --api-only/--draft-only, not both.")

    logger.info("main.py starting -- api=%s draft=%s", "api" in tasks, "draft" in tasks)

    # Joint lifecycle: if either task ends, stop the other and exit, rather
    # than leaving an orphaned half-running process.
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
    ap.add_argument("--api-only", action="store_true", help="Run only the save API")
    ap.add_argument("--draft-only", action="store_true", help="Run only the draft poll loop")
    ap.add_argument("--interval", type=int, default=None, help="Override DRAFT_POLL_INTERVAL_SECONDS")
    args = ap.parse_args()
    if args.api_only and args.draft_only:
        raise SystemExit("--api-only and --draft-only are mutually exclusive.")

    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        print("Stopped by user (Ctrl+C).")


if __name__ == "__main__":
    main()
