"""
main.py

Entry point: runs the save-invoice API (save_api/app.py). This is the one
thing src/ is responsible for running -- receive an invoice from the
frontend, save it into the Excel tracker, nothing else.

Run:
    python src/main.py
"""
import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import uvicorn

from save_api.app import app, load_config


def main():
    cfg = load_config()
    uvicorn.run(app, host=cfg.get("host", "0.0.0.0"), port=cfg.get("port", 5000))


if __name__ == "__main__":
    main()
