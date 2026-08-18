"""
env_loader.py

Tiny, dependency-free .env file loader. Reads KEY=VALUE lines from a
`.env` file in the project root and sets them into os.environ — but only
for keys not already set, so a real environment variable you set yourself
in PowerShell always wins over the file.

No pip install needed for this — deliberately kept to stdlib only.
"""
import os
from pathlib import Path

DEFAULT_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"


def load_env_file(path=None):
    """Loads KEY=VALUE pairs from a .env file into os.environ.
    Returns the dict of keys it found (values not echoed to logs anywhere)."""
    path = Path(path) if path else DEFAULT_ENV_PATH
    if not path.exists():
        return {}

    found = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        if not key:
            continue
        found[key] = value
        if key not in os.environ:  # real env vars always take precedence
            os.environ[key] = value
    return found
