#!/usr/bin/env bash
# deploy.sh -- switch the running invoice-automation service to a given
# version folder: stops whatever instance is currently running, then
# starts the requested one.
#
# Usage (run from /application/mis_api_py_node):
#   bash deploy.sh v2
#
# Expects the version's code to already be cloned at:
#   /application/mis_api_py_node/<version>/<repo-folder>/
# with src/main.py inside it, and a working venv already created there
# (<repo-folder>/venv/bin/python). This script does NOT clone code, create
# the venv, or install dependencies -- only stops the old process and
# starts the new one. Do that setup first (git clone + venv + pip install)
# before running this.

set -euo pipefail

BASE_DIR="/application/mis_api_py_node"
LOG_DIR="$BASE_DIR/logs"
PID_FILE="$BASE_DIR/current.pid"

VERSION="${1:-}"
if [ -z "$VERSION" ]; then
  echo "Usage: bash deploy.sh <version>   (e.g. bash deploy.sh v2)"
  exit 1
fi

VERSION_DIR="$BASE_DIR/$VERSION"
if [ ! -d "$VERSION_DIR" ]; then
  echo "ERROR: $VERSION_DIR does not exist. Clone the code there first:"
  echo "  mkdir -p $VERSION_DIR && cd $VERSION_DIR && git clone <repo-url> ."
  exit 1
fi

# Find src/main.py under this version folder, regardless of the cloned
# repo folder's exact name -- APP_DIR is two levels up from main.py
# (main.py -> src/ -> APP_DIR).
MAIN_PY=$(find "$VERSION_DIR" -maxdepth 4 -type f -path "*/src/main.py" | head -n 1)
if [ -z "$MAIN_PY" ]; then
  echo "ERROR: couldn't find src/main.py anywhere under $VERSION_DIR"
  exit 1
fi
APP_DIR=$(dirname "$(dirname "$MAIN_PY")")
echo "Found app at: $APP_DIR"

if [ ! -x "$APP_DIR/venv/bin/python" ]; then
  echo "ERROR: no venv at $APP_DIR/venv -- set it up first:"
  echo "  cd $APP_DIR"
  echo "  python3 -m venv venv || {"
  echo "    python3 -m venv --without-pip venv && source venv/bin/activate"
  echo "    curl -sS https://bootstrap.pypa.io/get-pip.py -o get-pip.py && python get-pip.py"
  echo "  }"
  echo "  source venv/bin/activate && pip install -r requirements.txt"
  exit 1
fi

if [ ! -f "$APP_DIR/.env" ] || [ ! -f "$APP_DIR/credentials/service_account.json" ]; then
  echo "ERROR: $APP_DIR is missing .env or credentials/service_account.json"
  echo "(these are gitignored -- upload them into this version's folder via SFTP first)"
  exit 1
fi

# --- Stop whatever is currently running ---
STOPPED_ANY=false

if [ -f "$PID_FILE" ]; then
  OLD_PID=$(cat "$PID_FILE")
  if kill -0 "$OLD_PID" 2>/dev/null; then
    echo "Stopping previous run (PID $OLD_PID, tracked in $PID_FILE)..."
    kill "$OLD_PID" 2>/dev/null || true
    STOPPED_ANY=true
  fi
  rm -f "$PID_FILE"
fi

# Fallback: catch anything matching src/main.py that the pidfile didn't
# know about (e.g. started manually, or the pidfile is stale after a
# server restart). The [s] trick keeps this grep from matching itself.
for PID in $(ps aux | grep '[s]rc/main\.py' | awk '{print $2}'); do
  if kill -0 "$PID" 2>/dev/null; then
    echo "Stopping stray process (PID $PID)..."
    kill "$PID" 2>/dev/null || true
    STOPPED_ANY=true
  fi
done

if [ "$STOPPED_ANY" = true ]; then
  echo "Waiting for port 8162 to free up..."
  for i in $(seq 1 10); do
    if ! (exec 3<>/dev/tcp/127.0.0.1/8162) 2>/dev/null; then
      break
    fi
    exec 3>&- 2>/dev/null || true
    sleep 1
  done
else
  echo "No previous instance was running."
fi

# --- Start the new version ---
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/app_${VERSION}.log"

cd "$APP_DIR"
nohup "$APP_DIR/venv/bin/python" src/main.py > "$LOG_FILE" 2>&1 &
NEW_PID=$!
disown

echo "$NEW_PID" > "$PID_FILE"
ln -sf "$LOG_FILE" "$LOG_DIR/current.log"

sleep 2
if kill -0 "$NEW_PID" 2>/dev/null; then
  echo ""
  echo "Started $VERSION -- PID $NEW_PID"
  echo "App dir: $APP_DIR"
  echo "Log:     $LOG_FILE   (tail -f $LOG_DIR/current.log)"
  echo "Health:  curl http://localhost:8162/health"
else
  echo "ERROR: process died immediately after starting -- check $LOG_FILE"
  exit 1
fi
