# Invoice Automation Backend — UAT Deployment Guide

Server: `INT-expanse-mgmt` (`35.154.72.144`)
App path: `/application/mis_api_py_node/mis_invoice_automation`
Logs: `/application/mis_api_py_node/logs/app.log`
Port: `8162`

This document is the exact, proven sequence used to deploy and run this service
on the UAT server. Follow it in order for any fresh deployment or redeployment.

---

## 1. Connect

**File transfer (SFTP)** — use for uploading code/config changes:
```
sftp -P 9822 mis_api_py_sftp@35.154.72.144
```
Note: this SFTP account can only browse/write under `/var` (e.g.
`/var/www/html/mis_api_py`) — it **cannot** reach `/application/mis_api_py_node`
directly. That path is only reachable over a real SSH shell session, or via
`git clone`/`cp` run from within that shell (see Step 2).

**Shell access (SSH)** — use for running commands, git, venv, starting the app:
```
ssh mis_api_py_sftp@35.154.72.144 -p 9822
```
(Same account also has an interactive shell, separate from its restricted SFTP scope.)

---

## 2. Get the code onto the server

The working deployment path is `/application/mis_api_py_node`. Since the SFTP
account can't browse there directly, code was brought in via `git clone` run
from inside the SSH session:
```
cd /application/mis_api_py_node
git clone <repo-url> mis_invoice_automation
cd mis_invoice_automation
```

**Important — these files are gitignored and will NOT come through git clone.
Upload them separately via SFTP into this same folder:**
- `.env` (real values — see Step 4)
- `credentials/service_account.json`

Verify both are present before continuing:
```
ls -la
ls -la credentials
```

---

## 3. Confirm Python

```
python3 --version
```
Needs 3.11+. This server has 3.12.3 confirmed working. Always use `python3`
explicitly — plain `python` is not aliased on this server.

---

## 4. Configure `.env`

Copy `.env.example` to `.env` if not already uploaded, and fill in real values:

| Key | Required? | Notes |
|---|---|---|
| `EMAIL_ADDRESS` | Yes | Gmail address the draft-mailer poller uses (IMAP) |
| `EMAIL_APP_PASSWORD` | Yes | Must be a Gmail **App Password**, not the normal login password. Requires 2-Step Verification enabled on that Gmail account, generated at Google Account → Security → App Passwords. A regular password will fail IMAP login with `AUTHENTICATIONFAILED`. |
| `GEMINI_API_KEY` | Optional | If missing, email drafting falls back to a plain template — doesn't block anything. |
| `GEMINI_MODEL` | Optional | Defaults if unset. |
| `MIS_API_KEY` | Not used | Leftover from a removed feature — ignore. |

Double-check for hidden formatting issues if IMAP login fails:
```
cat -A .env | grep EMAIL_APP_PASSWORD
```
(watch for stray quotes/spaces/trailing characters).

---

## 5. Confirm config points to the right things

`config/save_api_config.json` and `config/draft_poller_config.json` should already have:
- `"port": 8162`
- correct `google_sheet_id` and `google_service_account_json` path (relative to the config file's own folder)

No changes needed here for a standard redeploy — just confirm they weren't reverted by a fresh `git clone`.

---

## 6. Create the virtual environment

This server has **no `sudo` access** for this account, and the system Python is
missing `ensurepip` (`apt install python3.12-venv` is not usable here). Use the
no-root fallback:

```
python3 -m venv --without-pip venv
source venv/bin/activate
curl -sS https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python get-pip.py
```

Confirm pip landed inside the venv (not system-wide):
```
which pip
pip --version
```

(If a future server *does* have working `ensurepip`/root access, the simpler
`python3 -m venv venv` alone is enough — try that first before the fallback.)

---

## 7. Install dependencies

```
pip install -r requirements.txt
```

---

## 8. Test run (foreground)

```
python src/main.py
```

Expect to see:
```
INFO:     Uvicorn running on http://0.0.0.0:8162 (Press CTRL+C to quit)
2026-... [INFO] Draft cycle: 0 row(s) drafted
```
No `AUTHENTICATIONFAILED` errors should appear — if they do, revisit Step 4.

In a second terminal/session, confirm the app itself is healthy:
```
curl http://35.154.72.144:8162/health
```
Note: hit `/health` directly — there is **no** `/application/mis_api_py_node`
prefix in the app's own routes; that's just the folder it runs from on disk,
not part of any URL path.

Stop the test run with `Ctrl+C` once confirmed.

---

## 9. Run permanently in the background

No process supervisor (systemd/pm2) is set up on this account yet (needs root
or Node — see "Known limitations" below), so `nohup` is used:

```
cd /application/mis_api_py_node/mis_invoice_automation
source venv/bin/activate
nohup python src/main.py > /application/mis_api_py_node/logs/app.log 2>&1 &
disown
```

Confirm it's running and detached:
```
ps aux | grep main.py
curl http://35.154.72.144:8162/health
```

You can now safely close the SSH session — the app keeps running.

---

## 10. Day-to-day operations

**View live logs:**
```
tail -f /application/mis_api_py_node/logs/app.log
```

**Check if it's running:**
```
ps aux | grep main.py
```

**Stop it:**
```
kill <PID>
```
(get `<PID>` from the `ps aux` output above)

**Restart after a code/config update:**
```
kill <old PID>
cd /application/mis_api_py_node/mis_invoice_automation
git pull            # if updating via git
source venv/bin/activate
nohup python src/main.py > /application/mis_api_py_node/logs/app.log 2>&1 &
disown
```
Remember: `git pull` will never touch `.env` or `credentials/service_account.json`
since both are gitignored — no need to re-upload them unless they actually changed.

---

## Known limitations / open items

- **No auto-restart on crash or server reboot.** `nohup` keeps the process
  alive after the SSH session ends, but won't restart it if the process dies
  or the server reboots. Fixing this properly needs a process supervisor
  (systemd or pm2), which needs either root access or Node.js — ask DevOps.
- **SFTP account is scoped to `/var` only.** It cannot browse or upload
  directly into `/application/mis_api_py_node` — code has to be moved there
  via `git clone`/`cp` run from the SSH shell instead (Step 2). Flagged to
  DevOps as a possible SFTP scope fix.
- **Public URL path stripping unconfirmed.** If `http://35.154.72.144/mis_api_py/health`
  (via whatever reverse proxy sits on port 80) needs to work, that proxy must
  strip its path prefix before forwarding to port 8162 — the app itself only
  ever answers at `/health`, with no prefix. Confirm with DevOps whether this
  proxy layer exists and is configured correctly.
