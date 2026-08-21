# Run Commands

## Setup (one time)

From the project root (`E:\Test Project\invoice\invoice-automation_be`):

**Git Bash / MINGW64** (what you're actually running — the prompt reading
`MINGW64` means bash, not PowerShell; `.ps1` scripts don't run there):
```bash
python -m venv .venv
source .venv/Scripts/activate
pip install -r requirements.txt
```

**PowerShell** (if you ever run this from a `powershell.exe`/`pwsh` prompt
instead — same venv, different activation script):
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Either way, re-run `pip install -r requirements.txt` any time
`requirements.txt` changes — activating the venv alone doesn't install
anything new.

## Credentials (.env)

Copy `.env.example` to `.env` (if you haven't already) and fill in:

```
EMAIL_ADDRESS=your-gmail-address@gmail.com
EMAIL_APP_PASSWORD=your-16-character-app-password
GEMINI_API_KEY=your-gemini-key
PAYMENT_DUE_DAYS=7
DRAFT_POLL_INTERVAL_SECONDS=300
```

- `EMAIL_ADDRESS` / `EMAIL_APP_PASSWORD` are only needed for the draft
  loop (creating Gmail drafts over IMAP) — generate the app password at
  myaccount.google.com/apppasswords, not your real Google password.
- `GEMINI_API_KEY` is only needed for the LLM-drafted email body. If it's
  missing or the call fails, the draft loop falls back to a plain
  template automatically — it never blocks a draft from being created.
- `PAYMENT_DUE_DAYS` is optional — how many days after the invoice date
  the auto-filled "Payment Due Date" column is set to. Defaults to 7 if
  unset or not a valid whole number. Only affects the save API.
- `DRAFT_POLL_INTERVAL_SECONDS` is optional — how often (in seconds) the
  draft loop re-scans the tracker. Defaults to 300 (5 minutes) if unset or
  not a valid whole number. `--interval` on the command line overrides
  both.
- The save API itself needs neither the email nor Gemini credentials.

## Run everything (recommended)

```bash
python src/main.py
```

Starts the save API (`http://0.0.0.0:5000` by default — see
`config/save_api_config.json`) and the draft poll loop together in one
process. `Ctrl+C` stops both.

Options:
```bash
python src/main.py --api-only        # just the save API
python src/main.py --draft-only      # just the draft poll loop
python src/main.py --interval 30     # override DRAFT_POLL_INTERVAL_SECONDS (default 300)
```

Sanity-check it's actually up, from a second terminal:
```bash
curl http://localhost:5000/health
# -> {"success":true,"status":"UP"}
```

## Try the save endpoint directly

```bash
curl -X POST http://localhost:5000/invoice/api/v1/invoice-generation \
  -H "Content-Type: application/json" \
  -d '{
        "pfId": "PF-TEST-1",
        "clientName": "Test Client",
        "invoiceDescription": "Services for August26",
        "contactPersonName": "Jane Doe",
        "clientMailTo": "jane@example.com",
        "intCcMailId": "accounts@intglobal.com",
        "currency": "USD",
        "invoiceValue": "1000",
        "entity": "usa"
      }'
```

This appends a row into the real "USA" tab, assigns the next invoice
number in that tab's series, stamps a "Created At" timestamp, and
auto-fills **Payment Status** (`Not Paid`), **Payment Due Date**
(invoice date + `PAYMENT_DUE_DAYS`), and **Review Status**
(`Pending Review`).

## Draft loop trigger

The draft loop only picks up a row once its **Review Status** cell is
set to exactly `Reviewed` by hand in the tracker (every new row starts as
`Pending Review`). Once that's set, the
next poll cycle (within `DRAFT_POLL_INTERVAL_SECONDS`, default 300s)
renders the PDF, drafts the email, uploads it as a real Gmail draft, and
flips that row's **Email Drafted** to `True`.

## Where things land

- Tracker: the live Google Sheet identified by `google_sheet_id` in
  `config/save_api_config.json` and `config/draft_poller_config.json`
  (both point at the same sheet), reached via the service-account
  credentials at `google_service_account_json` (default
  `credentials/service_account.json`).
- Generated invoice PDFs: `output/`
- Gmail drafts: uploaded straight to the monitored account's Drafts
  folder — nothing is saved locally as `.eml`

## Config files

- `config/save_api_config.json` — Google Sheet id/credentials path, API
  host/port.
- `config/draft_poller_config.json` — Google Sheet id/credentials path
  (same sheet), IMAP host/port/timeout, optional `drafts_folder_override`.
- `config/tabs/usa.json` / `uk.json` / `poland.json` — each tab's real
  column layout and invoice-number series. Keep these in sync with the
  real sheet's actual columns — see the architecture notes if you're
  changing sheet structure.

`Ctrl+C` stops any of the continuously-running commands above.
