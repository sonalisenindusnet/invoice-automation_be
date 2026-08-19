# Invoice Automation

This project has two independent jobs:

1. **Invoice API** — accepts invoice data and appends it to the configured
   Google Sheet.
2. **Draft server** — polls the USA, UK, and Poland 2026 tabs. For each row
   whose **Email Drafted** cell is explicitly `FALSE`, it generates the PDF,
   creates a Gmail draft with the PDF attached, then sets **Email Drafted**
   to `TRUE`.

The project does not read, parse, or mark emails in the Gmail inbox.

For normal day-to-day use, run both together with `main.py` (see "Run
everything" below) — they still only ever talk to each other through the
tracker, exactly as before; `main.py` just gives them one process/one
lifecycle instead of two terminals to babysit. `scripts/api_server.py` and
`src/server/email_server.py` still work standalone too, for isolated
testing.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Create `.env` from `.env.example` and set `EMAIL_ADDRESS` and
`EMAIL_APP_PASSWORD`. The Google service-account file configured in
`config/email_server_config.json` must have access to the target Google Sheet.

## Run everything (recommended for normal use)

```powershell
python main.py
```

Starts the API (`http://0.0.0.0:5000`) and the draft/poll loop together in
one process, one Ctrl+C stops both. Options:

```powershell
python main.py --api-only        # just the API
python main.py --draft-only      # just the draft/poll loop
python main.py --interval 30     # override poll_interval_seconds
python main.py --port 8080       # override the API port
```

## Run the API only

The API listens on `http://localhost:5000`. Its health check is
`GET /health`; invoice creation is
`POST /invoice/api/v1/invoice-generation`.

## Run the draft server only

Run one safe test cycle:

```powershell
python src\server\email_server.py --once
```

Run continuously (default: every 60 seconds):

```powershell
python src\server\email_server.py
```

For a row to be processed, set **Email Drafted** to the boolean `FALSE`.
Blank cells are skipped intentionally, so historical rows cannot generate
drafts by accident. The generated PDF and runtime log are written under
`output/`.

## Project layout

```text
main.py                               Run-everything entry point (API + draft loop, one process)
scripts/api_server.py                 API entry point (standalone)
src/server/email_server.py            Google Sheets → PDF → Gmail-draft server (standalone)
src/excel/                            Sheet read/write and entity schemas
src/mailer/                           Gmail draft composition and IMAP client
src/pdf/generate_invoice_pdf_intl.py  USA, UK, and Poland PDF renderer
src/mis/                              MIS project lookup
src/utils/                            environment, Google Sheets, and Excel helpers
config/                               runtime and entity configuration
```
