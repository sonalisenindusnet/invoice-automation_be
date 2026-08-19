# Invoice Automation

This project has two active entry points:

1. **Invoice API** — accepts invoice data and appends it to the configured
   Google Sheet.
2. **Draft server** — polls the USA, UK, and Poland 2026 tabs. For each row
   whose **Email Drafted** cell is explicitly `FALSE`, it generates the PDF,
   creates a Gmail draft with the PDF attached, then sets **Email Drafted**
   to `TRUE`.

The project does not read, parse, or mark emails in the Gmail inbox.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Create `.env` from `.env.example` and set `EMAIL_ADDRESS` and
`EMAIL_APP_PASSWORD`. The Google service-account file configured in
`config/email_server_config.json` must have access to the target Google Sheet.

## Run the API

```powershell
python scripts\api_server.py
```

The API listens on `http://localhost:5000`. Its health check is
`GET /health`; invoice creation is
`POST /invoice/api/v1/invoice-generation`.

## Run the draft server

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
scripts/api_server.py                 API entry point
src/server/email_server.py            Google Sheets → PDF → Gmail-draft server
src/excel/                            Sheet read/write and entity schemas
src/mailer/                           Gmail draft composition and IMAP client
src/pdf/generate_invoice_pdf_intl.py  USA, UK, and Poland PDF renderer
src/mis/                              MIS project lookup
src/utils/                            environment, Google Sheets, and Excel helpers
config/                               runtime and entity configuration
```
