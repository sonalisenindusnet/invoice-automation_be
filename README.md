# Invoice Automation

This project has two independent jobs, run together by `src/main.py` in
one process:

1. **Save API** — accepts invoice data from the frontend (a normal
   single-resource invoice, or a dedicated multi-resource invoice) and
   appends it as a new row into the live Google Sheet tracker.
2. **Draft poll loop** — scans the tracker's USA/UK/Poland/Singapore tabs
   for rows an accountant has marked **Reviewed**, and for each one not yet
   drafted: renders the invoice PDF, composes the email body, and uploads
   it as a real Gmail draft (never sent automatically) with the PDF
   attached.

They only ever talk to each other through the tracker itself (a live
Google Sheet, not a local file — see `src/utils/tracker_io.py`) — `main.py`
just gives them one shared process/lifecycle instead of two terminals to
babysit.

See `RUNNING_COMMANDS.md` for exact setup and run commands, and
`.env.example` for every environment variable this project reads.

## Code layout

```text
src/
  main.py             Run-everything entry point (save API + draft loop, one process)
  api/
    routes.py          FastAPI app + all HTTP endpoints (HTTP layer only)
  models/
    invoice_models.py  Pydantic request schemas for the save API
  services/            Business logic -- no HTTP concerns
    excel_writer.py             Tracker row read/write, invoice numbering
    dedicated_invoice.py        Resource email grouping/summing for dedicated invoices
    tax_calculator.py           GST/VAT calculation (single source of truth)
    generate_invoice_pdf_intl.py  Per-entity invoice PDF rendering
    poller.py                   Draft poll loop (scans tracker, drafts Gmail emails)
    email_composer.py           Tracker row -> {to, cc, subject, body}
    gmail_imap.py                Gmail draft creation over IMAP
    llm_drafter.py                LLM (Gemini) email body drafting, with template fallback
  utils/
    tracker_io.py        Google Sheets-backed tracker adapter (duck-types openpyxl)
    xlsx_io.py            Shared TRACKER_LOCK
    env_loader.py          Loads .env into the process
config/
  save_api_config.json      Save API host/port
  draft_poller_config.json  IMAP host/port/timeout, drafts folder override
  tabs/<entity>.json         Each tab's real column layout + invoice-number series
  entities/<entity>.json     Each entity's PDF layout, tax rule, bank details
credentials/
  tracker_sheet.json    Single source of truth for which Google Sheet is the
                        live tracker + which service-account credential opens it
  service_account.json  The Google service-account credential itself
scripts/                One-off Singapore tab setup/maintenance scripts
output/                 Generated invoice PDFs land here
```

The project does not read, parse, or mark emails in the Gmail inbox — the
draft loop only ever creates new Drafts, never sends or touches existing
mail.

## Setup

See `RUNNING_COMMANDS.md` for full setup (virtualenv, `requirements.txt`,
`.env`) and run commands for both Git Bash and PowerShell.

## Run everything (recommended for normal use)

```bash
python src/main.py
```

Starts the save API (`http://0.0.0.0:8162` by default — see
`config/save_api_config.json`) and the draft poll loop together in one
process; `Ctrl+C` stops both. Options:

```bash
python src/main.py --api-only     # just the save API
python src/main.py --draft-only   # just the draft poll loop
python src/main.py --interval 30  # override DRAFT_POLL_INTERVAL_SECONDS (default 300)
```

## Endpoints

- `GET /health`
- `POST /invoice/api/v1/invoice-generation` — normal, single-line invoice
- `POST /invoice/api/v1/dedicated-invoice-generation` — multi-resource
  invoice; resources sharing the same email are summed into one merged
  line before the client ever sees it (see
  `services/dedicated_invoice.py`)
- `POST /invoice/api/v1/mis-verification` — flips a row's MIS Verified flag

See `RUNNING_COMMANDS.md` for a worked `curl` example and the full field
reference.
