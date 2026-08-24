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
TAX_RATE_SINGAPORE_LOCAL=0.09
TAX_RATE_SINGAPORE_FOREIGN=0.0
TAX_RATE_UK_LOCAL=0.20
TAX_RATE_UK_FOREIGN=0.0
TAX_RATE_POLAND_LOCAL=0.0
TAX_RATE_POLAND_FOREIGN=0.0
TAX_RATE_USA_LOCAL=0.0
TAX_RATE_USA_FOREIGN=0.0
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
- `TAX_RATE_<ENTITY>_LOCAL` / `TAX_RATE_<ENTITY>_FOREIGN` (all 8 above are
  optional — every default shown matches the actual tax rule already in
  effect) control `src/tax/tax_calculator.py`, the single source of truth
  for invoice tax. Tax depends on whether the **client's own country**
  matches the entity the invoice is raised from: `LOCAL` is the rate
  charged when it matches (e.g. a Singapore client on a Singapore
  invoice), `FOREIGN` is the rate otherwise. Only Singapore (9%/0%) and UK
  (20%/0%) actually vary today — Poland and USA are 0% either way, per
  explicit instruction, but still have their own env vars in case that
  ever changes. The tax name shown (GST for Singapore, VAT for the other
  three) is fixed by each country's own tax law, not configurable.
- The save API itself needs neither the email nor Gemini credentials, but
  does use the `TAX_RATE_*` vars (to report the tax breakdown in its
  response).

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
        "companyLocation": "United States",
        "raisedByEmail": "atanub@intglobal.com",
        "currency": "USD",
        "invoiceValue": "1000",
        "entity": "usa"
      }'
```

`raisedByEmail` is who internally raised/requested the invoice (added
2026-08-24) — stored in the tracker's own **"Invoice Advised By"** column
(`requested_by`). Distinct from `clientMailTo`, which is the CLIENT's
email address (used for the "Client Mail To" column and, later, as the
drafted email's recipient) — the two used to be conflated (that column was
populated from `clientMailTo` as a placeholder before this field existed).
Optional — an empty/missing value just leaves that column blank.

The client's own country/location field. **Either spelling is accepted**:
`companyLocation` (camelCase, matching this request's other fields) or
`company_location` (snake_case). This was originally specified as
snake_case only, but a live request on 2026-08-24 showed camelCase being
sent instead — since an unrecognized JSON key is silently ignored (not an
error), that request's country was quietly dropped and treated as
"foreign" with no visible error, until the blank-value WARNING log caught
it. The API now accepts both spellings so it works either way; if both
are somehow sent in the same request, `company_location` wins. It's what
`src/tax/tax_calculator.py` compares against the entity to decide the
LOCAL vs FOREIGN tax rate (see the `.env` section above); for
UK/Poland/Singapore it's also stored in that tab's existing "Country"
column. Matching tolerates punctuation, case, and compound values
("U.K.", "England, UK") — see that module's own docstring for exactly
what's recognized. Safe to omit entirely — an empty/missing value is
treated as "foreign" (the lower rate), never accidentally triggers the
higher local rate, but the save API logs a WARNING when it's blank so a
frontend-side issue doesn't fail silently.

This appends a row into the real tab, assigns the next invoice number in
that tab's series, stamps a "Created At" timestamp, and auto-fills
**Payment Status** (`Not Paid`), **Payment Due Date** (invoice date +
`PAYMENT_DUE_DAYS`), and **Review Status** (`Pending Review`). Tax is
computed from `companyLocation`/`company_location` and the entity (see
above) and **is baked into the tracker itself**: the tab's own **"Total
Amount"** column is set to the POST-tax grand total (subtotal + tax —
matching the column's own real-world meaning; Poland's actual header is
literally "Total Amount (Including VAT)"), and where the tab has its own
tax column (UK/Poland's "VAT (GBP)" today), that column gets the tax
amount. USA (and Singapore, until it gets a GST column) has no separate
tax column — fine, since its tax is always 0 either way. The response
also echoes the same breakdown (`tax_name`, `tax_rate`, `tax_amount`,
`subtotal` — the pre-tax figure the frontend originally sent,
`total_with_tax` — same as the sheet's own "Total Amount"). The SAME
figures are what actually show on the PDF and in the drafted email
later, once a human marks the row "Reviewed" and MIS-verifies it — the
draft loop reconstructs them from what's already saved in the sheet
rather than recomputing from scratch, so they can never disagree with
what an accountant sees in the tracker.

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
