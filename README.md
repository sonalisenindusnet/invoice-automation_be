# Invoice Email Automation — Demo

Automates the first slice of the OneSpace flow: CP sends an **"Invoice Summary"**
email → **update the Invoice Tracker Excel** (deterministic, no LLM) → **read
that Excel row back out and draft the invoice PDF + client email** (this is
where the LLM comes in). Tally sync and bank-statement reconciliation are
deliberately out of scope for this phase.

**Scope right now: foreign clients only — USA, UK, and Poland 2026. No tax
is applied to any of them.** India isn't wired up (see "Multi-entity
routing" below for exactly what that means and how small the change would
be to bring it back if you ever need it).

## Pipeline

```
CP email  --[parse_invoice_summary.py]-->  structured fields
          --[entity_resolver.py]-->        WHICH tab does this client belong to?
                                             USA / UK / Poland 2026
                                             (config/company_tab_map.json, or scan existing
                                              tabs — never a guess; see "Multi-entity routing")
          --[append_invoice_to_excel.py]--> new row in THAT tab
                                             (tax calc — none of these three have any tax —
                                              dedupe, invoice numbering; plain rule-based
                                              code, no LLM)
          --[draft_email_from_excel_row.py]--> generates the invoice PDF using that
                                             entity's own template (generate_invoice_pdf_intl.py)
                                             and the To/CC/Subject/Body draft
                                             (this is the "LLM reads the excel" step — I read
                                              the row and use judgment on wording, rather than
                                              a fixed script)
          --[mark_email_sent.py]-->        tries to write the outcome BACK into that same row —
                                             none of USA/UK/Poland 2026 have an EMAIL STATUS
                                             column in your real tracker, so this is currently
                                             always a safe no-op
```

`demo/run_demo.py` runs this once, by hand, on sample data — good for
testing the logic. **`src/server/email_server.py` is the actual always-on
server**: it watches a real inbox and runs this same pipeline automatically.
See "Running the server" below.

(Code is organized under `src/` by what it does — `src/parsing/`,
`src/excel/`, `src/pdf/`, `src/mailer/`, `src/server/`, `src/utils/` — see
"Folder structure" below for the full layout and why.)

## Running the server

`email_server.py` polls an inbox on an interval and, for every unread email
whose Subject contains `"Invoice Summary"` (configurable), runs the full
pipeline automatically and drops a ready-to-review draft into that account's
Drafts folder — no manual step in between.

It connects over plain IMAP with a Google **App Password**, not the Claude
Gmail connector — this is the "locally connected" route that sidesteps the
account-mismatch problem entirely, since it's a direct login to whichever
Gmail account you choose, same as configuring Outlook or Thunderbird.

### One-time setup

1. **Turn on 2-Step Verification** on the Gmail account you want this to
   monitor (Google Account → Security). App Passwords require it.
2. **Generate an App Password**: myaccount.google.com/apppasswords → create
   one (name it e.g. "Invoice Automation") → copy the 16-character password.
3. **Provide the credentials** — two ways, pick one:

   **Option A — `.env` file (persists, so you don't retype it every time):**
   Copy `.env.example` (in the project root) to a new file named `.env` in
   the same folder, then edit `.env` and fill in your real values:
   ```
   EMAIL_ADDRESS=the-account@yourdomain.com
   EMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
   ```
   `email_server.py` reads this automatically. `.env` is listed in
   `.gitignore` and is never read, uploaded, or synced by me — it's yours
   alone, on your machine. **Never paste its contents into this chat.**

   **Option B — environment variables (this shell session only):**
   ```
   $env:EMAIL_ADDRESS="the-account@yourdomain.com"
   $env:EMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"
   ```
   A real environment variable always overrides whatever is in `.env`, if
   you ever want to test with different credentials temporarily.
4. `config/email_server_config.json`'s `tracker_xlsx_path` now points at
   `../Invoice Traker.xlsx` — your real tracker file, sitting at the project
   root next to `.env` and `README.md`. Every matching email edits that exact
   file directly (opens it, appends the row, saves back to the same path) —
   no copy is created. If you ever move or rename that file, update this one
   path and nothing else needs to change.

### Running it

Run these from the project's root folder (the one with `.env`, `config/`,
`src/`, `demo/` in it) — the paths below are relative to that:

```
python src/server/email_server.py --once        # single pass — do this FIRST to test
python src/server/email_server.py                # runs forever, polling every 60s (configurable)
python src/server/email_server.py --interval 30  # override the poll interval
```

(On Windows that's `python src\server\email_server.py ...` — either slash
direction works in PowerShell.)

Leave that terminal window open and it runs continuously, checking the inbox
every `poll_interval_seconds`. To actually run as a background service
(survives closing the terminal / machine restart), the simplest options on
Windows are:
- A **Scheduled Task** set to run `python src\server\email_server.py` at logon, with
  "repeat task every X minutes" — or just have it call the script once and
  let the script's own internal loop keep running
- **NSSM** (Non-Sucking Service Manager) to wrap it as a proper Windows
  service, if you want it to survive without a logged-in session

### What it does per matching email

1. Parses the email body (handles both HTML and plain-text CP emails)
2. Figures out which tab the named client belongs to — USA, UK, or Poland
   2026 (see "Multi-entity routing" below)
3. Appends a row to that tab (no tax for any of the three, invoice
   numbering, duplicate check — a repeat email with the same PF ID + amount
   is skipped, not double-logged)
4. Generates the invoice PDF using that entity's own template
5. Builds the client email (To/CC/Subject/Body) — no tax lines, since none
   of these three entities have any right now
6. **Appends it directly into the account's Drafts folder as a real Gmail
   draft, PDF attached** — this is a plain IMAP APPEND with the `\Draft`
   flag, not an actual send; nothing goes out until a human opens Gmail and
   hits Send
7. Tries to write `EMAIL STATUS = "Draft Created"` + sent date/to/cc back
   into that Excel row — none of USA/UK/Poland 2026 have this column in
   your real tracker, so this step currently always reports "not
   supported" and changes nothing
8. Marks the source email read and records its Message-ID in
   `output/processed_emails.json`, so it's never processed twice — even if
   you restart the server

If step 2 can't determine the client's tab (not in
`config/company_tab_map.json` and not found sitting in any existing tab),
the email is logged as `entity_unresolved` and left **unread/unprocessed**
— nothing is appended anywhere. Add the company to
`config/company_tab_map.json` and the next poll picks it up automatically;
no data is ever written to a guessed tab.

Every run is logged to `output/email_server.log` (and printed to the
console), and one failing email is caught and logged without stopping the
loop or skipping the rest of that poll's batch.

### Honest caveat

I built and unit-tested this pipeline end-to-end using a mocked IMAP
connection (verified: parsing, Excel append, PDF generation, and the exact
MIME draft bytes it would upload) — but I could **not** test it against a
real Gmail server, because this cloud sandbox's network doesn't allow
outbound connections to `imap.gmail.com`. That part only runs on your
machine. Please run `python src/server/email_server.py --once` first and send me
whatever it prints (or the contents of `output/email_server.log`) if
anything errors — IMAP folder-naming and auth quirks vary enough by account
that I'd expect at least one round of fixes.

## How it maps to the diagram

| Diagram step | This code |
|---|---|
| 1. Invoice Request Email (CP → Accounts inbox) | Gmail MCP `search_threads` (subject: "Invoice Summary") + `get_thread` |
| 2. AI Agent captures email & auto-fills Excel | `parse_invoice_summary.py` + `entity_resolver.py` (which tab?) + `append_invoice_to_excel.py` |
| 4. AI Agent identifies new entries & validates data | duplicate check inside `append_invoice_to_excel.py` |
| 5. Generate invoice PDF + personalized email draft | `draft_email_from_excel_row.py` + `generate_invoice_pdf_intl.py` |
| 6. Store in Draft Workspace for Accounts review | Gmail MCP `create_draft` (pending connector) |
| 9. Delivery status logged for audit/search | `mark_email_sent.py` → currently a no-op for all three active tabs (none of them have an EMAIL STATUS column) |
| 3, 10 (MIS check, Tally sync, bank reconciliation) | **Not built yet** — next phase |

## Multi-entity routing — which tab does a client go into?

A CP email can be for a client billed by any of your foreign entities — USA,
UK, or Poland — so the server has to work out which tab a client belongs to
before appending anything, using `src/excel/entity_resolver.py`:

1. **`config/company_tab_map.json`** — the file you maintain. A flat list of
   `{company_name, entity_key}`, checked first with an exact match
   (case/punctuation/spacing-insensitive, and it also strips a trailing
   "Ltd"/"Inc"/"Pvt Ltd"/etc. so small formatting differences don't matter).
   This is the file to edit when you take on a new client — not the Python
   code.
2. **Fallback — scan the existing tabs.** If a company isn't in the mapping
   file, the code looks for that exact company name already sitting in a
   row on the USA, UK, or Poland 2026 tab. If found, it uses that tab **and
   writes the match back into `company_tab_map.json` automatically** — so
   the next invoice for that same client resolves instantly from step 1, no
   scanning needed.
3. **If neither finds a match**, the invoice comes back `entity_unresolved`
   and is **not appended anywhere**. In the live server this means the
   source email is left unread/unprocessed (not stuck — just retried on the
   next poll) so that once you add the company to the mapping file, it goes
   through automatically without you having to resend anything. This never
   guesses a tab for money.

Where each tab's real column layout lives:

| Tab | Schema file | Tax rule |
|---|---|---|
| USA | `config/tabs/usa.json` | none |
| UK | `config/tabs/uk.json` | VAT 20% |
| Poland 2026 | `config/tabs/poland.json` | VAT 0% (edit `tax.rate` if a client needs otherwise) |

Every column in these files matches your real tracker's actual headers
exactly (verified directly against your real "Invoice Traker.xlsx") —
nothing here creates a new column or restructures an existing tab; it only
appends rows using the columns already there.

**India and Kar Ventures are intentionally not wired up right now** — per
your instruction, foreign clients only, no GST anywhere. If an Indian client
ever needs to go through this pipeline, tell me and I'll add GST support
(CGST/SGST) plus an India tab back in — the underlying tax code
(`compute_tax()` in `append_invoice_to_excel.py`) already knows how to
handle a GST-style split, it's just unused by any active schema today, so
this would be a small, contained change rather than a rebuild.

**Singapore has no tab yet** in your tracker either. If a CP email names a
Singapore client, it comes back `entity_unresolved` until you either tell
me to add a Singapore tab or map that client to one of the existing three
tabs instead.

**Two honest limitations from routing to these tabs**, both because the
CP's "Invoice Summary" email doesn't carry every field a full invoice needs:
- USA's real invoice shows one row per resource with its own cost; the CP
  email only gives one total, so the generated PDF uses a single line item
  for the whole amount rather than a resource-by-resource breakdown.
- UK and Poland 2026 have no "Invoice Description" column in your tracker,
  so that text only exists in memory right when an email is processed — the
  live server and demo both draft the PDF/email from that in-memory row
  immediately, so this doesn't lose anything in normal use. It only matters
  if you manually re-draft an old UK/Poland invoice later purely from the
  sheet (via `draft_email_from_excel_row.py`'s CLI) — that re-draft falls
  back to a generic "Services" description, since the sheet itself has
  nowhere to keep the original wording.

## Known gaps, called out honestly

- **GSTIN, HSN, Client Type, Business Mode, POS** aren't in the CP's
  "Invoice Summary" email, so they're left blank in the new row rather than
  guessed. If you have a client → GSTIN/HSN lookup, tell me and I'll add it
  as a config file the same way `field_schema.json` works.
- **"Invoice Requested By"** — `email_server.py` fills this from the source
  email's `From` header automatically. The manual scripts (`run_demo.py`,
  `append_invoice_to_excel.py`'s CLI) still take it as a `--requested-by`
  parameter since there's no live email to read it from.
- Opening the workbook with openpyxl prints `UserWarning: Unknown extension
  is not supported and will be removed` — this is Excel-internal metadata
  (not your data) that openpyxl doesn't preserve on save. Sheet names, all
  values, and formatting I checked came through intact; flagging this so
  it isn't a surprise, and I'd suggest opening the saved file once in Excel
  directly to confirm before trusting it for anything beyond this demo.
- International templates (USA/UK/Poland) are wired into the live pipeline
  via the routing described above — Singapore still has no tab (see
  "Multi-entity routing"), and the USA/UK/Poland caveats there (no
  per-resource breakdown, no stored description) are worth a read.

## USA / UK / Poland invoice templates

You sent 4 real invoice PDFs (USA, UK, Singapore, Poland) and each entity
has its own legal name, tax rule, table shape, and bank-details format. Only
USA, UK, and Poland are wired into the live pipeline right now (see
"Multi-entity routing"); Singapore's template exists but has no tab to
route into yet.

| Entity | Title | Currency | Table shape | Tax | Amount in words |
|---|---|---|---|---|---|
| USA | Invoice | USD | one row per resource, "Total Monthly Billing" + "Sub-Total" rows | none | Yes |
| UK | Invoice | GBP | single multi-line description row | Vat 20%, shown **after** Sub-Total, final total unlabelled | No |
| Poland | Invoice | USD (bills a Canadian client) | single multi-line description row, has PO No./PO Dt. fields | VAT 0%, shown **before** Sub-Total | Yes (replaces a "Total" label) |
| Singapore *(template only — no tab yet)* | **Tax Invoice** | SGD | single multi-line description row | GST 9%, verbose labels ("Total amount payable excluding/including GST") | No |

This is implemented as:

- `config/entities/usa.json`, `uk.json`, `singapore.json`, `poland.json` —
  one config per entity: legal name, footer address, CIN/VAT/NIP/GST Reg No,
  bank details (and PayPal for Singapore), tax rate/position/labels, table
  columns, invoice-number prefix. Edit these directly for any correction —
  nothing is hardcoded in the render code.
- `src/pdf/generate_invoice_pdf_intl.py` —
  `render_international_invoice(entity_key, row, out_path)`. Handles both
  table shapes (per-resource for USA, single description row for the other
  three), each entity's own tax positioning, and an amount-in-words line for
  USA/Poland (plain integer words, no currency-name suffix like "Dollars",
  matching what's in the real PDFs).
- Tested against reconstructed data matching all 4 real invoices exactly
  (same invoice numbers, client names, amounts, bank details) and visually
  checked — all four render correctly.

**Wired into the live server** — the client company name in the CP email is
looked up in `config/company_tab_map.json` (falling back to scanning the
existing tabs), and that lookup decides both the Excel tab AND which PDF
renderer runs. Singapore is the one exception — no tab exists for it yet,
so a Singapore client still comes back `entity_unresolved` rather than
silently using the wrong template.

## Folder structure

Everything lives under one parent project folder. Code is grouped by what it
does, not dumped into one flat `scripts/` folder:

```
invoice-automation/                    <- parent folder, everything below lives here
├── .env                               (you create this — real credentials, gitignored, never sent to me)
├── .env.example                       template for your credentials file — copy to ".env" and fill in
├── .gitignore                         keeps .env, logs, and __pycache__ out of any repo you put this in
├── requirements.txt                   pip install -r requirements.txt
├── README.md
│
├── config/                            all settings and per-entity data — no code
│   ├── field_schema.json              field map for the CP's Invoice Summary email table
│   ├── company_tab_map.json           WHICH client goes to WHICH tab — the file you maintain
│   │                                    (see "Multi-entity routing")
│   ├── email_server_config.json       server behavior: subject filter, poll interval, tracker path
│   │                                    (credentials are NOT here — see "Running the server")
│   ├── tabs/                          Excel column schema for each active destination tab —
│   │   ├── usa.json                    matches your tracker's REAL columns exactly (see
│   │   ├── uk.json                     "Multi-entity routing" for the full table)
│   │   └── poland.json
│   └── entities/                      one file per international entity for PDF rendering —
│       ├── usa.json                    legal name, bank details, tax rule, table shape (this
│       ├── uk.json                     is separate from config/tabs/ above: entities/ is "how
│       ├── singapore.json              the PDF looks", tabs/ is "which Excel columns exist")
│       └── poland.json                 (singapore.json has no matching tabs/ file — no tab yet)
│
├── src/                               all Python source, one subfolder per concern
│   ├── parsing/
│   │   └── parse_invoice_summary.py    email body (HTML/text/dict) -> structured dict
│   ├── excel/
│   │   ├── entity_resolver.py          decides WHICH tab a client belongs to (see
│   │   │                                "Multi-entity routing"); loads that tab's schema
│   │   └── append_invoice_to_excel.py  structured dict -> new row in the resolved tab (tax calc,
│   │                                    dedupe, numbering); also update_email_status()/find_row_index()
│   ├── pdf/
│   │   └── generate_invoice_pdf_intl.py  renders USA/UK/Poland invoice PDFs, driven
│   │                                    entirely by config/entities/*.json
│   ├── mailer/
│   │   ├── gmail_imap.py               raw IMAP: connect (with a timeout — fails fast instead of
│   │   │                                hanging), find matching emails, build/append a Drafts MIME message
│   │   ├── draft_email_from_excel_row.py  turns a row into an invoice PDF + email draft, dispatching
│   │   │                                to the right PDF renderer and tax wording per entity
│   │   └── mark_email_sent.py          writes EMAIL STATUS / SENT DATE / SENT TO / SENT CC back into that
│   │                                    row — a safe no-op today, since none of USA/UK/Poland 2026
│   │                                    have that column in your real tracker
│   ├── server/
│   │   └── email_server.py             the actual always-on server — see "Running the server"
│   └── utils/
│       └── env_loader.py               tiny dependency-free .env file reader
│
├── demo/
│   └── run_demo.py                    manual one-shot run of the whole flow on a sample USA client
│
└── output/                            generated at runtime — nothing here is source code
    ├── tracker_demo.xlsx              a throwaway COPY used only by demo/run_demo.py for safe
    │                                   testing — NOT what the live server writes to (see below)
    ├── invoice_INT-USA-26-27-001.pdf  invoices generated from tracker rows
    ├── draft_preview_INT-USA-26-27-001.eml  human-readable preview of a draft email
    ├── processed_emails.json         (created by email_server.py) Message-IDs already handled
    ├── email_server.log               (created by email_server.py) run log
    └── debug_last_email_body.txt      (created by email_server.py) raw body of the last matching email
```

Each `src/` subfolder groups files that work on the same concern, so
`draft_email_from_excel_row.py` and `mark_email_sent.py` sit together under
`mailer/` (both act on the outgoing email/draft), while `email_server.py` —
the orchestrator that calls into every other subfolder — has its own
`server/` folder. Every module still resolves `config/` and `output/` paths
from its own file location, so it works the same regardless of which folder
you run a command from.

If you're running any single script directly instead of through
`email_server.py`, run it with its full path from the project root, e.g.
`python src/pdf/generate_invoice_pdf_intl.py usa row.json out.pdf` or
`python src/mailer/mark_email_sent.py output/tracker_demo.xlsx --invoice-no ...`
— each script finds its own imports and config regardless of your current
directory.

**A few files from earlier rounds are no longer used and can be deleted
from your project folder** — they were part of an India/GST-specific setup
that's been removed now that the active scope is foreign clients only:
`config/india_tab_schema.json`, `config/tabs/kar_ventures.json`,
`config/company_profile.json`, `src/pdf/generate_invoice_pdf.py`. Nothing
in the current code imports or reads any of these, so removing them is
safe. (If you still have an old flat `scripts/` folder from an even earlier
reorganization, that can go too — everything moved into `src/`/`demo/`.)

## Demo run vs. the live server — which file each one touches

- `demo/run_demo.py` always writes to `output/tracker_demo.xlsx`, a
  throwaway copy — it's for trying out the pipeline without risking your
  real file, and its path is hardcoded, separate from
  `email_server_config.json`.
- `src/server/email_server.py` (the real, always-on server) writes to
  whatever `tracker_xlsx_path` says in `config/email_server_config.json` —
  currently `../Invoice Traker.xlsx`, your real tracker at the project root.
  It edits that file directly, in place, every time a matching email comes
  in. No copy, no separate demo file involved.

```
python3 demo/run_demo.py
```

using a sample "Somax Inc" invoice request (already mapped to the USA tab
in `config/company_tab_map.json`), this:

1. Parsed it
2. Looked up "Somax Inc" and resolved it to the USA tab, then appended one
   row to the USA tab of `output/tracker_demo.xlsx` — no tax, since USA has
   none — with a correctly-computed $14,000 total
3. Read the row back out (from memory, not re-read from the sheet) and
   produced the invoice PDF and draft email
4. Tried to write the outcome back into that same row — reported
   "not_supported" honestly, since the USA tab has no EMAIL STATUS column
   in your real tracker
5. Ran a bonus scenario with a completely unmapped company name, confirming
   it comes back `entity_unresolved` and is never appended anywhere

## What happens once Gmail is connected (your side)

You're connecting a Gmail account to the Gmail MCP connector (search, read
threads, create drafts). Once it's enabled in this chat, the live loop is:

1. `search_threads(query='subject:"Invoice Summary"')` to find the CP's email
2. `get_thread(...)` to pull the body (HTML or plain text — the parser
   handles either)
3. Run `parse_invoice_summary.py` → `entity_resolver.py` →
   `append_invoice_to_excel.py` on that body
4. Run `draft_email_from_excel_row.py` on the new row
5. `create_draft(...)` with the composed To/CC/Subject/Body — **I'll need to
   check the connector's actual `create_draft` schema for attachment support
   once it's live**; if it doesn't accept attachments directly, the fallback
   is a raw-MIME draft with the PDF inlined as a base64 attachment (need to
   confirm which the connector exposes).

Nothing here talks to Gmail as a separate program with its own login — I (the
agent, in this session) call the MCP tools directly once you've connected the
account, using these scripts to do the parsing/generation work in between.

## Open questions for you

- Whether "INT CC mail id" should always be CC'd on the *client* email
  (current assumption, matching the diagram's To/CC step), or is internal-only
- Any real invoice template you'd like the PDF to match, if the USA/UK/Poland
  ones ever need corrections
- Whether Singapore should get a real tab added, or foreign clients stay at
  USA/UK/Poland only for now
