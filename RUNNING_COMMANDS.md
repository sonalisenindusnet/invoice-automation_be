# Project Run Commands

This file is a quick command reference for running the invoice automation
project from the repository root.

## 1. Create a virtual environment

PowerShell on Windows:

```powershell
python -m venv .venv
```

## 2. Activate the virtual environment

PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run this once in the same window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## 3. Install dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## 4. Configure environment variables

Copy the example file and fill in your Gmail credentials:

```powershell
Copy-Item .env.example .env
```

Then edit `.env` and set:

```text
EMAIL_ADDRESS=your-email@example.com
EMAIL_APP_PASSWORD=your-16-character-app-password
```

You can also set them for the current PowerShell session only:

```powershell
$env:EMAIL_ADDRESS="your-email@example.com"
$env:EMAIL_APP_PASSWORD="your-16-character-app-password"
```

## 5. Run the demo

This runs the sample end-to-end flow on a copy of the tracker:

```powershell
python scripts\run_demo.py
```

## 6. Run the email server

Single test pass:

```powershell
python src\server\email_server.py --once
```

Run continuously:

```powershell
python src\server\email_server.py
```

Run continuously with a custom polling interval:

```powershell
python src\server\email_server.py --interval 30
```

## 7. Run the individual scripts

Parse an invoice summary email body:

```powershell
python src\parsing\parse_invoice_summary.py path\to\email_body.txt
```

Append parsed invoice data into the tracker:

```powershell
python src\excel\append_invoice_to_excel.py path\to\tracker.xlsx path\to\parsed_data.json --entity usa --requested-by "Sender Name"
```

Generate a draft email from an Excel row:

```powershell
python src\mailer\draft_email_from_excel_row.py path\to\tracker.xlsx --entity usa --invoice-no "INT/USA/26-27/001"
```

Generate an invoice PDF:

```powershell
python src\pdf\generate_invoice_pdf.py path\to\row.json path\to\output.pdf
```

## 8. Useful files

- `README.md` for the full project overview
- `config/email_server_config.json` for mailbox, polling, and file paths
- `.env` for Gmail credentials
- `output/` for logs, generated PDFs, draft previews, and processed state

## 9. Quick start

If you just want the shortest path:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
Copy-Item .env.example .env
python src\server\email_server.py --once
```
