# Run Commands

From the project root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Start the API:

```powershell
python scripts\api_server.py
```

Test the Google Sheets draft server once:

```powershell
python src\server\email_server.py --once
```

Run the draft server continuously:

```powershell
python src\server\email_server.py
```

Use `Ctrl+C` to stop the continuous server.
