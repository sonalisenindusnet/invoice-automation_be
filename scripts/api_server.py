"""
api_server.py

FastAPI endpoint for creating an invoice row directly from the frontend.

Endpoint:
    POST /invoice/api/v1/invoice-generation

The request is converted into the internal format append_invoice() expects,
then the existing invoice pipeline does the rest: pick the entity/tab, read
the configured tracker (Google Sheets or local Excel), generate the next
invoice number, calculate tax and due date, run MIS verification, append the
row, and return the result.

Which tracker is used is controlled by "tracker_source" in
config/email_server_config.json ("google_sheets" or "excel").
"""

import json
import logging
import sys
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Project paths -- src/ on sys.path so imports below use the same bare
# "package.module" convention as the rest of the codebase.
# ---------------------------------------------------------------------------

BASE = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE / "config" / "email_server_config.json"

SRC = BASE / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from utils.env_loader import load_env_file
from excel.append_invoice_to_excel import append_invoice

load_env_file()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("invoice_api")

app = FastAPI(
    title="Invoice API",
    description="Invoice generation API",
    version="1.0.0",
)

# The frontend calls this API cross-origin, so the browser sends a CORS
# preflight OPTIONS request before every POST. allow_origins is "*" since
# this API uses no cookies/credentials; tighten this if it's ever exposed
# beyond your own frontend.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


class InvoiceGenerationRequest(BaseModel):
    pfId: str = Field(..., min_length=1)
    accountName: str = ""
    clientName: str = Field(..., min_length=1)
    invoiceDescription: str = ""
    contactPersonName: str = ""
    clientMailTo: str = ""
    clientMailCc: str = ""
    intCcMailId: str = ""
    workOrder: str = ""
    masterProjectId: str = ""
    currency: str = "USD"
    projectValue: str = ""
    invoiceValue: str = Field(..., min_length=1)
    invoiceType: str = ""
    entity: str = Field(..., min_length=1)


def load_config():
    """Load email_server_config.json (shared with the email server)."""
    if not CONFIG_PATH.exists():
        raise RuntimeError(f"Configuration file not found: {CONFIG_PATH}")
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Tracker configuration
# ---------------------------------------------------------------------------

def _google_sheets_tracker_ref(cfg):
    sheet_id = cfg.get("google_sheet_id")
    service_account_json = cfg.get("google_service_account_json")

    if not sheet_id:
        raise RuntimeError("google_sheet_id is missing from email_server_config.json")
    if not service_account_json:
        raise RuntimeError("google_service_account_json is missing from email_server_config.json")

    credentials_path = Path(service_account_json)
    if not credentials_path.is_absolute():
        # email_server_config.json lives in config/, so relative paths
        # resolve from there.
        credentials_path = CONFIG_PATH.parent / credentials_path
    credentials_path = credentials_path.resolve()

    if not credentials_path.exists():
        raise RuntimeError(f"Google service account file not found: {credentials_path}")
    if not credentials_path.is_file():
        raise RuntimeError(f"Google service account path is not a file: {credentials_path}")

    logger.info("Tracker source: Google Sheets")
    logger.info("Google Sheet ID: %s", sheet_id)
    logger.info("Service account: %s", credentials_path)

    return {
        "type": "google_sheets",
        "sheet_id": sheet_id,
        "service_account_json": str(credentials_path),
    }


def _excel_tracker_ref(cfg):
    tracker_xlsx_path = cfg.get("tracker_xlsx_path")
    if not tracker_xlsx_path:
        raise RuntimeError("tracker_xlsx_path is missing from email_server_config.json")

    tracker_path = Path(tracker_xlsx_path)
    if not tracker_path.is_absolute():
        tracker_path = CONFIG_PATH.parent / tracker_path
    tracker_path = tracker_path.resolve()

    logger.info("Tracker source: Excel")
    logger.info("Excel tracker: %s", tracker_path)

    return str(tracker_path)


def get_tracker_reference(cfg):
    """Return the tracker reference expected by tracker_io.py: a dict for
    Google Sheets, or a plain path string for local Excel."""
    tracker_source = cfg.get("tracker_source", "excel")
    if not isinstance(tracker_source, str):
        raise RuntimeError("tracker_source must be a string")
    tracker_source = tracker_source.strip().lower()

    if tracker_source == "google_sheets":
        return _google_sheets_tracker_ref(cfg)
    if tracker_source == "excel":
        return _excel_tracker_ref(cfg)

    raise RuntimeError(
        f"Unsupported tracker_source: {tracker_source}. Expected 'google_sheets' or 'excel'."
    )


# ---------------------------------------------------------------------------
# API -> internal invoice format
# ---------------------------------------------------------------------------

def convert_request_to_invoice_data(request: InvoiceGenerationRequest):
    """Convert the frontend API format into the internal format expected by
    append_invoice()."""
    try:
        invoice_amount = float(request.invoiceValue)
    except (TypeError, ValueError):
        raise HTTPException(status_code=400, detail="invoiceValue must be a valid number")

    if invoice_amount < 0:
        raise HTTPException(status_code=400, detail="invoiceValue cannot be negative")

    client_mail_to = [request.clientMailTo.strip()] if request.clientMailTo.strip() else []
    int_cc_mail = [request.intCcMailId.strip()] if request.intCcMailId.strip() else []

    return {
        "pf_id": request.pfId.strip(),
        "client_company": request.clientName.strip(),
        "client_contact_person": request.contactPersonName.strip(),
        "invoice_description": request.invoiceDescription.strip(),
        "master_project_id": request.masterProjectId.strip(),
        "work_order": request.workOrder.strip(),
        "currency": request.currency.strip().upper(),
        "invoice_value": {"amount": invoice_amount},
        "client_mail_to": client_mail_to,
        "int_cc_mail": int_cc_mail,
    }


SUPPORTED_ENTITIES = {"usa", "uk", "poland"}


def normalize_entity(entity: str) -> str:
    """Validate and normalize the requested entity."""
    entity_key = entity.strip().lower()
    if entity_key not in SUPPORTED_ENTITIES:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported entity '{entity}'. "
                f"Supported entities: {', '.join(sorted(SUPPORTED_ENTITIES))}"
            ),
        )
    return entity_key


@app.get("/health")
def health_check():
    return {"success": True, "status": "UP"}


# ---------------------------------------------------------------------------
# Invoice generation API
# ---------------------------------------------------------------------------

def _failure_response(result):
    logger.warning("Invoice was not appended: %s", result)
    return {
        "success": False,
        "status": result.get("status"),
        "message": result.get("message", "Invoice was not appended"),
    }


def _success_response(result, row):
    invoice_no = row.get("invoice_no")
    logger.info("Invoice successfully created: %s", invoice_no)
    return {
        "success": True,
        "status": "saved",
        "message": "Invoice saved successfully",
        "invoice_no": invoice_no,
        "entity": result.get("entity_key"),
        "sheet": result.get("sheet"),
        "pf_id": row.get("pf_id"),
        "client_name": row.get("client_company"),
        "invoice_value": row.get("base_amount"),
        "total": row.get("total"),
        "currency": row.get("currency"),
        "due_date": row.get("due_date"),
        "review_status": row.get("review_status"),
        "email_drafted": row.get("email_drafted"),
        "mis_verification": result.get("mis_verification"),
    }


@app.post("/invoice/api/v1/invoice-generation")
def invoice_generation(request: InvoiceGenerationRequest):
    """Create a new invoice row in the configured tracker (tracker_source =
    google_sheets or excel, per config/email_server_config.json)."""
    logger.info(
        "Invoice generation request received: PF ID=%s, entity=%s",
        request.pfId, request.entity,
    )

    entity_key = normalize_entity(request.entity)
    data = convert_request_to_invoice_data(request)

    try:
        cfg = load_config()
        tracker_ref = get_tracker_reference(cfg)

        logger.info("TRACKER SOURCE: %s", cfg.get("tracker_source"))
        logger.info("TRACKER REFERENCE: %s", tracker_ref)
        logger.info("ENTITY: %s", entity_key)

        # append_invoice() already knows how to read either tracker
        # reference via utils.tracker_io.load_tracker_with_retry().
        result = append_invoice(
            xlsx_path=tracker_ref,
            data=data,
            requested_by=request.clientMailTo,
            entity_key=entity_key,
        )

        if result.get("status") != "appended":
            return _failure_response(result)

        return _success_response(result, result.get("row", {}))

    except HTTPException:
        raise

    except Exception as exc:
        logger.exception("Invoice generation failed for PF ID=%s", request.pfId)
        raise HTTPException(status_code=500, detail=f"Invoice generation failed: {str(exc)}")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("scripts.api_server:app", host="0.0.0.0", port=5000, reload=True)
