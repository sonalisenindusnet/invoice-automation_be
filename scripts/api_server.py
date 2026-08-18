
"""
api_server.py

FastAPI endpoint for creating an invoice row directly from the frontend.

Endpoint:
    POST /invoice/api/v1/invoice-generation

The API converts the frontend request into the same internal data format
used by append_invoice().

It then uses the existing invoice pipeline to:

    1. Select the requested entity/tab
    2. Read the configured Google Sheet / Excel tracker
    3. Generate the next invoice number
    4. Calculate tax
    5. Calculate payment due date
    6. Call MIS verification
    7. Append the row
    8. Return the generated invoice number

The tracker destination is controlled by:

    config/email_server_config.json

For Google Sheets:

    "tracker_source": "google_sheets"

For local Excel:

    "tracker_source": "excel"

The existing append_invoice() implementation is reused for both.
"""

import json
import logging
import sys
from pathlib import Path

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Project paths
# ---------------------------------------------------------------------------

BASE = Path(__file__).resolve().parent.parent

CONFIG_PATH = BASE / "config" / "email_server_config.json"


# ---------------------------------------------------------------------------
# Make sure project root is importable
# ---------------------------------------------------------------------------

if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))


# ---------------------------------------------------------------------------
# Existing project imports
# ---------------------------------------------------------------------------

from scripts.env_loader import load_env_file

from src.excel.append_invoice_to_excel import append_invoice


# ---------------------------------------------------------------------------
# Load environment
# ---------------------------------------------------------------------------

load_env_file()


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

logger = logging.getLogger("invoice_api")


# ---------------------------------------------------------------------------
# FastAPI
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Invoice API",
    description="Invoice generation API",
    version="1.0.0",
)


# ---------------------------------------------------------------------------
# Request model
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

def load_config():
    """
    Load email_server_config.json.

    The same configuration file is used by the email server and this API.
    """

    if not CONFIG_PATH.exists():
        raise RuntimeError(
            f"Configuration file not found: {CONFIG_PATH}"
        )

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Tracker configuration
# ---------------------------------------------------------------------------

def get_tracker_reference(cfg):
    """
    Return the tracker reference expected by tracker_io.py.

    Google Sheets:
        {
            "type": "google_sheets",
            "sheet_id": "...",
            "service_account_json": "..."
        }

    Excel:
        "/absolute/path/to/Invoice Traker.xlsx"
    """

    tracker_source = cfg.get(
        "tracker_source",
        "excel",
    )

    if not isinstance(tracker_source, str):
        raise RuntimeError(
            "tracker_source must be a string"
        )

    tracker_source = tracker_source.strip().lower()

    # ------------------------------------------------------------------
    # Google Sheets
    # ------------------------------------------------------------------

    if tracker_source == "google_sheets":

        sheet_id = cfg.get("google_sheet_id")

        service_account_json = cfg.get(
            "google_service_account_json"
        )

        if not sheet_id:
            raise RuntimeError(
                "google_sheet_id is missing from "
                "email_server_config.json"
            )

        if not service_account_json:
            raise RuntimeError(
                "google_service_account_json is missing from "
                "email_server_config.json"
            )

        credentials_path = Path(
            service_account_json
        )

        # email_server_config.json is inside config/,
        # so resolve relative credential paths from config/.
        if not credentials_path.is_absolute():
            credentials_path = (
                CONFIG_PATH.parent / credentials_path
            )

        credentials_path = credentials_path.resolve()

        if not credentials_path.exists():
            raise RuntimeError(
                "Google service account file not found: "
                f"{credentials_path}"
            )

        if not credentials_path.is_file():
            raise RuntimeError(
                "Google service account path is not a file: "
                f"{credentials_path}"
            )

        tracker_ref = {
            "type": "google_sheets",
            "sheet_id": sheet_id,
            "service_account_json": str(credentials_path),
        }

        logger.info(
            "Tracker source: Google Sheets"
        )

        logger.info(
            "Google Sheet ID: %s",
            sheet_id,
        )

        logger.info(
            "Service account: %s",
            credentials_path,
        )

        return tracker_ref

    # ------------------------------------------------------------------
    # Excel fallback
    # ------------------------------------------------------------------

    if tracker_source == "excel":

        tracker_xlsx_path = cfg.get(
            "tracker_xlsx_path"
        )

        if not tracker_xlsx_path:
            raise RuntimeError(
                "tracker_xlsx_path is missing from "
                "email_server_config.json"
            )

        tracker_path = Path(
            tracker_xlsx_path
        )

        if not tracker_path.is_absolute():
            tracker_path = (
                CONFIG_PATH.parent / tracker_path
            )

        tracker_path = tracker_path.resolve()

        logger.info(
            "Tracker source: Excel"
        )

        logger.info(
            "Excel tracker: %s",
            tracker_path,
        )

        return str(tracker_path)

    raise RuntimeError(
        f"Unsupported tracker_source: {tracker_source}. "
        "Expected 'google_sheets' or 'excel'."
    )

# ---------------------------------------------------------------------------
# API → internal invoice format
# ---------------------------------------------------------------------------

def convert_request_to_invoice_data(
    request: InvoiceGenerationRequest,
):
    """
    Convert the frontend API format into the internal format expected
    by append_invoice().
    """

    try:
        invoice_amount = float(
            request.invoiceValue
        )

    except (TypeError, ValueError):
        raise HTTPException(
            status_code=400,
            detail="invoiceValue must be a valid number",
        )

    if invoice_amount < 0:
        raise HTTPException(
            status_code=400,
            detail="invoiceValue cannot be negative",
        )

    client_mail_to = []

    if request.clientMailTo.strip():
        client_mail_to.append(
            request.clientMailTo.strip()
        )

    int_cc_mail = []

    if request.intCcMailId.strip():
        int_cc_mail.append(
            request.intCcMailId.strip()
        )

    return {
        "pf_id": request.pfId.strip(),

        "client_company": request.clientName.strip(),

        "client_contact_person": (
            request.contactPersonName.strip()
        ),

        "invoice_description": (
            request.invoiceDescription.strip()
        ),

        "master_project_id": (
            request.masterProjectId.strip()
        ),

        "work_order": request.workOrder.strip(),

        "currency": request.currency.strip().upper(),

        "invoice_value": {
            "amount": invoice_amount
        },

        "client_mail_to": client_mail_to,

        "int_cc_mail": int_cc_mail,
    }


# ---------------------------------------------------------------------------
# Entity validation
# ---------------------------------------------------------------------------

SUPPORTED_ENTITIES = {
    "usa",
    "uk",
    "poland",
}


def normalize_entity(entity: str) -> str:
    """
    Validate and normalize the requested entity.
    """

    entity_key = entity.strip().lower()

    if entity_key not in SUPPORTED_ENTITIES:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Unsupported entity '{entity}'. "
                f"Supported entities: "
                f"{', '.join(sorted(SUPPORTED_ENTITIES))}"
            ),
        )

    return entity_key


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------

@app.get("/health")
def health_check():
    return {
        "success": True,
        "status": "UP",
    }


# ---------------------------------------------------------------------------
# Invoice generation API
# ---------------------------------------------------------------------------

@app.post("/invoice/api/v1/invoice-generation")
def invoice_generation(
    request: InvoiceGenerationRequest,
):
    """
    Create a new invoice row in the configured tracker.

    The destination is controlled by:

        tracker_source = google_sheets
            OR
        tracker_source = excel
    """

    logger.info(
        "Invoice generation request received: "
        "PF ID=%s, entity=%s",
        request.pfId,
        request.entity,
    )

    # -----------------------------------------------------------------------
    # Validate entity
    # -----------------------------------------------------------------------

    entity_key = normalize_entity(
        request.entity
    )

    # -----------------------------------------------------------------------
    # Convert frontend payload
    # -----------------------------------------------------------------------

    data = convert_request_to_invoice_data(
        request
    )

    try:

        # -------------------------------------------------------------------
        # Load configuration
        # -------------------------------------------------------------------

        cfg = load_config()

        # -------------------------------------------------------------------
        # Build tracker reference
        #
        # This is the important part:
        #
        # tracker_source = google_sheets
        #     -> Google Sheets reference
        #
        # tracker_source = excel
        #     -> local Excel path
        # -------------------------------------------------------------------

        tracker_ref = get_tracker_reference(
            cfg
        )

        logger.info(
            "TRACKER SOURCE: %s",
            cfg.get("tracker_source"),
        )

        logger.info(
            "TRACKER REFERENCE: %s",
            tracker_ref,
        )

        logger.info(
            "ENTITY: %s",
            entity_key,
        )

        # -------------------------------------------------------------------
        # Reuse the existing invoice pipeline
        #
        # append_invoice() already uses:
        #
        #     utils.tracker_io.load_tracker_with_retry()
        #
        # which is designed to handle the configured tracker reference.
        # -------------------------------------------------------------------

        result = append_invoice(
            xlsx_path=tracker_ref,
            data=data,
            requested_by=request.clientMailTo,
            entity_key=entity_key,
        )

        # -------------------------------------------------------------------
        # append_invoice() did not append
        # -------------------------------------------------------------------

        if result.get("status") != "appended":

            logger.warning(
                "Invoice was not appended: %s",
                result,
            )

            return {
                "success": False,
                "status": result.get("status"),
                "message": result.get(
                    "message",
                    "Invoice was not appended",
                ),
            }

        # -------------------------------------------------------------------
        # Successful append
        # -------------------------------------------------------------------

        row = result.get(
            "row",
            {},
        )

        invoice_no = row.get(
            "invoice_no"
        )

        logger.info(
            "Invoice successfully created: %s",
            invoice_no,
        )

        return {
            "success": True,

            "status": "saved",

            "message": "Invoice saved successfully",

            "invoice_no": invoice_no,

            "entity": result.get(
                "entity_key"
            ),

            "sheet": result.get(
                "sheet"
            ),

            "pf_id": row.get(
                "pf_id"
            ),

            "client_name": row.get(
                "client_company"
            ),

            "invoice_value": row.get(
                "base_amount"
            ),

            "total": row.get(
                "total"
            ),

            "currency": row.get(
                "currency"
            ),

            "due_date": row.get(
                "due_date"
            ),

            "review_status": row.get(
                "review_status"
            ),

            "email_drafted": row.get(
                "email_drafted"
            ),

            "mis_verification": result.get(
                "mis_verification"
            ),
        }

    except HTTPException:
        raise

    except Exception as exc:

        logger.exception(
            "Invoice generation failed for PF ID=%s",
            request.pfId,
        )

        raise HTTPException(
            status_code=500,
            detail=(
                f"Invoice generation failed: {str(exc)}"
            ),
        )


# ---------------------------------------------------------------------------
# Run directly
# ---------------------------------------------------------------------------

if __name__ == "__main__":

    import uvicorn

    uvicorn.run(
        "scripts.api_server:app",
        host="0.0.0.0",
        port=5000,
        reload=True,
    )
