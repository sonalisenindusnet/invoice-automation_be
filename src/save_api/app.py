"""
app.py

The save API: one endpoint that takes an invoice request from the frontend
and appends it as a new row into the correct tab of the tracker (a live
Google Sheet -- see utils/tracker_io.py). Nothing else -- no tax
calculation, no PDF, no email drafting.

    POST /invoice/api/v1/invoice-generation
    GET  /health
"""
import json
import logging
import sys
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

SRC_DIR = Path(__file__).resolve().parent.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from save_api.excel_writer import (
    save_invoice, update_mis_verified, UnknownEntityError, RowNotFoundError, PfIdMismatchError,
)
from utils.tracker_io import tracker_ref_from_config

PROJECT_ROOT = SRC_DIR.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "save_api_config.json"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("save_api")

app = FastAPI(title="Save Invoice API", version="1.0.0")

# The frontend calls this cross-origin, so the browser preflights every
# POST with an OPTIONS request -- without this middleware that preflight
# has nowhere to land and the browser never sends the real request.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

SUPPORTED_ENTITIES = {"usa", "uk", "poland"}


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


class MisVerificationUpdateRequest(BaseModel):
    invoiceNo: str = Field(..., min_length=1)
    entity: str = Field(..., min_length=1)
    pfId: str = ""
    misUpdateFlag: bool


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def tracker_path_from_config(cfg):
    """Despite the name (kept for call-site stability), this returns a
    `tracker_ref` dict pointing at the live Google Sheet -- see
    utils.tracker_io.tracker_ref_from_config for the resolution logic."""
    return tracker_ref_from_config(cfg, CONFIG_PATH.parent)


def normalize_entity(entity: str) -> str:
    entity_key = entity.strip().lower()
    if entity_key not in SUPPORTED_ENTITIES:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported entity '{entity}'. Supported entities: {', '.join(sorted(SUPPORTED_ENTITIES))}",
        )
    return entity_key


def convert_request(request: InvoiceGenerationRequest):
    """Converts the frontend's request shape into what save_invoice() expects."""
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


@app.get("/health")
def health_check():
    return {"success": True, "status": "UP"}


@app.post("/invoice/api/v1/invoice-generation")
def invoice_generation(request: InvoiceGenerationRequest):
    logger.info("Invoice save request received: PF ID=%s, entity=%s", request.pfId, request.entity)

    entity_key = normalize_entity(request.entity)
    data = convert_request(request)

    try:
        cfg = load_config()
        tracker_ref = tracker_path_from_config(cfg)
        result = save_invoice(
            data, entity_key, tracker_ref, requested_by=request.clientMailTo,
        )
    except UnknownEntityError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Invoice save failed for PF ID=%s", request.pfId)
        raise HTTPException(status_code=500, detail=f"Invoice save failed: {exc}")

    row = result["row"]
    logger.info("Invoice saved: %s into '%s' for PF ID=%s", row.get("invoice_no"), result["sheet"], row.get("pf_id"))

    return {
        "success": True,
        "status": "saved",
        "message": "Invoice saved successfully",
        "invoice_no": row.get("invoice_no"),
        "entity": entity_key,
        "sheet": result["sheet"],
        "pf_id": row.get("pf_id"),
        "client_name": row.get("client_company"),
        "invoice_value": row.get("total"),
        "currency": row.get("currency"),
        "created_at": row.get("created_at"),
    }


@app.post("/invoice/api/v1/mis-verification")
def mis_verification_update(request: MisVerificationUpdateRequest):
    logger.info(
        "MIS verification update received: invoiceNo=%s, entity=%s, pfId=%s, flag=%s",
        request.invoiceNo, request.entity, request.pfId, request.misUpdateFlag,
    )

    entity_key = normalize_entity(request.entity)

    try:
        cfg = load_config()
        tracker_ref = tracker_path_from_config(cfg)
        result = update_mis_verified(
            entity_key, request.invoiceNo.strip(), request.pfId.strip(), request.misUpdateFlag, tracker_ref,
        )
    except UnknownEntityError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except RowNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except PfIdMismatchError as exc:
        raise HTTPException(status_code=409, detail=str(exc))
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("MIS verification update failed for invoiceNo=%s", request.invoiceNo)
        raise HTTPException(status_code=500, detail=f"MIS verification update failed: {exc}")

    logger.info(
        "MIS verification updated: %s in '%s' set to %s",
        result["invoice_no"], result["sheet"], result["mis_verification_done"],
    )

    return {
        "success": True,
        "status": "updated",
        "entity": entity_key,
        "sheet": result["sheet"],
        "invoice_no": result["invoice_no"],
        "mis_verification_done": result["mis_verification_done"],
    }
