"""
routes.py

The save API: HTTP layer only -- the FastAPI app, its endpoints, and the
thin request/response plumbing around them. Request/response shapes live in
models/invoice_models.py; all business logic (Excel/tracker writes, tax,
PDF, email) lives in services/.

    POST /invoice/api/v1/invoice-generation
    POST /invoice/api/v1/dedicated-invoice-generation
    POST /invoice/api/v1/mis-verification
    GET  /health

Moved here from save_api/app.py on 2026-10-06 as part of the
api/models/services/utils restructure -- see invoice-automation-architecture.md's
2026-10-06 update #4 for the full story. Content/behavior is otherwise
unchanged from save_api/app.py.
"""
import json
import logging
import sys
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

SRC_DIR = Path(__file__).resolve().parent.parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from models.invoice_models import (
    InvoiceGenerationRequest, ResourceItem, DedicatedInvoiceGenerationRequest,
    MisVerificationUpdateRequest,
)
from services.excel_writer import (
    save_invoice, update_mis_verified, UnknownEntityError, RowNotFoundError, PfIdMismatchError,
    SchemaDriftError,
)
from services.dedicated_invoice import consolidate_resources_by_email
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

SUPPORTED_ENTITIES = {"usa", "uk", "poland", "singapore"}


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
        "client_country": request.company_location.strip(),
        "currency": request.currency.strip().upper(),
        "invoice_value": {"amount": invoice_amount},
        "client_mail_to": client_mail_to,
        "int_cc_mail": int_cc_mail,
        "company_address": request.companyAddress.strip(),
        "business_model": request.clientType.strip(),
    }


@app.get("/health")
def health_check():
    return {"success": True, "status": "UP"}


@app.post("/invoice/api/v1/invoice-generation")
def invoice_generation(request: InvoiceGenerationRequest):
    logger.info(
        "Invoice save request received: PF ID=%s, entity=%s, company_location=%r",
        request.pfId, request.entity, request.company_location,
    )
    if not request.company_location.strip():
        # Not an error -- a blank/missing value is deliberately treated as
        # "foreign" (see services.tax_calculator), so this never blocks a
        # save. But it's worth a visible WARNING (not just silence) since a
        # wrong field name/transport on the frontend's side would look
        # exactly like this -- every invoice quietly getting the foreign
        # tax rate with no error anywhere. If you're expecting a value here
        # and see this warning instead, check what the frontend is
        # actually sending.
        logger.warning(
            "Invoice save request for PF ID=%s has no company_location -- "
            "will be treated as a foreign client (0%% local tax rate) for tax purposes",
            request.pfId,
        )

    entity_key = normalize_entity(request.entity)
    data = convert_request(request)

    try:
        cfg = load_config()
        tracker_ref = tracker_path_from_config(cfg)
        result = save_invoice(
            data, entity_key, tracker_ref, requested_by=request.raisedByEmail,
        )
    except UnknownEntityError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except SchemaDriftError as exc:
        logger.error("Schema drift detected for PF ID=%s: %s", request.pfId, exc)
        raise HTTPException(status_code=500, detail=str(exc))
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Invoice save failed for PF ID=%s", request.pfId)
        raise HTTPException(status_code=500, detail=f"Invoice save failed: {exc}")

    row = result["row"]
    tax = result["tax"]
    if request.company_location.strip() and not tax["country_recognized"]:
        # The frontend's client-location field is a fixed dropdown as of
        # 2026-08-24 (only ever "United States"/"United Kingdom"/
        # "Singapore"/"Poland") -- so a non-blank value that matches NONE
        # of them is now a real bug signal (a dropdown change, an encoding
        # issue, a new value nobody told this service about), not normal
        # variation. The save still proceeds with the foreign tax rate
        # either way -- this is visibility, not a block.
        logger.warning(
            "Invoice save request for PF ID=%s has company_location=%r, which "
            "doesn't match any known country -- expected one of the fixed "
            "dropdown values. Tax was computed as if this were a foreign client.",
            request.pfId, request.company_location,
        )
    logger.info(
        "Invoice saved: %s into '%s' for PF ID=%s (%s %.2f%% -> %s tax on %s subtotal)",
        row.get("invoice_no"), result["sheet"], row.get("pf_id"),
        tax["tax_name"], tax["rate"] * 100, tax["tax_amount"], tax["subtotal"],
    )

    return {
        "success": True,
        "status": "saved",
        "message": "Invoice saved successfully",
        "invoice_no": row.get("invoice_no"),
        "entity": entity_key,
        "sheet": result["sheet"],
        "pf_id": row.get("pf_id"),
        "client_name": row.get("client_company"),
        "company_location": row.get("country"),
        "invoice_value": row.get("total"),
        "currency": row.get("currency"),
        "created_at": row.get("created_at"),
        # Tax breakdown -- computed from (entity, client_country, the
        # REQUEST's raw pre-tax invoiceValue) via
        # services.tax_calculator.compute_tax(). "invoice_value" above is
        # the row's own "Total Amount" column, which (2026-08-24) is the
        # POST-tax grand total -- i.e. it equals "total_with_tax" below,
        # NOT "subtotal". "subtotal" here is what the frontend originally
        # sent as invoiceValue (before tax was added); "total" is
        # subtotal + tax_amount, which is what actually got written to
        # the tracker's "Total Amount" (and, where the tab has one, its
        # VAT/GST column got "tax_amount").
        "tax_name": tax["tax_name"],
        "tax_rate": tax["rate"],
        "tax_amount": tax["tax_amount"],
        "subtotal": tax["subtotal"],
        "total_with_tax": tax["total"],
    }


@app.post("/invoice/api/v1/dedicated-invoice-generation")
def dedicated_invoice_generation(request: DedicatedInvoiceGenerationRequest):
    logger.info(
        "Dedicated invoice save request received: PF ID=%s, entity=%s, resources=%d",
        request.pfId, request.entity, len(request.resources),
    )

    entity_key = normalize_entity(request.entity)

    try:
        invoice_amount = float(request.invoiceValue)
    except (TypeError, ValueError):
        raise HTTPException(status_code=400, detail="invoiceValue must be a valid number")
    if invoice_amount < 0:
        raise HTTPException(status_code=400, detail="invoiceValue cannot be negative")

    # Consolidate resources by email
    consolidated = consolidate_resources_by_email(
        [r.model_dump() for r in request.resources]
    )

    # Build description for Excel from consolidated resources
    excel_description = consolidated["description"]

    # Prepare base invoice data
    client_mail_to = [request.clientMailTo.strip()] if request.clientMailTo.strip() else []
    int_cc_mail = [request.intCcMailId.strip()] if request.intCcMailId.strip() else []

    data = {
        "pf_id": request.pfId.strip(),
        "client_company": request.clientName.strip(),
        "client_contact_person": request.contactPersonName.strip(),
        "invoice_description": request.invoiceDescription.strip(),  # Original description for email and Excel
        "client_country": request.companyLocation.strip(),
        "currency": request.currency.strip().upper(),
        "invoice_value": {"amount": invoice_amount},
        "client_mail_to": client_mail_to,
        "int_cc_mail": int_cc_mail,
        "company_address": request.companyAddress.strip(),
        # Store consolidated resources in resource_description column for PDF generation
        "resource_description": excel_description,
        "business_model": request.clientType.strip(),
    }

    try:
        cfg = load_config()
        tracker_ref = tracker_path_from_config(cfg)
        result = save_invoice(
            data, entity_key, tracker_ref, requested_by=request.raisedByEmail,
        )
    except UnknownEntityError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except SchemaDriftError as exc:
        logger.error("Schema drift detected for PF ID=%s: %s", request.pfId, exc)
        raise HTTPException(status_code=500, detail=str(exc))
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Dedicated invoice save failed for PF ID=%s", request.pfId)
        raise HTTPException(status_code=500, detail=f"Dedicated invoice save failed: {exc}")

    row = result["row"]
    tax = result["tax"]

    logger.info(
        "Dedicated invoice saved: %s into '%s' for PF ID=%s with %d consolidated resources",
        row.get("invoice_no"), result["sheet"], row.get("pf_id"),
        len(consolidated["consolidated_resources"]),
    )

    return {
        "success": True,
        "status": "saved",
        "message": "Dedicated invoice saved successfully",
        "invoice_no": row.get("invoice_no"),
        "entity": entity_key,
        "sheet": result["sheet"],
        "pf_id": row.get("pf_id"),
        "client_name": row.get("client_company"),
        "company_location": row.get("country"),
        "invoice_value": row.get("total"),
        "currency": row.get("currency"),
        "created_at": row.get("created_at"),
        "resources_count": len(consolidated["consolidated_resources"]),
        "resources": consolidated["consolidated_resources"],
        "excel_description": consolidated["description"],
        "tax_name": tax["tax_name"],
        "tax_rate": tax["rate"],
        "tax_amount": tax["tax_amount"],
        "subtotal": tax["subtotal"],
        "total_with_tax": tax["total"],
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


# Per-entity currency used for dummy data -- independent of each tab's
# "default_currency" in config/tabs/*.json (that's just what gets written
# when a request omits currency-parsing logic entirely; a real request can
# and does send other currencies, as seen from the frontend).
DUMMY_ENTITY_CURRENCY = {
    "usa": "USD",
    "uk": "GBP",
    "poland": "EUR",
    "singapore": "SGD",
}


def _build_dummy_dedicated_request(entity_key):
    """Builds one DedicatedInvoiceGenerationRequest-shaped dummy payload for
    `entity_key`, with a fresh unique pfId so repeated calls never collide
    with real data or with each other."""
    run_id = uuid4().hex[:8]
    dummy_pf_id = f"DUMMY/{entity_key.upper()}/{run_id}"
    return DedicatedInvoiceGenerationRequest(
        pfId=dummy_pf_id,
        clientName=f"Dummy Client {entity_key.upper()}",
        contactPersonName="Dummy Contact",
        clientMailTo="dummy.client@example.com",
        clientMailCc="dummy.cc@example.com",
        intCcMailId="dummy.int.cc@example.com",
        clientType="ECB",
        raisedByEmail="dummy.raisedby@example.com",
        companyLocation="India",
        companyAddress="123 Dummy Street",
        companyGeography="India",
        entity=entity_key.upper(),
        currency=DUMMY_ENTITY_CURRENCY[entity_key],
        invoiceValue="1500",
        invoiceDescription=f"Dummy invoice for {entity_key.upper()} column-layout check",
        resources=[
            # Distinct emails -- resouceEmail is mandatory as of 2026-10-06
            # (see models.invoice_models.ResourceItem), and these two are
            # deliberately different people so this dummy call still
            # demonstrates two separate line items, not the same-email
            # merge (see services.dedicated_invoice.consolidate_resources_by_email()).
            ResourceItem(
                pfId=f"{dummy_pf_id}/1", resourceName="Dummy Resource One",
                invoiceAmount="1000", resouceEmail="dummy.resource1@example.com",
            ),
            ResourceItem(
                pfId=f"{dummy_pf_id}/2", resourceName="Dummy Resource Two",
                invoiceAmount="500", resouceEmail="dummy.resource2@example.com",
            ),
        ],
    )


@app.post("/dummy-all-country-data")
def dummy_all_country_data():
    """Dev-only helper: saves one dummy invoice per supported entity by
    calling dedicated_invoice_generation() directly (in-process -- no HTTP
    call to this same server), so the tracker's per-entity column layout
    can be eyeballed manually in each sheet without hand-crafting a request
    per country."""
    results = []
    for entity_key in sorted(SUPPORTED_ENTITIES):
        try:
            request = _build_dummy_dedicated_request(entity_key)
            response = dedicated_invoice_generation(request)
            results.append({
                "entity": entity_key,
                "success": True,
                "invoice_no": response["invoice_no"],
                "sheet": response["sheet"],
                "pf_id": response["pf_id"],
            })
        except HTTPException as exc:
            results.append({"entity": entity_key, "success": False, "error": exc.detail})
        except Exception as exc:
            logger.exception("Dummy data generation failed for entity=%s", entity_key)
            results.append({"entity": entity_key, "success": False, "error": str(exc)})

    return {
        "success": all(r["success"] for r in results),
        "results": results,
    }
