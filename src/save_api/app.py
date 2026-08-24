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
from pydantic import AliasChoices, BaseModel, Field, model_validator

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

SUPPORTED_ENTITIES = {"usa", "uk", "poland", "singapore"}

# Frontend forms commonly send an explicit JSON `null` for an empty optional
# field instead of omitting the key. Plain `str` fields reject `null`
# outright with a 422 validation error (confirmed via a direct test:
# sending raisedByEmail=null on its own crashed the whole request, not just
# that field) -- so every genuinely-optional, blank-string-default field on
# InvoiceGenerationRequest is normalized from `null` to `""` before
# Pydantic's own type validation runs (see that model's `_null_optional_
# strings_to_empty` validator). Deliberately excludes the required fields
# (pfId, clientName, invoiceValue, entity) and `currency` (whose real
# default is "USD", not "") -- a `null` for any of those should keep
# failing loudly rather than being silently papered over. Module-level (not
# a class attribute) because Pydantic v2 treats an underscore-prefixed
# CLASS attribute as a private model field slot, not a plain constant --
# confirmed the hard way (`TypeError: 'ModelPrivateAttr' object is not
# iterable`) before moving it here.
NULLABLE_TO_EMPTY_FIELDS = (
    "accountName", "invoiceDescription", "contactPersonName", "clientMailTo",
    "clientMailCc", "intCcMailId", "workOrder", "masterProjectId",
    "company_location", "companyLocation", "projectValue", "invoiceType",
    "raisedByEmail",
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
    # Originally told the frontend sends this as "company_location"
    # (snake_case, unlike every other field on this request). A live test
    # on 2026-08-24 showed a request actually using "companyLocation"
    # (camelCase, matching this model's other fields) instead -- which
    # Pydantic silently ignored, since an unrecognized key just falls back
    # to the default "" rather than erroring. Accepting BOTH spellings here
    # means whichever one the frontend actually sends works, instead of
    # this field quietly going blank (and every invoice being taxed as
    # "foreign") again if it changes back.
    company_location: str = Field(
        default="",
        validation_alias=AliasChoices("company_location", "companyLocation"),
    )
    # Required, no default -- per explicit instruction, a save request that
    # doesn't specify a currency should fail loudly (422), not be silently
    # assumed to be USD. This matters most for Singapore (multi-currency:
    # SGD, USD, EUR, ... depending on the client), but applies to every
    # entity now -- the frontend must always say which currency an invoice
    # is actually in.
    currency: str = Field(..., min_length=1)
    projectValue: str = ""
    invoiceValue: str = Field(..., min_length=1)
    invoiceType: str = ""
    entity: str = Field(..., min_length=1)
    # Who actually raised/requested this invoice internally -- distinct from
    # clientMailTo (the CLIENT's email, used for recipient/that column).
    # Written into the tracker's "Invoice Advised By" column (key
    # "requested_by"). Previously that column was populated from
    # clientMailTo as a placeholder; this is the real source now that the
    # frontend sends it.
    raisedByEmail: str = ""

    @model_validator(mode="before")
    @classmethod
    def _null_optional_strings_to_empty(cls, data):
        """See NULLABLE_TO_EMPTY_FIELDS above for why this exists."""
        if isinstance(data, dict):
            for key in NULLABLE_TO_EMPTY_FIELDS:
                if key in data and data[key] is None:
                    data[key] = ""
        return data


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
        "client_country": request.company_location.strip(),
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
    logger.info(
        "Invoice save request received: PF ID=%s, entity=%s, company_location=%r",
        request.pfId, request.entity, request.company_location,
    )
    if not request.company_location.strip():
        # Not an error -- a blank/missing value is deliberately treated as
        # "foreign" (see tax.tax_calculator), so this never blocks a save.
        # But it's worth a visible WARNING (not just silence) since a wrong
        # field name/transport on the frontend's side would look exactly
        # like this -- every invoice quietly getting the foreign tax rate
        # with no error anywhere. If you're expecting a value here and see
        # this warning instead, check what the frontend is actually sending.
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
        # tax.tax_calculator.compute_tax(). "invoice_value" above is the
        # row's own "Total Amount" column, which (2026-08-24) is the
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
