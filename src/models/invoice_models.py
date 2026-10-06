"""
invoice_models.py

Pydantic request schemas for the save API (api/routes.py). Extracted out of
routes.py on 2026-10-06 as part of the api/models/services/utils restructure
-- these are pure data-shape definitions with no HTTP or business logic of
their own, so they live in their own layer rather than mixed into the route
handlers.
"""
from typing import List

from pydantic import AliasChoices, BaseModel, Field, model_validator

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
    "raisedByEmail", "companyAddress",
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
    clientType: str = ""
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
    # SGD, USD, EURO, ... depending on the client), but applies to every
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
    companyAddress: str = ""

    @model_validator(mode="before")
    @classmethod
    def _null_optional_strings_to_empty(cls, data):
        """See NULLABLE_TO_EMPTY_FIELDS above for why this exists."""
        if isinstance(data, dict):
            for key in NULLABLE_TO_EMPTY_FIELDS:
                if key in data and data[key] is None:
                    data[key] = ""
        return data


class ResourceItem(BaseModel):
    pfId: str = Field(..., min_length=1)
    resourceName: str = Field(..., min_length=1)
    invoiceAmount: str = Field(..., min_length=1)
    # Mandatory as of 2026-10-06, per explicit instruction -- this is the
    # grouping key services.dedicated_invoice.consolidate_resources_by_email()
    # sums same-resource amounts on, so a blank/missing email would silently
    # defeat that grouping (every blank would collide into one group, or
    # none would ever merge, depending on how it's handled -- rather than
    # guess, this is now required and a request without it gets a loud 422).
    resouceEmail: str = Field(..., min_length=1)  # Note: typo preserved to match frontend


class DedicatedInvoiceGenerationRequest(BaseModel):
    pfId: str = Field(..., min_length=1)
    clientName: str = Field(..., min_length=1)
    contactPersonName: str = ""
    clientMailTo: str = ""
    clientMailCc: str = ""
    intCcMailId: str = ""
    clientType: str = ""
    raisedByEmail: str = ""
    companyLocation: str = ""
    companyAddress: str = ""
    companyGeography: str = ""
    entity: str = Field(..., min_length=1)
    currency: str = Field(..., min_length=1)
    invoiceValue: str = Field(..., min_length=1)
    invoiceDescription: str = ""
    resources: List[ResourceItem] = Field(..., min_items=1)

    @model_validator(mode="before")
    @classmethod
    def _null_optional_strings_to_empty(cls, data):
        if isinstance(data, dict):
            nullable_fields = (
                "contactPersonName", "clientMailTo", "clientMailCc", "intCcMailId",
                "clientType", "raisedByEmail", "companyLocation", "companyAddress",
                "companyGeography", "invoiceDescription",
            )
            for key in nullable_fields:
                if key in data and data[key] is None:
                    data[key] = ""
        return data


class MisVerificationUpdateRequest(BaseModel):
    invoiceNo: str = Field(..., min_length=1)
    entity: str = Field(..., min_length=1)
    pfId: str = ""
    misUpdateFlag: bool
