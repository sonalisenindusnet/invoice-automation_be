"""
email_composer.py

Turns a tracker row into {to, cc, subject, body}. Subject is deterministic;
body is LLM-drafted (llm_drafter.py) with a plain-template fallback if that
call fails for any reason -- a flaky API call should never block a draft
from being created.

Moved here from draft_mailer/email_composer.py on 2026-10-06 as part of
the api/models/services/utils restructure; content/behavior unchanged.
"""
import logging

from services.llm_drafter import draft_invoice_email_body_llm

logger = logging.getLogger("draft_mailer.email_composer")


def _recipients(row):
    to_list = [e.strip() for e in (row.get("client_mail_to") or "").split(",") if e.strip()]
    cc_list = [e.strip() for e in (row.get("int_cc_mail") or "").split(",") if e.strip()]
    return to_list, cc_list


def _build_subject(row, description):
    return f"Invoice {row.get('invoice_no')} - {row.get('client_company')} - {description}".strip()


def _money(value, currency):
    try:
        return f"{currency} {float(value):,.2f}".strip()
    except (TypeError, ValueError):
        return ""


def _payable_block(row):
    """Sub-Total / Tax / Total breakdown when tax fields are present (see
    poller.process_row(), which always sets them before calling
    compose_email() now) -- matches the PDF's own Sub-Total -> Tax -> Total
    structure, and states the tax name/rate explicitly even when it's 0%,
    per explicit instruction ("the tax name is important for the draft
    email body"). Falls back to a single "Total Payable" line if this row
    has no tax fields at all (e.g. a direct/test call to compose_email()).

    NOTE: the Sub-Total line reads `row["subtotal"]` (the PRE-tax amount,
    set by poller.process_row() alongside the other tax fields) --
    deliberately NOT `row["total"]`, which is the tracker's own POST-tax
    grand total (see excel_writer.build_row()). Using "total" here would
    double-count the tax in what's shown as the Sub-Total."""
    currency = row.get("currency") or ""
    tax_name = row.get("tax_name")
    if not tax_name:
        return f"Total Payable: {_money(row.get('total'), currency)}"

    tax_rate_pct = round((row.get("tax_rate") or 0) * 100)
    return (
        f"Sub-Total: {_money(row.get('subtotal'), currency)}\n"
        f"{tax_name} ({tax_rate_pct}%): {_money(row.get('tax_amount'), currency)}\n"
        f"Total Payable: {_money(row.get('total_with_tax'), currency)}"
    )


def _fallback_body(row, description):
    return f"""Dear Team,

Please find attached the invoice {row.get('invoice_no')} for {description}.

{_payable_block(row)}

Kindly process the payment at your earliest convenience. Please let us know if any further information or supporting documents are required.

Best regards,
Accounts Team
INTglobal
"""


def compose_email(row):
    to_list, cc_list = _recipients(row)
    description = row.get("invoice_description") or "Services"
    subject = _build_subject(row, description)

    try:
        body = draft_invoice_email_body_llm(row)
    except Exception as exc:
        logger.warning(
            "LLM email draft failed for %s (%s) -- falling back to the standard template",
            row.get("invoice_no"), exc,
        )
        body = _fallback_body(row, description)

    return {"to": to_list, "cc": cc_list, "subject": subject, "body": body}
