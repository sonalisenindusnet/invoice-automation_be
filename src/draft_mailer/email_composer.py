"""
email_composer.py

Turns a tracker row into {to, cc, subject, body}. Subject is deterministic;
body is LLM-drafted (llm_drafter.py) with a plain-template fallback if that
call fails for any reason -- a flaky API call should never block a draft
from being created.
"""
import logging

from draft_mailer.llm_drafter import draft_invoice_email_body_llm

logger = logging.getLogger("draft_mailer.email_composer")


def _recipients(row):
    to_list = [e.strip() for e in (row.get("client_mail_to") or "").split(",") if e.strip()]
    cc_list = [e.strip() for e in (row.get("int_cc_mail") or "").split(",") if e.strip()]
    return to_list, cc_list


def _build_subject(row, description):
    return f"Invoice {row.get('invoice_no')} - {row.get('client_company')} - {description}".strip()


def _fallback_body(row, description):
    total = row.get("total")
    currency = row.get("currency") or ""
    total_str = f"{currency} {float(total):,.2f}" if total not in (None, "") else ""
    return f"""Dear Team,

Please find attached the invoice {row.get('invoice_no')} for {description}.

Total Payable: {total_str}

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
