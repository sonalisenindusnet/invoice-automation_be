"""
llm_drafter.py

Drafts the client-facing invoice email BODY with Gemini. The subject line
is deterministic (see email_composer.py) -- only the body is LLM-written.

Reliability: draft_invoice_email_body_llm() raises on any failure (missing
API key, API error, malformed JSON reply) so the caller can fall back to a
plain template instead of blocking the draft from being created.

Setup: set GEMINI_API_KEY in .env. Optionally set GEMINI_MODEL to override
the model name below.

Uses the google-genai SDK's Interactions API (`client.interactions.create()`
/ `.output_text`).
"""
import json
import logging
import os

from google import genai

logger = logging.getLogger("draft_mailer.llm_drafter")

DEFAULT_MODEL = "gemini-3.6-flash"


def _build_prompt(fields):
    # NOTE: only None/"" are dropped here -- a genuine 0 (e.g. a 0% tax
    # rate) is real, known information, not a missing field, so it must
    # stay in the prompt rather than be silently filtered out.
    known = "\n".join(f"- {label}: {value}" for label, value in fields if value not in (None, ""))
    return f"""Draft a professional, concise invoice email BODY (a complete email,
including greeting and sign-off -- NOT a subject line).

Invoice information (this is everything known -- do not invent anything else):
{known}

Rules:
1. Do not invent or assume any missing facts.
2. Always mention the invoice number and the total payable.
3. If a tax rate/amount is listed above, mention the tax name and rate explicitly, even if it is 0% -- do not omit it.
4. Do not claim that payment has been received or completed -- this is a request for payment.
5. Keep it professional, polite, and concise. Address the client contact by name if listed, otherwise use a generic greeting. Sign off as "Accounts Team, INTglobal".
6. Return ONLY valid JSON with exactly this field: {{"body": "..."}}
"""


def _call_gemini(prompt, model):
    client = genai.Client()
    model_name = model or os.environ.get("GEMINI_MODEL", DEFAULT_MODEL)
    try:
        interaction = client.interactions.create(model=model_name, input=prompt, store=False)
    except Exception as exc:
        logger.error("Gemini API request failed: %s", exc)
        raise RuntimeError(f"Gemini API request failed: {exc}") from exc
    return (getattr(interaction, "output_text", None) or "").strip()


def _parse_llm_json(text):
    if text.startswith("```"):
        text = text.strip("`").strip()
        if text.lower().startswith("json"):
            text = text[4:].strip()
    try:
        # strict=False: models sometimes emit a literal newline inside the
        # "body" string instead of an escaped \n.
        draft = json.loads(text, strict=False)
    except json.JSONDecodeError as exc:
        logger.error("Gemini returned invalid JSON: %s", text)
        raise RuntimeError("Gemini returned an invalid email draft.") from exc

    body = draft.get("body")
    if not body:
        raise RuntimeError("Gemini response did not include a body.")
    return body


def draft_invoice_email_body_llm(row, model=None):
    """Returns the drafted body text (str). Raises RuntimeError on any
    failure so the caller can fall back to a static template."""
    if not os.environ.get("GEMINI_API_KEY"):
        raise RuntimeError("GEMINI_API_KEY is not set -- cannot draft the email body with the LLM.")

    tax_name = row.get("tax_name")
    fields = [
        ("Client company", row.get("client_company") or ""),
        ("Invoice number", row.get("invoice_no") or ""),
        ("Invoice date", row.get("invoice_date") or ""),
        ("Due date", row.get("due_date") or ""),
        ("Currency", row.get("currency") or ""),
        # "subtotal" (pre-tax, set by poller.process_row() alongside the
        # other tax_* fields) is deliberately distinct from row["total"],
        # which is the tracker's own POST-tax grand total -- see
        # excel_writer.build_row(). Using "total" here would tell the LLM
        # the tax-inclusive total is the pre-tax figure.
        ("Sub-Total (before tax)", row.get("subtotal") if row.get("subtotal") not in (None, "") else ""),
        ("Description", row.get("invoice_description") or ""),
    ]
    if tax_name:
        # Appended unconditionally (not filtered by _build_prompt's normal
        # None/"" check) so a genuine 0% rate still reaches the prompt --
        # see llm_drafter._build_prompt's rule #3 and email_composer's
        # _payable_block(), both of which state the tax explicitly either way.
        fields.append((f"{tax_name} rate", f"{round((row.get('tax_rate') or 0) * 100)}%"))
        fields.append((f"{tax_name} amount", row.get("tax_amount") if row.get("tax_amount") not in (None, "") else 0))
        fields.append(("Total payable (incl. tax)", row.get("total_with_tax") if row.get("total_with_tax") not in (None, "") else ""))
    else:
        fields.append(("Total payable", row.get("total") if row.get("total") not in (None, "") else ""))
    prompt = _build_prompt(fields)

    text = _call_gemini(prompt, model)
    body = _parse_llm_json(text)

    logger.info("LLM-drafted email body for %s", row.get("invoice_no"))
    return body
