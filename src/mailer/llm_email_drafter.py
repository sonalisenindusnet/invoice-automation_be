"""
llm_email_drafter.py

Added 2026-08-14, per explicit instruction ("use the LLM to generate the
message [body] only"), adapted from a sample script provided directly by
the user. Drafts the client-facing invoice email BODY with an LLM (Google
Gemini) instead of the fixed template previously in compose_email()
(draft_email_from_excel_row.py). The subject line stays the existing
deterministic format -- predictable and searchable in a mail client, and
"the message" in the instruction reads as the body, not the subject; say
the word if you want the subject LLM-drafted too.

Trimmed from the user's sample to match what USA/UK/Poland actually need:
CGST/SGST were dropped entirely -- those are India/GST-specific fields,
and none of the three live entities ever charge GST (see
entity_resolver.py's module docstring: "Foreign clients only... none of
these three entities charge GST"). Only VAT (UK/Poland; always 0 for USA,
so never mentioned there) is in scope.

Reliability: a live email to a real client should not silently fail just
because an LLM call hiccups (rate limit, auth, network, a malformed JSON
reply). draft_invoice_email_body_llm() raises on any such failure --
compose_email() catches that and falls back to the previous deterministic
template, logging a warning -- so a flaky API call degrades the email's
polish, never blocks it from going out.

Setup: set GEMINI_API_KEY (same .env-driven pattern as EMAIL_ADDRESS /
EMAIL_APP_PASSWORD -- see utils/env_loader.py and .env.example). Optionally
set GEMINI_MODEL to override the model name below -- kept configurable
rather than hardcoded, since model availability changes over time (see
correction note right below).

CORRECTION, 2026-08-14: an earlier version of this file called
`client.models.generate_content(...)` instead of the user's original
`client.interactions.create(...)` / `.output_text`, on the assumption the
latter was mixed up with OpenAI's Responses API. That was wrong -- a live
run hit `404 This model models/gemini-2.0-flash is no longer available...
We recommend you to use the Interactions API`, which is Google's own
server confirming the Interactions API is the real, current one.
Confirmed directly against the installed google-genai SDK (2.18.1):
`client.interactions` is a real resource with a `.create()` method, its
response type has a real `.output_text` field, and 'gemini-3.6-flash' is
one of the SDK's own recognized model literals. Switched back to exactly
the user's original call shape below.
"""
import json
import logging
import os

from google import genai

logger = logging.getLogger("email_server.llm_email_drafter")

# gemini-2.0-flash (this file's first guess) turned out to be retired.
# gemini-2.5-flash is used as the fallback default here only because it's
# very likely to still be live if GEMINI_MODEL is ever unset -- but
# GEMINI_MODEL in .env (set to gemini-3.6-flash) is what actually
# controls this in practice, and takes priority every time.
DEFAULT_MODEL = "gemini-3.6-flash"


def _build_prompt(fields):
    known = "\n".join(f"- {label}: {value}" for label, value in fields if value not in (None, "", 0, 0.0))
    return f"""Draft a professional, concise invoice email BODY (a complete email,
including greeting and sign-off -- NOT a subject line).

Invoice information (this is everything known -- do not invent anything else):
{known}

Rules:
1. Do not invent or assume any missing facts.
2. Always mention the invoice number and the total payable.
3. Mention the base amount only if it's listed above and differs from the total payable.
4. Mention VAT only if a non-zero VAT value is listed above -- never invent or calculate a tax amount, and never mention VAT if it isn't listed.
5. If remarks are listed above, incorporate them naturally. If not listed, do not mention remarks at all.
6. Do not claim that payment has been received or completed -- this is a request for payment.
7. Keep it professional, polite, and concise. Address the client contact by name if listed. Sign off as "Accounts Team, INTglobal".
8. Return ONLY valid JSON with exactly this field: {{"body": "..."}}
"""


def draft_invoice_email_body_llm(row, model=None):
    """Returns the drafted body text (str). Raises RuntimeError on any
    failure -- missing API key, API error, or a malformed/incomplete JSON
    reply -- so the caller (compose_email) can fall back to the static
    template instead of blocking the invoice from going out."""
    if not os.environ.get("GEMINI_API_KEY"):
        raise RuntimeError("GEMINI_API_KEY is not set -- cannot draft the email body with the LLM.")

    base_amount = row.get("base_amount")
    total = row.get("total")
    vat = row.get("vat")
    fields = [
        ("Client contact", row.get("client_contact_person") or ""),
        ("Client company", row.get("client_company") or ""),
        ("Invoice number", row.get("invoice_no") or ""),
        ("Invoice date", row.get("invoice_date") or ""),
        ("Due date", row.get("due_date") or ""),
        ("Currency", row.get("currency") or ""),
        ("Base amount", base_amount if base_amount not in (None, "") else ""),
        ("Total payable", total if total not in (None, "") else ""),
        ("VAT", vat if vat not in (None, "", 0, 0.0) else ""),
        ("Remarks", row.get("remarks") or ""),
    ]
    prompt = _build_prompt(fields)

    client = genai.Client()
    model_name = model or os.environ.get("GEMINI_MODEL", DEFAULT_MODEL)
    try:
        interaction = client.interactions.create(
            model=model_name,
            input=prompt,
            store=False,
        )
    except Exception as exc:
        logger.error("Gemini API request failed: %s", exc)
        raise RuntimeError(f"Gemini API request failed: {exc}") from exc

    text = (getattr(interaction, "output_text", None) or "").strip()
    if text.startswith("```"):
        # Models sometimes fence the JSON in ```json ... ``` despite being
        # asked for raw JSON -- strip that before parsing rather than
        # failing on it.
        text = text.strip("`").strip()
        if text.lower().startswith("json"):
            text = text[4:].strip()

    try:
        # strict=False: models frequently emit a literal newline inside the
        # "body" string value instead of an escaped \n, even when asked for
        # valid JSON -- strict JSON parsing rejects that as a control
        # character. This is a real, observed failure mode (caught while
        # testing this module), not a hypothetical one.
        draft = json.loads(text, strict=False)
    except json.JSONDecodeError as exc:
        logger.error("Gemini returned invalid JSON: %s", text)
        raise RuntimeError("Gemini returned an invalid email draft.") from exc

    body = draft.get("body")
    if not body:
        raise RuntimeError("Gemini response did not include a body.")

    logger.info("LLM-drafted email body for %s", row.get("invoice_no"))
    return body
