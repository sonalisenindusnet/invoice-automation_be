"""
tax_calculator.py

Single source of truth for invoice tax (GST/VAT) calculation across every
entity. Per explicit instruction, tax depends on whether the CLIENT is in
the same country as the entity the invoice is being raised FROM:

    Singapore entity + Singapore client  -> 9%  GST
    Singapore entity + any other client  -> 0%  GST
    UK entity        + UK client         -> 20% VAT
    UK entity        + any other client  -> 0%  VAT
    Poland entity     + any client       -> 0%  VAT   (never varies)
    USA entity        + any client       -> 0%  VAT   (never varies)

Rates are read from environment variables (see _RATE_ENV below), per
explicit instruction ("tax percentage will store in ENV"), so they can be
changed without a code deploy. Every default below matches the rule table
above exactly, so an unset .env still produces the correct result.

Nothing here ever raises for a missing/unrecognized client_country -- it's
simply treated as "not local" (the foreign rate), same as an empty string.
This fails toward under- rather than over-charging tax, which is the safer
direction for a mistake to fall in.

Country matching (see is_local_client()) tolerates common real-world
variation on its own -- punctuation ("U.K.", "U.S.A."), case, extra
whitespace, and a compound value like "England, UK" or "London / United
Kingdom" (checked segment-by-segment on top of the whole string). It
deliberately never does substring/"contains" matching on short codes --
"uk" must match a whole segment, not appear inside one -- specifically so
"Ukraine" can never be mistaken for "UK", nor "Uspallata" for "US".

As of 2026-08-24, the frontend's client-location field is a fixed
dropdown -- confirmed to only ever send exactly one of "United States",
"United Kingdom", "Singapore", "Poland" (see KNOWN_DROPDOWN_COUNTRIES).
All four already match correctly against the alias sets above with no
changes needed. Since the value is no longer free text, anything else
arriving here is a real signal something's wrong (a dropdown/frontend
change, an encoding issue, a new entity nobody told this module about) --
see is_recognized_country() / compute_tax()'s "country_recognized" field,
which the save API logs a warning on.

As of 2026-08-24 (later the same day): the tracker's own "Total Amount"
column is confirmed to mean the POST-tax grand total (Poland's real
column header literally says "Total Amount (Including VAT)"), with the
tax amount itself stored separately in that tab's own VAT/GST column --
NOT "Total Amount stays the pre-tax subtotal" as an earlier version of
this module's callers assumed. See compute_tax() (used once, at SAVE
time, from the raw pre-tax amount the frontend sent) vs.
tax_result_from_stored() (used at DRAFT time, reconstructing the same
breakdown from what was actually saved) for how this is now split.

Moved here from tax/tax_calculator.py on 2026-10-06 as part of the
api/models/services/utils restructure; content/behavior unchanged.
"""
import os
import re

# Tax name shown on the invoice / in API responses / in drafted emails,
# per entity. Fixed by each country's actual tax law -- not configurable.
TAX_NAME_BY_ENTITY = {
    "singapore": "GST",
    "uk": "VAT",
    "poland": "VAT",
    "usa": "VAT",
}

# Country name(s) that count as "local" for each entity's own tax rule,
# matched case-insensitively after stripping whitespace. Extend a set here
# if the frontend ever sends a spelling/code not covered yet.
_LOCAL_COUNTRY_ALIASES = {
    "singapore": {"singapore", "sg"},
    "uk": {"uk", "united kingdom", "gb", "great britain", "england", "scotland", "wales", "northern ireland"},
    "poland": {"poland", "pl"},
    "usa": {"usa", "us", "united states", "united states of america"},
}

# The exact 4 values the frontend's client-location dropdown sends today
# (confirmed 2026-08-24) -- one per entity's own country. Kept here purely
# as documentation/for is_recognized_country() below; is_local_client()
# itself still goes through the alias sets above, not this set, since a
# dropdown value being "known" and being "local to THIS entity" are two
# different questions.
KNOWN_DROPDOWN_COUNTRIES = {"United States", "United Kingdom", "Singapore", "Poland"}

# ENV var name + default for each entity's LOCAL rate (client is in the
# same country as the entity) and FOREIGN rate (client is anywhere else).
# FOREIGN defaults to 0% for every entity today per the rule table above,
# but is still its own env var (not hardcoded) so this can change later
# without a code deploy.
_RATE_ENV = {
    "singapore": {"local": ("TAX_RATE_SINGAPORE_LOCAL", 0.09), "foreign": ("TAX_RATE_SINGAPORE_FOREIGN", 0.0)},
    "uk": {"local": ("TAX_RATE_UK_LOCAL", 0.20), "foreign": ("TAX_RATE_UK_FOREIGN", 0.0)},
    "poland": {"local": ("TAX_RATE_POLAND_LOCAL", 0.23), "foreign": ("TAX_RATE_POLAND_FOREIGN", 0.0)},
    "usa": {"local": ("TAX_RATE_USA_LOCAL", 0.0), "foreign": ("TAX_RATE_USA_FOREIGN", 0.0)},
}


class UnknownTaxEntityError(ValueError):
    """Raised when there's no tax rule configured for the given entity."""


def _rate_from_env(env_name, default):
    raw = os.environ.get(env_name)
    if raw is None or not raw.strip():
        return default
    try:
        return float(raw.strip())
    except ValueError:
        return default


# Parentheses become a segment boundary (so "United Kingdom (UK)" splits
# into "united kingdom" and "uk" rather than merging into one run-on
# string with no separator between them) -- see _SEGMENT_SPLIT_RE below.
_PAREN_TO_COMMA_RE = re.compile(r"[()]")
# Other punctuation stripped out entirely -- periods/apostrophes so
# "U.K." / "U.S.A." collapse to "uk" / "usa".
_PUNCT_STRIP_RE = re.compile(r"[.’']")
# Characters that separate independent alternatives within one string --
# e.g. "England, UK" or "London / United Kingdom" -- each side is checked
# on its own in addition to the string as a whole.
_SEGMENT_SPLIT_RE = re.compile(r"[,/|]")


def _normalize_country(raw):
    s = (raw or "").strip().lower()
    s = _PAREN_TO_COMMA_RE.sub(",", s)
    s = _PUNCT_STRIP_RE.sub("", s)
    return re.sub(r"\s+", " ", s).strip()


def _country_segments(raw):
    """The normalized whole string, plus each comma/slash-separated piece
    of it, trimmed and de-duplicated. Never splits on whitespace alone --
    a multi-word alias like "united kingdom" must still match as one
    segment, not two separate "united"/"kingdom" tokens."""
    whole = _normalize_country(raw)
    if not whole:
        return []
    pieces = [whole] + [p.strip() for p in _SEGMENT_SPLIT_RE.split(whole)]
    seen = []
    for p in pieces:
        if p and p not in seen:
            seen.append(p)
    return seen


def is_local_client(entity_key, client_country):
    aliases = _LOCAL_COUNTRY_ALIASES.get((entity_key or "").strip().lower(), set())
    if not aliases:
        return False
    # Exact match only, per segment -- never "contains"/substring matching,
    # so a short code like "uk"/"us"/"sg"/"pl" can only match when it IS
    # the whole segment, never when it merely appears inside a longer word
    # (e.g. "Ukraine", "Uspallata", "Pluto" must never match).
    return any(segment in aliases for segment in _country_segments(client_country))


# Every alias across every entity, combined -- used only by
# is_recognized_country() below to answer "is this ANY known country at
# all", regardless of which entity the invoice happens to be for.
_ALL_KNOWN_ALIASES = set().union(*_LOCAL_COUNTRY_ALIASES.values())


def is_recognized_country(client_country):
    """True if `client_country` matches ANY entity's alias set -- i.e. it's
    a real, known value, just not necessarily "local" to the entity this
    particular invoice happens to be for. A blank value is NOT considered
    unrecognized (that's the normal "not provided" case, handled
    separately) -- this is for flagging a non-blank value that doesn't
    match anything at all, which now that the frontend uses a fixed
    dropdown (see KNOWN_DROPDOWN_COUNTRIES) should never actually happen."""
    if not (client_country or "").strip():
        return True
    return any(segment in _ALL_KNOWN_ALIASES for segment in _country_segments(client_country))


def compute_tax(entity_key, client_country, subtotal):
    """Returns the full tax breakdown for one invoice:

        {"tax_name": "GST"/"VAT", "rate": float, "is_local": bool,
         "subtotal": float, "tax_amount": float, "total": float,
         "country_recognized": bool}

    `subtotal` is the pre-tax invoice amount. Raises UnknownTaxEntityError
    if `entity_key` isn't one of the four configured entities -- this one
    DOES raise (unlike the client-country handling above) because an
    unrecognized entity means the caller's own routing is broken, not a
    normal missing-data case.

    "country_recognized" is False only for a non-blank client_country that
    doesn't match ANY entity's alias set -- a real signal worth logging
    (see is_recognized_country()), not something this function itself
    reacts to; the tax math always proceeds using the foreign rate either
    way, same as for a blank value."""
    key = (entity_key or "").strip().lower()
    if key not in _RATE_ENV:
        raise UnknownTaxEntityError(f"No tax rule configured for entity '{entity_key}'")

    try:
        subtotal = float(subtotal or 0)
    except (TypeError, ValueError):
        subtotal = 0.0

    local = is_local_client(key, client_country)
    env_name, default = _RATE_ENV[key]["local" if local else "foreign"]
    rate = _rate_from_env(env_name, default)

    tax_amount = round(subtotal * rate, 2)
    total = round(subtotal + tax_amount, 2)

    return {
        "tax_name": TAX_NAME_BY_ENTITY.get(key, "VAT"),
        "rate": rate,
        "is_local": local,
        "subtotal": round(subtotal, 2),
        "tax_amount": tax_amount,
        "total": total,
        "country_recognized": is_recognized_country(client_country),
    }


def tax_result_from_stored(entity_key, client_country, stored_total, stored_tax_amount):
    """Reconstructs the same shape compute_tax() returns, but from a row's
    OWN already-saved Total Amount (post-tax) and VAT/GST column, instead
    of recomputing the tax amount fresh from today's env-var rate.

    Use this at DRAFT time (once a row has already been through
    save_invoice()), not compute_tax() -- review + MIS-verification can
    happen days after a row is saved, and re-deriving the amount from the
    *current* env rate at that point could show the client a different
    figure than what was actually agreed/saved at invoice time if the
    rate changed in between. Reading it back from the row instead means
    the PDF and drafted email always match the sheet exactly, by
    construction, regardless of any later env change.

    `stored_total` is the row's Total Amount (subtotal + tax, per the
    tracker's actual column semantics -- e.g. Poland's real header is
    literally "Total Amount (Including VAT)"). `stored_tax_amount` is the
    row's own VAT/GST column value (0 for an entity/tab with no such
    column, e.g. USA -- meaning subtotal == total there, correctly).
    `subtotal` is derived as `stored_total - stored_tax_amount`, and
    `rate` as `stored_tax_amount / subtotal` (0.0 if subtotal is 0) --
    both exact reconstructions of what was actually charged, not fresh
    lookups. `is_local`/`country_recognized` are still derived live from
    entity+country, same as compute_tax() -- those never depend on the
    rate, so there's no drift risk in computing them fresh.

    Raises UnknownTaxEntityError for an unrecognized entity_key, same as
    compute_tax()."""
    key = (entity_key or "").strip().lower()
    if key not in _RATE_ENV:
        raise UnknownTaxEntityError(f"No tax rule configured for entity '{entity_key}'")

    try:
        stored_total = float(stored_total or 0)
    except (TypeError, ValueError):
        stored_total = 0.0
    try:
        stored_tax_amount = float(stored_tax_amount or 0)
    except (TypeError, ValueError):
        stored_tax_amount = 0.0

    subtotal = round(stored_total - stored_tax_amount, 2)
    rate = round(stored_tax_amount / subtotal, 6) if subtotal else 0.0

    return {
        "tax_name": TAX_NAME_BY_ENTITY.get(key, "VAT"),
        "rate": rate,
        "is_local": is_local_client(key, client_country),
        "subtotal": subtotal,
        "tax_amount": round(stored_tax_amount, 2),
        "total": round(stored_total, 2),
        "country_recognized": is_recognized_country(client_country),
    }
