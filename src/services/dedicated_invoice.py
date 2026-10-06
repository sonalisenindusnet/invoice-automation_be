"""
dedicated_invoice.py

Handler for dedicated invoice generation with multiple resources.
Consolidates resources by email and generates line items for PDF.

As of 2026-10-06, per explicit instruction: resource email is mandatory on
every resource (enforced on `ResourceItem` in models/invoice_models.py, not
here -- this module just relies on it always being a real, non-blank
value), and resources sharing the same email are summed into ONE merged
line rather than kept as separate line items -- "if 2 email is same then
sum the money against the single resource name... we show the client the
final result only this resource." The merged line's name/PF ID come from
whichever of the matching resources was listed first; only the amount is
summed.

Moved here from save_api/dedicated_invoice.py on 2026-10-06 as part of the
api/models/services/utils restructure; also dropped the unused
build_line_items_for_pdf() helper (defined, never called anywhere in the
repo -- the draft-mailer poller reconstructs PDF line items straight from
parse_consolidated_description() instead, see services/poller.py).

Entry point:
    process_dedicated_invoice(request_data) -> dict with consolidated resources
"""
from typing import List, Dict, Any


class DedicatedResource:
    """Represents a single resource in a dedicated invoice."""

    def __init__(self, pf_id: str, resource_name: str, invoice_amount: float, resource_email: str = ""):
        self.pf_id = pf_id
        self.resource_name = resource_name
        self.invoice_amount = float(invoice_amount)
        self.resource_email = resource_email.strip().lower()


def consolidate_resources_by_email(resources: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Groups resources by email (case-insensitive, trimmed -- see
    DedicatedResource.__init__) and SUMS their invoice amounts into one
    merged entry per unique email, in the order each email was first seen.
    If the same resource email appears on more than one line item (e.g. two
    separate time entries for the same person), the client-facing invoice
    shows ONE combined line for that person, not two -- this is what makes
    grouping by email meaningful rather than just a label: resource email
    is mandatory on every item precisely so this grouping key is always
    reliable (see ResourceItem in models/invoice_models.py).

    A merged entry's `resource_name`/`pf_id` come from whichever of its
    matching resources was listed first in the request; only
    `invoice_amount` is summed across all of them. (If the same email is
    ever sent with genuinely different resource names, that's a data
    problem on the caller's side -- this function doesn't try to guess
    which name is "right," it just keeps the first one, same policy as a
    plain dict/group-by would.)

    Returns:
        {
            "consolidated_resources": [
                {
                    "pf_id": "2504/Insurance/1473",
                    "resource_name": "Swarna Sekhar Dhar",
                    "invoice_amount": 1000.0,
                    "resource_email": "swarna@example.com"
                },
                {
                    "pf_id": "2507/Insurance/2072",
                    "resource_name": "Ipsita Langal",
                    "invoice_amount": 2000.0,
                    "resource_email": "ipsita@example.com"
                },
                ...
            ],
            "description": "Swarna Sekhar Dhar - $1000.00; Ipsita Langal - $2000.00; ..."
        }

    (If two input resources above had shared one email, there would be only
    one entry in "consolidated_resources" for them, with "invoice_amount"
    equal to their sum -- and "description" would have one "Name - $Amount"
    segment for the pair, not two.)
    """

    groups: Dict[str, Dict[str, Any]] = {}  # normalized email -> merged entry
    order: List[str] = []  # first-seen order of each unique email

    for resource in resources:
        res = DedicatedResource(
            pf_id=resource.get("pfId", ""),
            resource_name=resource.get("resourceName", ""),
            invoice_amount=resource.get("invoiceAmount", 0),
            resource_email=resource.get("resouceEmail", ""),  # Note: typo in the JSON is "resouceEmail"
        )

        key = res.resource_email
        if key not in groups:
            groups[key] = {
                "pf_id": res.pf_id,
                "resource_name": res.resource_name,
                "invoice_amount": res.invoice_amount,
                "resource_email": res.resource_email,
            }
            order.append(key)
        else:
            # Same email seen again -- sum the money into the existing
            # entry, keep the first-seen name/PF ID untouched.
            groups[key]["invoice_amount"] += res.invoice_amount

    consolidated = [groups[key] for key in order]

    # Build description for Excel with amounts for verification -- one
    # merged line per unique email: "Resource Name - $Amount"
    description_parts = [
        f"{r['resource_name']} - ${r['invoice_amount']:.2f}" for r in consolidated
    ]
    # Create description string for Excel: "Resource 1 - $Amount1; Resource 2 - $Amount2"
    description = "; ".join(description_parts)

    return {
        "consolidated_resources": consolidated,
        "description": description,
    }


def parse_consolidated_description(description: str) -> List[Dict[str, Any]]:
    """
    Parses the consolidated description format from Excel to extract resource names and amounts.
    Format: "Resource Name 1 - $Amount1; Resource Name 2 - $Amount2; ..."

    Returns list of dicts with 'name' and 'amount' keys, or empty list if not in dedicated format.
    """
    if not description or " - $" not in description:
        return []

    resources = []
    items = description.split(";")

    for item in items:
        item = item.strip()
        if not item or " - $" not in item:
            continue

        # Split by " - $" to separate name from amount
        parts = item.rsplit(" - $", 1)
        if len(parts) != 2:
            continue

        try:
            name = parts[0].strip()
            # Remove $ from the amount string if present
            amount_str = parts[1].strip().lstrip('$')
            amount = float(amount_str)
            resources.append({"name": name, "amount": amount})
        except (ValueError, IndexError):
            continue

    return resources


def is_dedicated_invoice(description: str) -> bool:
    """
    Checks if this invoice description is from a dedicated invoice.
    Dedicated invoices have format: "Resource Name - $Amount[; Resource Name - $Amount ...]".

    CHANGED 2026-10-06: used to require at least 2 semicolon-separated
    entries, specifically to avoid mistaking a normal invoice's free-text
    description for a dedicated one if it happened to contain " - $" once,
    coincidentally. Now that same-email resources are summed into ONE
    merged line before this ever reaches the tracker (see
    consolidate_resources_by_email()), a real dedicated invoice can
    legitimately end up with just a single entry here -- e.g. every
    resource on the request shared one email. Requiring 2+ would wrongly
    fall back to the plain-description PDF/email path for that case,
    losing the resource name entirely. Dropped to >= 1: this column
    (`resource_description`) is exclusively populated by the
    dedicated-invoice-generation endpoint -- a normal invoice never writes
    to it at all -- so even a single well-formed "Name - $Amount" entry
    here unambiguously means "dedicated invoice," never a coincidence.
    """
    if not description:
        return False
    if " - $" not in description:
        return False
    # Count how many resources we can parse
    items = description.split(";")
    count = sum(1 for item in items if " - $" in item.strip())
    return count >= 1  # A single merged resource still counts as dedicated
