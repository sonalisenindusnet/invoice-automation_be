"""
dedicated_invoice.py

Handler for dedicated invoice generation with multiple resources.
Consolidates resources by email and generates line items for PDF.

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
    Groups resources by email for tracking. Each resource becomes a separate line item
    in the PDF with its own serial number (1, 2, 3...).

    Returns:
        {
            "consolidated_resources": [
                {
                    "pf_id": "2504/Insurance/1473",
                    "resource_name": "Swarna Sekhar Dhar",
                    "invoice_amount": 1000.0,
                    "resource_email": ""
                },
                {
                    "pf_id": "2507/Insurance/2072",
                    "resource_name": "Ipsita Langal",
                    "invoice_amount": 2000.0,
                    "resource_email": ""
                },
                ...
            ],
            "description": "Swarna Sekhar Dhar - $1000.00; Ipsita Langal - $2000.00; ..."
        }
    """

    consolidated = []
    description_parts = []

    for resource in resources:
        res = DedicatedResource(
            pf_id=resource.get("pfId", ""),
            resource_name=resource.get("resourceName", ""),
            invoice_amount=resource.get("invoiceAmount", 0),
            resource_email=resource.get("resouceEmail", "")  # Note: typo in the JSON is "resouceEmail"
        )

        consolidated.append({
            "pf_id": res.pf_id,
            "resource_name": res.resource_name,
            "invoice_amount": res.invoice_amount,
            "resource_email": res.resource_email,
        })

        # Build description for Excel with amounts for verification
        # Each resource as individual line: "Resource Name - $Amount"
        description_parts.append(f"{res.resource_name} - ${res.invoice_amount:.2f}")

    # Create description string for Excel: "Resource 1 - $Amount1; Resource 2 - $Amount2"
    description = "; ".join(description_parts)

    return {
        "consolidated_resources": consolidated,
        "description": description,
    }


def build_line_items_for_pdf(consolidated_resources: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Builds line items for PDF generation from resources.
    Each resource becomes a separate line item with its own serial number.

    Returns:
        [
            {
                "label": "Swarna Sekhar Dhar",
                "amount": 1000.0
            },
            {
                "label": "Ipsita Langal",
                "amount": 2000.0
            },
            ...
        ]
    """
    line_items = []

    for resource in consolidated_resources:
        label = resource['resource_name']

        line_items.append({
            "label": label,
            "amount": resource["invoice_amount"],
        })

    return line_items


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
    Dedicated invoices have format: "Resource Name - $Amount; Resource Name - $Amount"
    Must have at least 2 items separated by semicolon.
    """
    if not description:
        return False
    if " - $" not in description or ";" not in description:
        return False
    # Count how many resources we can parse
    items = description.split(";")
    count = sum(1 for item in items if " - $" in item.strip())
    return count >= 2  # Must have at least 2 resources to be a dedicated invoice
