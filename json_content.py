"""
json_content.py  – all JSON-building logic lives here
Edit this file whenever you need to change what gets sent to Acelynk.
"""

from __future__ import annotations
import math, numpy as np
from typing import Any


# ───────── tiny helpers ──────────
def _is_nan(v: Any) -> bool:
    return (
        v is None
        or (isinstance(v, float) and math.isnan(v))
        or (isinstance(v, np.floating) and np.isnan(v))
        or (isinstance(v, str) and v.strip() == "")
    )


def _strip_nans(o: Any) -> Any:
    """Recursively drop empty / NaN branches so we don’t send nulls."""
    if isinstance(o, dict):
        out = {k: _strip_nans(v) for k, v in o.items() if not _is_nan(v)}
        return {k: v for k, v in out.items() if v != {}}
    if isinstance(o, list):
        return [_strip_nans(v) for v in o if not _is_nan(v)]
    return o


def _clean_tariff(v: Any) -> str:
    digits = "".join(c for c in str(v) if c.isdigit())
    return digits.zfill(10)[:10]


# ───────── build_entry() is imported by api_ui.py ─────────
def build_entry(head_row) -> dict:
    """
    Return a *header-only* Entry Summary payload.
    `head_row` is the first row of the split Excel (Pandas Series).
    """
    invoice = str(head_row["Invoice_No"]).strip()

    # --- Importer (minimal) ---
    importer = {
        "Account": "SHEI00001",          # <-- static per Magaya
        "Filer":   "D9T",
        "Tax_id":  "86-371698000"
    }

    # --- General block ---
    general = {
        "Invoice_number":        invoice,
        "Invoice_date":          "06/01/2025",
        "Entry_filing_type":     "01",
        "Mode_of_transportation":"40",
        "Payment_type":          "7",

        "Entry_port":       "3901",
        "Port_of_lading":   "3901",
        "Port_of_unlading": "3901",
        "Firms":            "HB61",
        "broker_ref_num":   "TESTBROKER",

        "Scac":             "K4",
        "vessel_flight_no": "0933",

        "gross_weight":         "3099",
        "Charges":              "3099",
        "manifest_description": "TANK TOP",

        "Anticipated_arrival": { "Date": "06/05/2025" },
        "export_date":     "06/04/2025",
        "origin_country":  "CN",
        "export_country":  "CN",

        "BillLading": [{
            "BillType": "M",
            "BillNo":   invoice,
            "Quantity": 17,
            "QuantityUOM": "CTN",
            "SCAC": "K4",
            "HouseBill": [{
                "BillNo": invoice[3:],     # strip first 3 chars
                "Quantity": 17,
                "QuantityUOM": "CTN",
                "SCAC": "K4"
            }]
        }]
    }

    # --- Parties (new) ---
    parties = {
        "Seller":    { "Name": "ROADGET BUSINESS PTE. LTD." },
        "Consignee": { "Name": "SHEIN DISTRIBUTION CORPORATION" },
        "Buyer":     { "Name": "SHEIN DISTRIBUTION CORPORATION" }
    }

    # --- Header ---
    header = {
        "Importer_of_record": importer,
        "General":            general,
        "Parties":            parties           # <── added here
    }

    # --- No line items yet (empty list) ---
    shipment = {
        "Reference": invoice,
        "Header":    header,
        "Line":      { "Item": [] }
    }

    return _strip_nans({ "Shipment": [shipment] })
