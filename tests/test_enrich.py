"""Tests for matching SQL site cost, aged status, and SellBelowCost onto offers.

Targets the pure ``enrich_offers`` — the database read (``load_reference_data``)
is exercised by a real run.
"""
from __future__ import annotations

import pandas as pd

import run_ebay_best_offers as script


def _offers(rows):
    """Build the listings frame from ``[item_number, title, sku, price]`` rows."""
    items = [
        {"item_number": r[0], "title": r[1], "sku": r[2],
         "current_price": float(str(r[3]).replace("$", "").replace(",", "")),
         "quantity": 5, "quantity_sold": 1, "quantity_available": 4,
         "listing_status": "Active"}
        for r in rows
    ]
    return script.items_to_frame(items, "Acct", "2026-07-02")


def test_enrich_matches_by_sku_and_fills_missing():
    offers = _offers([
        ["1", "Sony A", "SKU-A", "$100.00", "href"],
        ["2", "Sony B", "SKU-B", "$200.00", "href"],  # no reference row at all
    ])
    site_costs = pd.DataFrame(
        {"sku": ["SKU-A"], "site_cost": [70.1234], "weight_oz": [40.0], "sell_below_cost": [True]})
    aged = pd.DataFrame({"sku": ["SKU-A"], "aged_status": ["Dead"]})

    out = script.enrich_offers(offers, site_costs, aged)
    by_sku = out.set_index("sku")

    assert by_sku.loc["SKU-A", "site_cost"] == 70.12   # rounded to 2 decimals
    assert by_sku.loc["SKU-A", "weight_oz"] == 40.0
    assert by_sku.loc["SKU-A", "aged_status"] == "Dead"
    assert by_sku.loc["SKU-A", "sell_below_cost"]      # True
    # missing -> 0 / 0 oz / N/A / False
    assert by_sku.loc["SKU-B", "site_cost"] == 0
    assert by_sku.loc["SKU-B", "weight_oz"] == 0
    assert by_sku.loc["SKU-B", "aged_status"] == "N/A"
    assert not by_sku.loc["SKU-B", "sell_below_cost"]
    assert len(out) == len(offers)                     # never add or drop rows


def test_enrich_does_not_duplicate_on_repeated_reference_sku():
    offers = _offers([["1", "Sony A", "SKU-A", "$100.00", "href"]])
    # a duplicate SKU in the reference data must not multiply the offer row
    site_costs = pd.DataFrame(
        {"sku": ["SKU-A", "SKU-A"], "site_cost": [70.0, 99.0],
         "weight_oz": [40.0, 40.0], "sell_below_cost": [False, False]})
    aged = pd.DataFrame({"sku": [], "aged_status": []})

    out = script.enrich_offers(offers, site_costs, aged)
    assert len(out) == 1
