"""Tests for shaping scraped offer rows into the offers table.

Targets the pure ``offers_to_frame`` — the browser side (``scrape_pending_offers``)
is exercised by a real run.
"""
from __future__ import annotations

import run_ebay_best_offers as script

COLUMNS = ["date", "account", "title", "sku", "current_price", "item_number", "out_of_stock"]


def test_offers_to_frame_shapes_and_orders_columns():
    rows = [
        ["234567890123", "Acme Studio Headphones", "ACM-HP-001", "$179.99",
         "https://www.ebay.com/bo/seller/showOffers?itemid=234567890123"],
        ["345678901234", "Acme Field Binoculars 10x42", "36014", "$3,849.00",
         "https://www.ebay.com/bo/seller/showOffers?itemid=345678901234"],
        # out of stock: no Respond link
        ["456789012345", "Acme Rifle Scope 4-12x40", "ACM-RS-003", "$129.99", ""],
    ]
    df = script.offers_to_frame(rows, "Acct1", "2026-07-02")

    assert list(df.columns) == COLUMNS
    assert df.iloc[0]["date"] == "2026-07-02"
    assert df.iloc[0]["account"] == "Acct1"
    assert df.iloc[0]["item_number"] == "234567890123"
    # "$3,849.00" -> 3849.0
    assert df.iloc[1]["current_price"] == 3849.0
    assert bool(df.iloc[0]["out_of_stock"]) is False
    assert bool(df.iloc[2]["out_of_stock"]) is True


def test_offers_to_frame_handles_empty():
    df = script.offers_to_frame([], "Acct1", "2026-07-02")
    assert list(df.columns) == COLUMNS
    assert len(df) == 0
