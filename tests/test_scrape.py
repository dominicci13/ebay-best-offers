"""Tests for reading the pending-offers grid and shaping it into the offers table.

Covers the pure ``offers_to_frame`` plus the one browser-side decision that is worth
pinning: how ``scrape_pending_offers`` tells "this account has no offers" apart from
"our row selector broke". The driver is faked; the rest of the browser work is
exercised by a real run.
"""
from __future__ import annotations

from types import SimpleNamespace

import pytest

import run_ebay_best_offers as script

COLUMNS = ["date", "account", "title", "sku", "current_price", "item_number", "out_of_stock"]

GRID_ROW = ["234567890123", "Acme Studio Headphones", "ACM-HP-001", "$179.99",
            "https://www.ebay.com/bo/seller/showOffers?itemid=234567890123"]


class FakeDriver:
    """Stand-in for the SeleniumBase driver, with just what the scrape touches."""

    def __init__(self, rows: list, zero_results: bool):
        self.rows = rows
        self.zero_results = zero_results

    def get(self, url: str) -> None:
        pass

    def switch_to_window(self, index: int) -> None:
        pass

    def execute_script(self, script_text: str) -> list:
        return self.rows

    def find_elements(self, by: str, selector: str) -> list:
        return ["zero-results-element"] if self.zero_results else []


@pytest.fixture
def browser_stubs(monkeypatch):
    """Neutralize the column reset, the element wait, and the screenshot."""
    monkeypatch.setattr(script.ebay, "customize_offers_table", lambda driver: None)
    monkeypatch.setattr(script, "WebDriverWait",
                        lambda driver, timeout: SimpleNamespace(until=lambda condition: None))
    monkeypatch.setattr(script, "save_debug_screenshot", lambda *args, **kwargs: "shot.png")


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


def test_scrape_reads_the_rows_the_grid_renders(browser_stubs):
    driver = FakeDriver([GRID_ROW], zero_results=False)

    df = script.scrape_pending_offers(driver, "Acct1", "2026-07-02")

    assert len(df) == 1
    assert df.iloc[0]["item_number"] == "234567890123"


def test_scrape_returns_empty_when_ebay_reports_zero_results(browser_stubs):
    """An account with no pending offers is normal, not a failure.

    eBay renders the table shell even when the filter matches nothing, so the
    empty case reaches the row check rather than the element-wait timeout.
    """
    driver = FakeDriver([], zero_results=True)

    df = script.scrape_pending_offers(driver, "Acct1", "2026-07-02")

    assert list(df.columns) == COLUMNS
    assert len(df) == 0


def test_scrape_raises_when_rows_vanish_with_no_zero_results_message(browser_stubs):
    """No rows AND no zero-results element means the selector really did break."""
    driver = FakeDriver([], zero_results=False)

    with pytest.raises(RuntimeError):
        script.scrape_pending_offers(driver, "Acct1", "2026-07-02")
