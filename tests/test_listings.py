"""Tests for reading offered listings from the Trading API and shaping them.

Replaces the old grid tests. The decisions worth pinning are the ones the Seller
Hub grid used to make for us: which listings get read at all, and how out-of-stock
is decided now that eBay's respond-link signal is gone.
"""
from __future__ import annotations

import pytest

import run_ebay_best_offers as script

COLUMNS = ["date", "account", "title", "sku", "current_price", "item_number", "out_of_stock"]


def item(item_number: str = "234567890123", sku: str | None = "ACM-HP-001",
         price: float = 179.99, available: int = 4, title: str = "Acme Studio Headphones") -> dict:
    return {
        "item_number": item_number, "title": title, "sku": sku,
        "current_price": price, "quantity": available + 1, "quantity_sold": 1,
        "quantity_available": available, "listing_status": "Active",
    }


def offer(item_number: str, amount: float, offer_id: str) -> dict:
    return {"item_number": item_number, "cx_offer": amount, "best_offer_id": offer_id,
            "quantity": 1, "code": "BuyerBestOffer"}


# --- shaping -----------------------------------------------------------------

def test_columns_are_shaped_and_ordered():
    frame = script.items_to_frame([item()], "Acct1", "2026-07-02")
    assert list(frame.columns) == COLUMNS
    row = frame.iloc[0]
    assert row["date"] == "2026-07-02"
    assert row["account"] == "Acct1"
    assert row["item_number"] == "234567890123"
    assert row["current_price"] == 179.99


def test_no_items_gives_an_empty_frame_that_still_has_the_columns():
    frame = script.items_to_frame([], "Acct1", "2026-07-02")
    assert frame.empty
    assert list(frame.columns) == COLUMNS


def test_a_listing_with_stock_is_not_out_of_stock():
    assert bool(script.items_to_frame([item(available=3)], "A", "2026-07-02").iloc[0]["out_of_stock"]) is False


def test_a_sold_out_listing_is_out_of_stock():
    # The grid used to tell us this by omitting the respond link; now it comes
    # from available quantity. Real case: qty 7, sold 7, listing still Active.
    assert bool(script.items_to_frame([item(available=0)], "A", "2026-07-02").iloc[0]["out_of_stock"]) is True


def test_a_missing_price_becomes_nan_rather_than_a_string():
    frame = script.items_to_frame([{**item(), "current_price": None}], "A", "2026-07-02")
    assert frame["current_price"].isna().all()


# --- which listings get read -------------------------------------------------

@pytest.fixture
def fake_get_item(monkeypatch):
    """Record the item ids requested and serve a listing for each."""
    requested: list[str] = []

    def get_item(token: str, item_id: str) -> dict:
        requested.append(item_id)
        return item(item_number=item_id)

    monkeypatch.setattr(script.ebay_api, "get_item", get_item)
    return requested


def test_only_listings_with_offers_are_read(fake_get_item):
    offers = [offer("111", 10.0, "o1"), offer("222", 20.0, "o2")]
    script.fetch_offer_items("tok", offers, "Acct", "2026-07-02")
    assert fake_get_item == ["111", "222"]


def test_a_listing_carrying_several_offers_is_read_once(fake_get_item):
    # Five offers on one listing is a shape seen in production. Reading it five
    # times would both waste calls and duplicate it in the merge.
    offers = [offer("111", amount, f"o{n}") for n, amount in enumerate([800, 600, 600, 675, 784])]
    frame = script.fetch_offer_items("tok", offers, "Acct", "2026-07-02")
    assert fake_get_item == ["111"]
    assert len(frame) == 1


def test_no_offers_reads_nothing_and_returns_an_empty_frame(fake_get_item):
    frame = script.fetch_offer_items("tok", [], "Acct", "2026-07-02")
    assert fake_get_item == []
    assert frame.empty
    assert list(frame.columns) == COLUMNS


# --- the merge invariant -----------------------------------------------------

def test_every_offer_becomes_exactly_one_row():
    offers = [offer("111", amount, f"o{n}") for n, amount in enumerate([800, 600, 600, 675, 784])]
    listings = script.items_to_frame([item(item_number="111")], "Acct", "2026-07-02")
    merged = script.attach_api_offers(listings, offers)
    assert len(merged) == 5
    assert sorted(merged["best_offer_id"]) == ["o0", "o1", "o2", "o3", "o4"]


def test_identical_offer_amounts_are_kept_as_separate_offers():
    # Two buyers at 600 is not a duplicate; de-duplicating on amount would drop a
    # real offer and leave a buyer unanswered.
    offers = [offer("111", 600.0, "o1"), offer("111", 600.0, "o2")]
    listings = script.items_to_frame([item(item_number="111")], "Acct", "2026-07-02")
    merged = script.attach_api_offers(listings, offers)
    assert len(merged) == 2
    assert sorted(merged["best_offer_id"]) == ["o1", "o2"]


def test_a_duplicated_listing_raises_instead_of_inventing_offers():
    # The 2026-07-29 duplicate rows came from the grid carrying an item twice,
    # which cross-products in the merge. That must fail loudly, not be reported.
    offers = [offer("111", 10.0, "o1")]
    listings = script.items_to_frame([item(item_number="111"), item(item_number="111")],
                                     "Acct", "2026-07-02")
    with pytest.raises(RuntimeError, match="changed in the merge"):
        script.attach_api_offers(listings, offers)
