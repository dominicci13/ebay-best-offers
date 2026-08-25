"""Branch tests for build_results — the full archive record per offer.

Covers every decision outcome and checks that not-applicable fields come back as
NULL (NaN in the frame, converted to SQL NULL at insert time) rather than 0.
"""
from __future__ import annotations

import pandas as pd

import run_ebay_best_offers as script
from conftest import FLAT_ACCOUNT

SETTINGS = {
    "commission": 0.12,
    "min_profit": 0.09,
    "slow_min_profit": 0.06,
    "dead_min_profit": 0.04,
    "sell_below_cost_min_profit": 0.04,
    "flat_min_profit::AccountFlat": 0.02,
    "min_discount": 0.05,
    "max_discount": 0.10,
}

# One row per outcome. Columns match an enriched offer with the API offer attached.
OFFERS = pd.DataFrame(
    [
        # date, account, title, sku, current_price, item_number, out_of_stock, site_cost, weight_oz, aged_status, cx_offer, sell_below_cost, best_offer_id, offer_quantity, offer_code
        ("2026-07-02", "Acct1", "Cold CS-22C",  "CS-22C", 109.99, "111", False, 69.17, 40.0, "N/A", 90.0, False, "b-111", 1, "BuyerBestOffer"),    # counter
        ("2026-07-02", "Acct1", "Sony A7",      "SNY",    250.00, "222", False, 100.00, 32.0, "Slow", 200.0, False, "b-222", 1, "BuyerBestOffer"), # accept
        ("2026-07-02", "Acct1", "Kodak EKMSD",  "EKMSD",  24.29,  "333", False, 35.64, 40.0, "N/A", 20.0, False, "b-333", 1, "BuyerBestOffer"),    # decline
        ("2026-07-02", "Acct1", "Sony Expired", "SZ",     100.00, "444", True,  50.00, 40.0, "N/A", 0.0, False, None, 1, ""),                      # expired
        ("2026-07-02", "Acct1", "Sony NoCost",  "SQ",     100.00, "555", False, 0.00,  40.0, "N/A", 80.0, False, "b-555", 1, "BuyerBestOffer"),    # missing cost
        ("2026-07-02", "Acct1", "Dead Item",    "DED",    3849.0, "666", False, 2800.0, 640.0, "Dead", 3500.0, False, "b-666", 1, "BuyerBestOffer"), # dead accepts
        ("2026-07-02", "Acct1", "OOS Item",     "OOS",    200.00, "777", True,  120.00, 40.0, "N/A", 150.0, False, "b-777", 1, "BuyerBestOffer"),  # out of stock
        ("2026-07-02", "Acct1", "Our Counter",  "OWN",    1022.99, "888", False, 733.27, 0.0, "N/A", 920.69, False, "b-888", 1, "SellerCounterOffer"), # awaiting buyer
    ],
    columns=["date", "account", "title", "sku", "current_price", "item_number",
             "out_of_stock", "site_cost", "weight_oz", "aged_status", "cx_offer",
             "sell_below_cost", "best_offer_id", "offer_quantity", "offer_code"],
)


def _by_item():
    """Run build_results and index the rows by item_number for easy assertions."""
    result = script.build_results(OFFERS, SETTINGS)
    assert list(result.columns) == script.RESULT_COLUMNS + ["best_offer_id", "offer_quantity"]
    return {row.item_number: row for row in result.itertuples(index=False)}


def test_shape_and_report_date():
    result = script.build_results(OFFERS, SETTINGS)
    assert len(result) == len(OFFERS)
    assert (result["report_date"] == "2026-07-02").all()


def test_counter_row_has_price_discount_and_margin():
    row = _by_item()["111"]
    assert row.action == "Counteroffer"
    assert 109.99 * 0.90 <= row.counter <= 109.99 * 0.95   # inside the discount band
    assert row.discount == round((109.99 - row.counter) / 109.99, 4)
    assert row.counter_margin >= SETTINGS["min_profit"]
    assert row.buyer_margin < SETTINGS["min_profit"]       # the buyer's own low offer wouldn't clear


def test_accepted_row_has_buyer_margin_only():
    row = _by_item()["222"]
    assert row.action == "Accepted"
    assert row.buyer_margin >= SETTINGS["min_profit"]
    assert pd.isna(row.counter)
    assert pd.isna(row.discount)
    assert pd.isna(row.counter_margin)


def test_declined_row_keeps_buyer_margin_no_counter():
    row = _by_item()["333"]
    assert row.action == "Declined"
    assert row.buyer_margin < SETTINGS["min_profit"]
    assert pd.isna(row.counter)
    assert pd.isna(row.counter_margin)


def test_expired_row_has_no_margins():
    row = _by_item()["444"]
    assert row.action == "Expired Offer"
    assert row.out_of_stock is True
    assert pd.isna(row.buyer_margin)
    assert pd.isna(row.counter)


def test_missing_cost_row_has_no_buyer_margin():
    row = _by_item()["555"]
    assert row.action == "Missing Site Cost"
    assert pd.isna(row.buyer_margin)  # can't price without a real cost
    assert pd.isna(row.counter)


def test_out_of_stock_row_is_skipped_no_counter():
    row = _by_item()["777"]
    assert row.action == "Out of Stock"
    assert pd.isna(row.counter)
    assert pd.isna(row.counter_margin)


def test_our_own_counteroffer_is_recorded_but_never_answered():
    # Priced as a buyer offer this clears the floor and would read "Accepted" — the row
    # eBay rejected with 21940 on 7/29 and 7/30. It is archived, never sent.
    row = _by_item()["888"]
    assert row.action == "Awaiting Buyer"
    assert pd.isna(row.counter)
    assert pd.isna(row.counter_margin)
    assert row.action not in script.RESPOND_ACTIONS   # nothing goes to eBay


def test_dead_item_accepts_below_default_margin():
    row = _by_item()["666"]
    assert row.action == "Accepted"
    # its margin clears the 4% Dead floor but not the 9% default
    assert SETTINGS["dead_min_profit"] <= row.buyer_margin < SETTINGS["min_profit"]
    assert pd.isna(row.counter)


def test_computed_costs_match_helpers():
    row = _by_item()["111"]
    assert row.est_shipping == script.est_shipping(40.0)             # keyed on weight (40 oz)
    assert row.total_cost == round(69.17 + script.est_shipping(40.0), 2)


# --- attach_api_offers (one row per offer) -----------------------------------

_GRID_COLS = ["date", "account", "title", "sku", "current_price", "item_number",
              "out_of_stock", "site_cost", "weight_oz", "aged_status", "sell_below_cost"]


def test_attach_api_offers_expands_multiple_offers_per_item():
    grid = pd.DataFrame(
        [("2026-07-02", "Acct1", "Multi", "SKU-M", 100.00, "900", False, 50.0, 40.0, "N/A", False)],
        columns=_GRID_COLS,
    )
    api_offers = [
        {"item_number": "900", "cx_offer": 80.0, "best_offer_id": "A", "quantity": 1},
        {"item_number": "900", "cx_offer": 60.0, "best_offer_id": "B", "quantity": 2},
    ]
    out = script.attach_api_offers(grid, api_offers)
    assert len(out) == 2                                  # one row per offer, not one per item
    assert set(out["best_offer_id"]) == {"A", "B"}
    assert sorted(out["cx_offer"]) == [60.0, 80.0]
    assert set(out["offer_quantity"]) == {1, 2}


def test_attach_api_offers_carries_the_offer_code():
    """Drop the code and we can't tell our own counteroffer from a buyer's offer."""
    grid = pd.DataFrame(
        [("2026-07-02", "Acct1", "Mixed", "SKU-X", 100.00, "902", False, 50.0, 40.0, "N/A", False)],
        columns=_GRID_COLS,
    )
    api_offers = [
        {"item_number": "902", "cx_offer": 60.0, "best_offer_id": "A", "quantity": 1,
         "code": "BuyerBestOffer"},
        {"item_number": "902", "cx_offer": 90.0, "best_offer_id": "B", "quantity": 1,
         "code": "SellerCounterOffer"},
    ]
    out = script.attach_api_offers(grid, api_offers)
    assert dict(zip(out["best_offer_id"], out["offer_code"])) == {
        "A": "BuyerBestOffer", "B": "SellerCounterOffer"
    }
    # ...and the code must reach the decision, or the whole guard is a no-op.
    rows = {r.best_offer_id: r for r in script.build_results(out, SETTINGS).itertuples(index=False)}
    assert rows["A"].action != "Awaiting Buyer"
    assert rows["B"].action == "Awaiting Buyer"


def test_attach_api_offers_no_offer_has_an_empty_code():
    grid = pd.DataFrame(
        [("2026-07-02", "Acct1", "None", "SKU-N", 100.00, "903", False, 50.0, 40.0, "N/A", False)],
        columns=_GRID_COLS,
    )
    out = script.attach_api_offers(grid, [])
    assert out.iloc[0]["offer_code"] == ""   # no offer, no code — must not read as Awaiting Buyer


def test_attach_api_offers_no_offer_reads_as_expired():
    grid = pd.DataFrame(
        [("2026-07-02", "Acct1", "None", "SKU-N", 100.00, "901", False, 50.0, 40.0, "N/A", False)],
        columns=_GRID_COLS,
    )
    out = script.attach_api_offers(grid, [])
    assert len(out) == 1
    assert out.iloc[0]["cx_offer"] == 0.0
    assert out.iloc[0]["best_offer_id"] is None
    result = script.build_results(out, SETTINGS)
    assert result.iloc[0]["action"] == "Expired Offer"


def test_account_reaches_the_decision_so_a_flat_floor_account_gets_its_floor():
    """The account must travel from the frame into decide_offer, or a flat-floor
    account silently keeps the 9% default and the whole rule is a no-op."""
    same_offer = [
        ("2026-07-02", FLAT_ACCOUNT, "Flat", "F1", 100.00, "801", False,
         50.00, 40.0, "Dead", 70.0, False, "b-801", 1, "BuyerBestOffer"),
        ("2026-07-02", "SomeOtherAccount", "Flat", "F1", 100.00, "802", False,
         50.00, 40.0, "Dead", 70.0, False, "b-802", 1, "BuyerBestOffer"),
    ]
    frame = pd.DataFrame(same_offer, columns=OFFERS.columns)
    rows = {r.item_number: r for r in script.build_results(frame, SETTINGS).itertuples(index=False)}
    # ~2.9% margin: clears the flat 2%, misses the 4% Dead floor everywhere else.
    assert rows["801"].action == "Accepted"
    assert rows["802"].action != "Accepted"
