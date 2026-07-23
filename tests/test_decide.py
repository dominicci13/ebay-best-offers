"""Branch tests for the pricing decision (est_shipping, effective_min_profit, decide_offer)."""
from __future__ import annotations

import run_ebay_best_offers as script

SETTINGS = {
    "commission": 0.12,
    "min_profit": 0.09,
    "slow_min_profit": 0.06,
    "dead_min_profit": 0.04,
    "sell_below_cost_min_profit": 0.04,
    "min_discount": 0.05,
    "max_discount": 0.10,
    "shipping_floor": 12.0,
}


def test_est_shipping_weight_tiers():
    # shipping is keyed on total weight in ounces (lbs*16 + oz)
    assert script.est_shipping(0) == 15.0      # missing/zero weight -> bad-weight rate
    assert script.est_shipping(1) == 15.0      # <= 1 oz
    assert script.est_shipping(16) == 8.0      # 1 lb
    assert script.est_shipping(24) == 10.0     # 1.5 lb rounds up into the 2-3 lb tier
    assert script.est_shipping(48) == 10.0     # 3 lb
    assert script.est_shipping(320) == 20.0    # 20 lb lands in the 11-20 lb tier
    assert script.est_shipping(560) == 25.0    # 35 lb
    assert script.est_shipping(1600) == 80.0   # 100 lb
    assert script.est_shipping(1700) == 100.0  # over 100 lb


def test_effective_min_profit_tiers_and_lowest_wins():
    emp = script.effective_min_profit
    assert emp("N/A", False, SETTINGS) == 0.09     # default
    assert emp("Slow", False, SETTINGS) == 0.06
    assert emp("Dead", False, SETTINGS) == 0.04
    assert emp("N/A", True, SETTINGS) == 0.04      # SellBelowCost
    assert emp("slow", False, SETTINGS) == 0.06    # case-insensitive
    assert emp("Slow", True, SETTINGS) == 0.04     # lowest applicable wins (6% vs 4%)


def test_expired_offer():
    action, counter, _ = script.decide_offer(0, 100, 50, 100, "N/A", False, SETTINGS)
    assert action == "Expired Offer"
    assert counter == 0.0


def test_missing_site_cost():
    for bad in (0, 0.01):
        action, _, _ = script.decide_offer(80, 100, bad, 100, "N/A", False, SETTINGS)
        assert action == "Missing Site Cost"


def test_out_of_stock_is_skipped_but_expired_wins():
    action, counter, _ = script.decide_offer(90, 100, 50, 100, "N/A", False, SETTINGS, out_of_stock=True)
    assert action == "Out of Stock"           # can't fulfill -> never answered
    assert counter == 0.0
    # No readable offer takes priority: label Expired, not Out of Stock.
    action, _, _ = script.decide_offer(0, 100, 50, 100, "N/A", False, SETTINGS, out_of_stock=True)
    assert action == "Expired Offer"


def test_accept_when_offer_clears_margin():
    action, counter, pct = script.decide_offer(200, 250, 100, 32, "N/A", False, SETTINGS)
    assert action == "Accepted"
    assert counter == 0.0
    assert pct >= SETTINGS["min_profit"]


def test_counter_within_band_that_clears():
    # weight 40 oz -> $10 shipping, total_cost 79.17: counters at the floor within the band
    action, counter, pct = script.decide_offer(90.0, 109.99, 69.17, 40, "N/A", False, SETTINGS)
    assert action == "Counteroffer"
    assert pct >= SETTINGS["min_profit"]
    assert 109.99 * 0.90 <= counter <= 109.99 * 0.95   # inside the discount band


def test_decline_when_even_shallow_discount_cannot_clear():
    action, counter, _ = script.decide_offer(20, 24.29, 35.64, 40, "N/A", False, SETTINGS)
    assert action == "Declined"
    assert counter == 0.0


def test_dead_and_sell_below_cost_accept_where_normal_would_not():
    # weight 40 oz -> $10 shipping, site_cost 50 -> total_cost 60; offer 74 yields
    # ~6.9% margin: clears the 4% Dead / below-cost floor but not the 9% default
    normal = script.decide_offer(74, 100, 50, 40, "N/A", False, SETTINGS)
    dead = script.decide_offer(74, 100, 50, 40, "Dead", False, SETTINGS)
    sbc = script.decide_offer(74, 100, 50, 40, "N/A", True, SETTINGS)
    assert normal[0] != "Accepted"     # 6.9% < 9%
    assert dead[0] == "Accepted"       # 6.9% >= 4%
    assert sbc[0] == "Accepted"        # 6.9% >= 4%
