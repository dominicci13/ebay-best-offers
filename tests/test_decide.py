"""Branch tests for the pricing decision (est_shipping, effective_min_profit, decide_offer)."""
from __future__ import annotations

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


def test_flat_floor_account_uses_one_floor_for_every_item():
    emp = script.effective_min_profit
    # "all items, not age dependent" — the aged / below-cost easings never apply here.
    assert emp("N/A", False, SETTINGS, FLAT_ACCOUNT) == 0.02
    assert emp("Slow", False, SETTINGS, FLAT_ACCOUNT) == 0.02
    assert emp("Dead", False, SETTINGS, FLAT_ACCOUNT) == 0.02
    assert emp("N/A", True, SETTINGS, FLAT_ACCOUNT) == 0.02
    assert emp("Dead", True, SETTINGS, FLAT_ACCOUNT) == 0.02


def test_other_accounts_keep_the_aged_tiers():
    emp = script.effective_min_profit
    for account in ("SomeOtherAccount", ""):
        assert emp("N/A", False, SETTINGS, account) == 0.09
        assert emp("Dead", False, SETTINGS, account) == 0.04


def test_flat_floor_account_accepts_where_other_accounts_counter():
    # weight 40 oz -> $10 shipping, site_cost 50 -> total_cost 60; offer 70 yields
    # ~2.9% margin: clears the 2% flat floor but none of the other accounts' floors.
    flat = script.decide_offer(70, 100, 50, 40, "N/A", False, SETTINGS, False, FLAT_ACCOUNT)
    other = script.decide_offer(70, 100, 50, 40, "N/A", False, SETTINGS, False, "SomeOtherAccount")
    assert flat[0] == "Accepted"
    assert flat[2] >= SETTINGS[script.flat_floor_key(FLAT_ACCOUNT)]
    assert other[0] != "Accepted"


def test_flat_floor_account_still_respects_its_floor_and_the_discount_band():
    # An offer below the 2% floor is countered, never accepted, and the counter still
    # clears 2% and stays inside the 5-10% discount band.
    action, counter, pct = script.decide_offer(
        50, 100, 50, 40, "N/A", False, SETTINGS, False, FLAT_ACCOUNT
    )
    assert action == "Counteroffer"
    assert pct >= SETTINGS[script.flat_floor_key(FLAT_ACCOUNT)]
    assert 100 * 0.90 <= counter <= 100 * 0.95


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


def test_our_own_outstanding_counteroffer_is_never_answered():
    # GetBestOffers returns our own live counteroffer alongside the buyer's offers. Its
    # price was built to clear the floor, so pricing it as a buyer offer always says
    # "Accepted" — and eBay rejects that with 21940 (you can't accept your own offer).
    action, counter, _ = script.decide_offer(
        90, 100, 50, 40, "N/A", False, SETTINGS, False, "", "SellerCounterOffer"
    )
    assert action == "Awaiting Buyer"
    assert counter == 0.0


def test_the_7_30_deals_drop_rejection_reproduces_and_is_fixed():
    # The exact row behind 21940 on 7/29 and 7/30: our $920.69 counter (10% off $1,022.99)
    # read back as an offer, 6.73% margin against a 2% flat floor -> "Accepted".
    args = (920.69, 1022.99, 733.27, 0, "N/A", False, SETTINGS, False, FLAT_ACCOUNT)
    assert script.decide_offer(*args)[0] == "Accepted"          # today's wrong answer
    assert script.decide_offer(*args, "SellerCounterOffer")[0] == "Awaiting Buyer"


def test_buyer_codes_are_answered_normally():
    for code in ("BuyerBestOffer", "BuyerCounterOffer"):
        action, _, _ = script.decide_offer(200, 250, 100, 32, "N/A", False, SETTINGS, False, "", code)
        assert action == "Accepted", code


def test_a_missing_code_still_answers_so_real_offers_are_never_dropped():
    # No code means no evidence the offer is ours; skipping it would cost a sale.
    action, _, _ = script.decide_offer(200, 250, 100, 32, "N/A", False, SETTINGS, False, "", "")
    assert action == "Accepted"


def test_expired_wins_over_awaiting_buyer():
    # A no-offer row carries no code; it must stay Expired, not become Awaiting Buyer.
    action, _, _ = script.decide_offer(0, 100, 50, 100, "N/A", False, SETTINGS, False, "", "")
    assert action == "Expired Offer"


def test_awaiting_buyer_wins_over_out_of_stock_and_missing_cost():
    # It isn't our offer to answer at all — that outranks the other skip reasons.
    for oos, cost in ((True, 50), (False, 0)):
        action, _, _ = script.decide_offer(
            90, 100, cost, 40, "N/A", False, SETTINGS, oos, "", "SellerCounterOffer"
        )
        assert action == "Awaiting Buyer"


def test_dead_and_sell_below_cost_accept_where_normal_would_not():
    # weight 40 oz -> $10 shipping, site_cost 50 -> total_cost 60; offer 74 yields
    # ~6.9% margin: clears the 4% Dead / below-cost floor but not the 9% default
    normal = script.decide_offer(74, 100, 50, 40, "N/A", False, SETTINGS)
    dead = script.decide_offer(74, 100, 50, 40, "Dead", False, SETTINGS)
    sbc = script.decide_offer(74, 100, 50, 40, "N/A", True, SETTINGS)
    assert normal[0] != "Accepted"     # 6.9% < 9%
    assert dead[0] == "Accepted"       # 6.9% >= 4%
    assert sbc[0] == "Accepted"        # 6.9% >= 4%
