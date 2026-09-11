"""Branch tests for the pricing decision.

Covers ``est_shipping``, ``margin``, ``round_up_to_cent``, ``effective_min_profit``,
``decide_offer`` and the discount-cap rule ``decide_by_discount_cap``.
"""
from __future__ import annotations

import pytest

import run_ebay_best_offers as script
from conftest import CAP_ACCOUNT, FLAT_ACCOUNT

SETTINGS = {
    "commission": 0.12,
    "min_profit": 0.09,
    "slow_min_profit": 0.06,
    "dead_min_profit": 0.04,
    "sell_below_cost_min_profit": 0.04,
    "flat_min_profit::AccountFlat": 0.02,
    "flat_min_profit::AccountCap": 0.02,
    "discount_cap::AccountCap": 0.05,
    "min_discount": 0.05,
    "max_discount": 0.10,
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


# =============================================================================
# Discount-cap accounts: accept within the cap, otherwise counter at exactly the
# cap — with the profit floor still applying underneath.
# =============================================================================
# Every case below is a $100 listing at a 5% cap, so the cap price is $95, and
# weight 40 oz -> $10 shipping. The flat floor for CAP_ACCOUNT is 2%, commission 12%,
# so break-even is total_cost / (1 - 0.12 - 0.02).


def _cap_settings(**overrides):
    settings = dict(SETTINGS)
    settings.update(overrides)
    return settings


@pytest.mark.parametrize(
    "cx_offer, site_cost, expected_action, expected_counter",
    [
        # The two worked examples the rule was specified with (total cost $60).
        (96.00, 50.0, "Accepted", 0.0),        # within 5% of $100 -> accept as offered
        (90.00, 50.0, "Counteroffer", 95.00),  # below the cap -> counter at exactly 5% off
        (95.00, 50.0, "Accepted", 0.0),        # exactly at the cap price is still within it
        (99.99, 50.0, "Accepted", 0.0),
        (1.00, 50.0, "Counteroffer", 95.00),   # however low the offer, the counter is the cap
        # Near-cost item: 5% off no longer clears the 2% floor, so the floor wins and the
        # counter comes back shallower than the cap instead of selling at a loss.
        (90.00, 72.0, "Counteroffer", 95.35),  # total cost $82 -> break-even $95.3488 -> $95.35
        # The same backstop on the accept side: inside the cap, but under the floor.
        # Break-even is $97.6744, rounded UP so the counter cannot land under the floor.
        (96.00, 74.0, "Counteroffer", 97.68),  # total cost $84
        # The floor is unreachable even at full list price -> nothing to counter with.
        (96.00, 90.0, "Declined", 0.0),
    ],
)
def test_discount_cap_account_prices_off_the_selling_price(
    cx_offer, site_cost, expected_action, expected_counter
):
    action, counter, _ = script.decide_offer(
        cx_offer, 100, site_cost, 40, "N/A", False, SETTINGS, False, CAP_ACCOUNT
    )
    assert action == expected_action
    assert counter == expected_counter


def test_discount_cap_accept_never_breaches_the_profit_floor():
    # The offer is inside the cap but below the floor, so it must NOT be accepted, and the
    # counter it gets instead must clear the floor.
    floor = SETTINGS[script.flat_floor_key(CAP_ACCOUNT)]
    action, counter, _ = script.decide_offer(
        96, 100, 74.0, 40, "N/A", False, SETTINGS, False, CAP_ACCOUNT
    )
    assert action == "Counteroffer"
    assert counter > 96
    # Assert on the UNROUNDED margin at the counter price: the returned pct is rounded to
    # 4 places, which would hide a counter sitting a fraction of a cent under break-even.
    assert script.margin(counter, 74.0 + 10.0, SETTINGS["commission"]) >= floor


@pytest.mark.parametrize("site_cost", [50.0, 72.0, 74.0, 82.0, 85.0])
def test_discount_cap_counter_never_rounds_below_break_even(site_cost):
    # round() rounds to NEAREST, so a break-even counter could land a half-cent under the
    # floor; the cap path rounds UP instead. Checked on the unrounded margin.
    floor = SETTINGS[script.flat_floor_key(CAP_ACCOUNT)]
    action, counter, _ = script.decide_offer(
        1.00, 100, site_cost, 40, "N/A", False, SETTINGS, False, CAP_ACCOUNT
    )
    if action == "Counteroffer":
        assert script.margin(counter, site_cost + 10.0, SETTINGS["commission"]) >= floor


def test_round_up_to_cent():
    assert script.round_up_to_cent(97.674418604) == 97.68
    assert script.round_up_to_cent(95.348837209) == 95.35
    assert script.round_up_to_cent(95.0) == 95.00      # already whole cents — unchanged
    assert script.round_up_to_cent(95.35) == 95.35     # and not nudged up by float noise


@pytest.mark.parametrize("cap, site_cost", [(0.00, 50.0), (0.001, 50.0), (0.05, 50.0),
                                            (0.05, 84.5), (0.20, 50.0)])
def test_discount_cap_counter_is_always_strictly_below_list(cap, site_cost):
    # eBay rejects a counteroffer that is not below the Buy It Now price, so a counter at
    # the list price would be recorded as sent and never actually happen. A 0% cap makes
    # the cap price equal the list price, which is the way in.
    settings = _cap_settings(**{script.discount_cap_key(CAP_ACCOUNT): cap})
    action, counter, _ = script.decide_offer(
        50, 100, site_cost, 40, "N/A", False, settings, False, CAP_ACCOUNT
    )
    if action == "Counteroffer":
        assert counter < 100
    else:
        assert (action, counter) == ("Declined", 0.0)


def test_discount_cap_ignores_the_counteroffer_discount_band():
    # The band is 5-10%, but a 20% cap must counter at 20% off — the band is not consulted
    # on a discount-cap account.
    settings = _cap_settings(**{script.discount_cap_key(CAP_ACCOUNT): 0.20})
    action, counter, _ = script.decide_offer(
        50, 100, 50, 40, "N/A", False, settings, False, CAP_ACCOUNT
    )
    assert action == "Counteroffer"
    assert counter == 80.00


def test_discount_cap_does_not_change_other_accounts():
    # Same numbers, an account with no cap: the margin-first rule still decides, so the
    # $90 offer that a cap account counters at $95 is simply accepted here.
    action, counter, _ = script.decide_offer(
        90, 100, 50, 40, "N/A", False, SETTINGS, False, "SomeOtherAccount"
    )
    assert (action, counter) == ("Accepted", 0.0)


@pytest.mark.parametrize(
    "kwargs, expected",
    [
        ({"cx_offer": 0}, "Expired Offer"),
        ({"out_of_stock": True}, "Out of Stock"),
        ({"site_cost": 0}, "Missing Site Cost"),
        ({"offer_code": "SellerCounterOffer"}, "Awaiting Buyer"),
    ],
)
def test_discount_cap_still_honours_the_skip_guards(kwargs, expected):
    # The cap rule runs after the guards, so an expired / out-of-stock / costless /
    # our-own-counteroffer row is skipped on a cap account exactly as anywhere else.
    args = {"cx_offer": 90, "current_price": 100, "site_cost": 50, "weight_oz": 40,
            "aged_status": "N/A", "sell_below_cost": False, "settings": SETTINGS,
            "out_of_stock": False, "account": CAP_ACCOUNT, "offer_code": ""}
    args.update(kwargs)
    assert script.decide_offer(**args)[0] == expected


def test_margin_path_never_counters_at_the_listing_price():
    # min_discount 0% makes the shallowest allowed discount the list price itself. eBay
    # rejects a counter that isn't below Buy It Now, and the row would still have been
    # recorded as a real Counteroffer, so this must decline instead.
    settings = dict(SETTINGS, min_discount=0.00)
    # site_cost 69 + $10 shipping -> total cost 79; break-even at the 9% floor is exactly $100.
    action, counter, _ = script.decide_offer(
        50, 100, 69.0, 40, "N/A", False, settings, False, "SomeOtherAccount"
    )
    assert (action, counter) == ("Declined", 0.0)


def test_margin_path_still_counters_at_the_shallowest_allowed_discount():
    # The guard above must not swallow the ordinary near-boundary case: break-even landing
    # on the shallowest allowed discount ($95 at the 5% min) is a valid counter.
    # site_cost 65.05 + $10 shipping -> total cost 75.05; break-even is 94.99999999999999,
    # a float-hair UNDER the $95 band edge (so this case does not pin the inclusive
    # `target <= highest_price` comparison), and rounds to $95.00.
    action, counter, pct = script.decide_offer(
        50, 100, 65.05, 40, "N/A", False, SETTINGS, False, "SomeOtherAccount"
    )
    assert (action, counter) == ("Counteroffer", 95.00)
    assert pct >= SETTINGS["min_profit"]
