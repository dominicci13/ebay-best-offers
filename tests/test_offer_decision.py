"""Branch-coverage tests for the `_decide_offer` pricing rule.

`_decide_offer` is a pure function from offer inputs to
`(action, counteroffer, log_pct)`. These tests exercise every branch of
the decision tree on inputs that exhibit the relevant condition.
"""
from __future__ import annotations

import pytest


# Standard defaults; individual cases override what they need.
COMMISSION = 0.091
FLOOR_FACTOR = 0.9       # MAX_DISCOUNT — 90% of list  → floor price
CEILING_FACTOR = 0.95    # MIN_DISCOUNT — 95% of list  → ceiling price


@pytest.mark.parametrize(
    "case_id, cx, site, list_price, total, brand, blocked, want_action, want_counter",
    [
        # 1) Removed By Buyer — any of four conditions trigger this branch.
        ("removed_cx_zero",          0,   100,  200, 150, "Sony",   [],          "Removed By Buyer", 0.0),
        ("removed_site_zero",      150,   0,    200, 150, "Sony",   [],          "Removed By Buyer", 0.0),
        ("removed_site_sentinel",  150,   0.01, 200, 150, "Sony",   [],          "Removed By Buyer", 0.0),
        ("removed_brand_blocked",  150,   100,  200, 150, "Vortex", ["Vortex"],  "Removed By Buyer", 0.0),

        # 2) Accept — customer's offer alone clears the 11% threshold.
        ("accept_high_margin",     200,   100,  250, 100, "Sony",   [],          "Accepted",         0.0),

        # 4) Counter at the ceiling — ceiling (95%) clears, customer's offer doesn't.
        ("counter_at_ceiling",     180,   100,  250, 162, "Sony",   [],          "Counteroffer",     round(250 * CEILING_FACTOR, 2)),

        # 5) Fall-through — nothing clears, counter at list minus a penny.
        ("counter_list_minus_001", 180,   100,  250, 200, "Sony",   [],          "Counteroffer",     round(250 - 0.01, 2)),
    ],
)
def test_decide_offer_branches(
    decide_offer, case_id, cx, site, list_price, total, brand, blocked, want_action, want_counter,
):
    action, counter, _ = decide_offer(
        cx_offer=cx,
        site_cost=site,
        current_price=list_price,
        total_cost=total,
        brand=brand,
        blocked_brands=blocked,
        commission=COMMISSION,
        floor_factor=FLOOR_FACTOR,
        ceiling_factor=CEILING_FACTOR,
    )
    assert action == want_action, f"[{case_id}] action mismatch"
    assert counter == want_counter, f"[{case_id}] counteroffer mismatch"


def test_accepted_returns_cx_profit_pct(decide_offer):
    """When accepting, log_pct should equal the customer-offer profit %."""
    action, _, log_pct = decide_offer(
        cx_offer=200, site_cost=100, current_price=250, total_cost=100,
        brand="Sony", blocked_brands=[],
        commission=COMMISSION, floor_factor=FLOOR_FACTOR, ceiling_factor=CEILING_FACTOR,
    )
    assert action == "Accepted"
    # profit_pct = 1 - total/cx - commission = 1 - 100/200 - 0.091 = 0.409
    assert log_pct == pytest.approx(0.409, abs=1e-9)


def test_blocked_brand_short_circuits_even_with_great_offer(decide_offer):
    """Brand block should win over a high-margin customer offer."""
    action, counter, _ = decide_offer(
        cx_offer=300, site_cost=100, current_price=250, total_cost=50,
        brand="Vortex", blocked_brands=["Vortex"],
        commission=COMMISSION, floor_factor=FLOOR_FACTOR, ceiling_factor=CEILING_FACTOR,
    )
    assert action == "Removed By Buyer"
    assert counter == 0.0
