"""Tests for the control-workbook settings check.

Targets the pure ``check_settings`` function, so no Excel is needed — the pandas
read in ``read_settings`` is a thin wrapper exercised by a real run.
"""
from __future__ import annotations

import pytest

import run_ebay_best_offers as script


def _good(**overrides):
    """A complete, valid label -> value mapping; override individual cells."""
    entered = {
        "ebay commission": 0.12,
        "minimum profit margin": 0.09,
        "slow item minimum profit margin": 0.06,
        "dead item minimum profit margin": 0.04,
        "sell below cost minimum profit margin": 0.04,
        "min counteroffer discount": 0.05,
        "max counteroffer discount": 0.10,
        "minimum estimated shipping": 12,
    }
    entered.update(overrides)
    return entered


def test_valid_settings_have_no_problems():
    settings, problems = script.check_settings(_good())
    assert problems == []
    assert settings == {
        "commission": 0.12,
        "min_profit": 0.09,
        "slow_min_profit": 0.06,
        "dead_min_profit": 0.04,
        "sell_below_cost_min_profit": 0.04,
        "min_discount": 0.05,
        "max_discount": 0.10,
        "shipping_floor": 12.0,
    }


@pytest.mark.parametrize(
    "missing",
    [
        "ebay commission",
        "minimum profit margin",
        "slow item minimum profit margin",
        "dead item minimum profit margin",
        "sell below cost minimum profit margin",
        "min counteroffer discount",
        "max counteroffer discount",
        "minimum estimated shipping",
    ],
)
def test_missing_label_is_reported(missing):
    entered = _good()
    del entered[missing]
    _, problems = script.check_settings(entered)
    assert any(missing.split()[0] in p.lower() for p in problems)


def test_blank_value_is_reported():
    _, problems = script.check_settings(_good(**{"ebay commission": ""}))
    assert problems


def test_non_numeric_is_reported():
    _, problems = script.check_settings(_good(**{"ebay commission": "twelve"}))
    assert any("commission" in p.lower() for p in problems)


def test_percent_out_of_range_uses_percent_hint():
    # 12 (== 1200%) is the classic "typed 12 instead of 12%" slip.
    _, problems = script.check_settings(_good(**{"ebay commission": 12}))
    assert any("commission" in p.lower() and "percentage" in p.lower() for p in problems)


def test_shipping_floor_out_of_range_uses_dollar_hint():
    _, problems = script.check_settings(_good(**{"minimum estimated shipping": 200}))
    assert any("shipping" in p.lower() and "dollar" in p.lower() for p in problems)


def test_shipping_floor_accepts_dollar_string():
    settings, problems = script.check_settings(_good(**{"minimum estimated shipping": "$12"}))
    assert problems == []
    assert settings["shipping_floor"] == 12.0


def test_inverted_discount_band_is_reported():
    entered = _good(**{"min counteroffer discount": 0.10, "max counteroffer discount": 0.05})
    _, problems = script.check_settings(entered)
    assert any("max" in p.lower() and "discount" in p.lower() for p in problems)


def test_every_problem_is_collected():
    # commission missing + margin non-numeric -> at least two problems.
    entered = _good(**{"minimum profit margin": "abc"})
    del entered["ebay commission"]
    _, problems = script.check_settings(entered)
    assert len(problems) >= 2
