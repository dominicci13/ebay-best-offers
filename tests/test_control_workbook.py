"""Tests for the control-workbook settings check.

Targets the pure ``check_settings`` function and the account maps that feed its
required-label list, so no Excel is needed — the pandas read in ``read_settings``
is a thin wrapper exercised by a real run.
"""
from __future__ import annotations

import pytest

import run_ebay_best_offers as script
from conftest import CAP_ACCOUNT, CAP_FLOOR_LABEL, CAP_LABEL, FLAT_ACCOUNT, FLAT_LABEL


def _good(**overrides):
    """A complete, valid label -> value mapping; override individual cells."""
    entered = {
        "ebay commission": 0.12,
        "minimum profit margin": 0.09,
        "slow item minimum profit margin": 0.06,
        "dead item minimum profit margin": 0.04,
        "sell below cost minimum profit margin": 0.04,
        "accountflat minimum profit margin": 0.02,
        "accountcap minimum profit margin": 0.02,
        "accountcap maximum discount": 0.05,
        "min counteroffer discount": 0.05,
        "max counteroffer discount": 0.10,
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
        "flat_min_profit::AccountFlat": 0.02,
        "flat_min_profit::AccountCap": 0.02,
        "discount_cap::AccountCap": 0.05,
        "min_discount": 0.05,
        "max_discount": 0.10,
    }


@pytest.mark.parametrize(
    "missing",
    [
        "ebay commission",
        "minimum profit margin",
        "slow item minimum profit margin",
        "dead item minimum profit margin",
        "sell below cost minimum profit margin",
        "accountflat minimum profit margin",
        "accountcap minimum profit margin",
        "accountcap maximum discount",
        "min counteroffer discount",
        "max counteroffer discount",
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


# --- the account maps that add required labels -------------------------------

def test_duplicate_label_is_reported_not_silently_overwritten(monkeypatch):
    # A cap label copy-pasted from the flat-floor label used to overwrite it in the spec:
    # no problem reported, the run starts, and the first offer raises KeyError deep in
    # effective_min_profit. It must stop the run at the settings check instead.
    monkeypatch.setattr(script, "DISCOUNT_CAP_ACCOUNTS", {CAP_ACCOUNT: CAP_FLOOR_LABEL})
    _, problems = script.check_settings(_good())
    assert any(CAP_FLOOR_LABEL in p and "more than one setting" in p for p in problems)


def test_duplicate_of_a_builtin_label_is_reported(monkeypatch):
    monkeypatch.setattr(script, "DISCOUNT_CAP_ACCOUNTS", {CAP_ACCOUNT: "ebay commission"})
    _, problems = script.check_settings(_good())
    assert any("ebay commission" in p and "more than one setting" in p for p in problems)


@pytest.mark.parametrize("config_key", ["AccountCap", "accountcap", " AccountCap ", "ACCOUNTCAP"])
def test_account_key_resolves_to_the_profile_spelling(monkeypatch, config_key):
    # The pricing path tests membership against the profile name, so a config key that
    # differs only by case or whitespace must still resolve — otherwise the rule silently
    # does not apply while the summary still claims it does.
    monkeypatch.setattr(script, "EBAY_PROFILES", {CAP_ACCOUNT: "Default"})
    assert script._accounts_by_label({config_key: CAP_LABEL}) == {CAP_ACCOUNT: CAP_LABEL}


def test_account_matching_no_profile_is_dropped(monkeypatch):
    monkeypatch.setattr(script, "EBAY_PROFILES", {CAP_ACCOUNT: "Default"})
    assert script._accounts_by_label({"NotAProfile": CAP_LABEL}) == {}


def test_summaries_name_only_accounts_whose_setting_resolved(monkeypatch):
    # An account that never resolved must not appear in the log line or the email footer
    # claiming its rule was applied.
    monkeypatch.setattr(script, "EBAY_PROFILES", {CAP_ACCOUNT: "Default", FLAT_ACCOUNT: "Profile 1"})
    monkeypatch.setattr(script, "DISCOUNT_CAP_ACCOUNTS",
                        script._accounts_by_label({CAP_ACCOUNT: CAP_LABEL, "Ghost": "ghost maximum discount"}))
    monkeypatch.setattr(script, "FLAT_MIN_PROFIT_ACCOUNTS",
                        script._accounts_by_label({FLAT_ACCOUNT: FLAT_LABEL}))
    settings, problems = script.check_settings(_good())
    assert problems == []
    assert "Ghost" not in script.discount_cap_summary(settings)
    assert CAP_ACCOUNT in script.discount_cap_summary(settings)
    assert FLAT_ACCOUNT in script.flat_floor_summary(settings)
