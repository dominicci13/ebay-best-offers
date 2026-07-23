"""Tests for build_summary_email — the HTML report built from in-memory results."""
from __future__ import annotations

import pandas as pd

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

# A counter, an accept, a decline; one account name carries an ampersand (must escape).
OFFERS = pd.DataFrame(
    [
        ("2026-07-07", "Acme & Co", "Cold CS-22C", "CS-22C&X", 109.99, "111", False, 69.17, 40.0, "N/A", 90.0, False, "b-111", 1),
        ("2026-07-07", "Acme & Co", "Sony A7", "SNY", 250.00, "222", False, 100.00, 32.0, "Slow", 200.0, False, "b-222", 1),
        ("2026-07-07", "Account Two", "Kodak EKMSD", "EK", 24.29, "333", False, 35.64, 40.0, "N/A", 20.0, False, "b-333", 1),
    ],
    columns=["date", "account", "title", "sku", "current_price", "item_number",
             "out_of_stock", "site_cost", "weight_oz", "aged_status", "cx_offer",
             "sell_below_cost", "best_offer_id", "offer_quantity"],
)
RESULTS = script.build_results(OFFERS, SETTINGS)


def test_has_date_accounts_and_totals():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "07/07/2026" in body            # date reformatted to the fleet's MM/DD/YYYY
    assert "Acme &amp; Co" in body         # account name escaped
    assert "Account Two" in body
    assert "All accounts" in body          # totals row


def test_counteroffer_detail_moved_to_workbook():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "CS-22C" not in body            # countered SKU not listed — detail lives in the workbook
    assert "Counter" in body               # but the summary still has a Counter column


def test_account_names_are_html_escaped():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "Acme &amp; Co" in body         # ampersand escaped
    assert "Acme & Co<" not in body        # raw unescaped ampersand not emitted


def test_footer_reflects_settings_not_hardcoded():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "12.0%" in body                 # commission
    assert "9.0%" in body                  # default minimum profit
    assert "Slow 6.0%" in body             # aged tiers echoed
    assert "5% to 10%" in body             # discount band from min/max
    assert "shipping estimated by item weight" in body


def test_avoids_dash_ai_tells():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "—" not in body            # em dash
    assert "–" not in body            # en dash


def test_greeting_intro_and_closing_present():
    body = script.build_summary_email(RESULTS, SETTINGS, greeting_text="Good morning")
    assert "Good morning," in body
    assert "Here is today's summary" in body
    assert "Let me know if you have any questions." in body
    assert "Thanks," in body
