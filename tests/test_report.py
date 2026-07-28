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
    "flat_min_profit::AccountFlat": 0.02,
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


# --- Partial-run banner ------------------------------------------------------
# An account can fail on its own now without stopping the run, so the email must
# say so. A partial report that looks complete is worse than no report.

def test_no_banner_when_every_account_succeeded():
    body = script.build_summary_email(RESULTS, SETTINGS)
    assert "did not run" not in body


def test_banner_names_each_failed_account():
    body = script.build_summary_email(RESULTS, SETTINGS, failed_accounts=["Acct2", "Acct3"])
    assert "2 accounts did not run" in body
    assert "Acct2" in body
    assert "Acct3" in body
    assert "were not read or answered" in body


def test_banner_singular_for_one_failed_account():
    body = script.build_summary_email(RESULTS, SETTINGS, failed_accounts=["Acct2"])
    assert "1 account did not run" in body


# --- Expired dropped from the table -----------------------------------------
# The 7/23 GetBestOffers pagination fix removed the cause of false "Expired"
# rows (none have landed since), so the column was dead weight in the email.
# A real one can still reach SQL, so the Total must not count what it can't show.

WITH_EXPIRED = pd.DataFrame(
    [
        ("2026-07-27", "Solo", "Sony A7", "SNY", 250.00, "222", False, 100.00, 32.0, "N/A", 200.0, False, "b-222", 1),
        ("2026-07-27", "Solo", "Kodak EK", "EK", 24.29, "333", False, 35.64, 40.0, "N/A", 20.0, False, "b-333", 1),
        ("2026-07-27", "Solo", "Nikon Z", "NZ", 500.00, "555", False, 200.00, 30.0, "N/A", 450.0, False, "b-555", 1),
        ("2026-07-27", "Solo", "Gone", "GX", 100.00, "444", False, 50.00, 40.0, "N/A", 0.0, False, None, 1),
    ],
    columns=OFFERS.columns,
)


def test_expired_is_not_a_column():
    body = script.build_summary_email(script.build_results(WITH_EXPIRED, SETTINGS), SETTINGS)
    assert "Expired" not in body
    assert "Missing Cost" in body          # the other skip reasons stay


def test_total_counts_only_what_the_table_shows():
    results = script.build_results(WITH_EXPIRED, SETTINGS)
    assert (results["action"] == "Expired Offer").sum() == 1   # fixture really has one

    body = script.build_summary_email(results, SETTINGS)
    totals_row = body.rsplit("<tr>", 1)[1]
    assert ">3<" in totals_row             # the 3 rows the columns actually show
    assert ">4<" not in totals_row         # not inflated by the hidden Expired row


def test_banner_escapes_names_and_avoids_dash_ai_tells():
    body = script.build_summary_email(RESULTS, SETTINGS, failed_accounts=["Acme & Co"])
    assert "Acme &amp; Co" in body
    assert "Acme & Co<" not in body
    assert "—" not in body
    assert "–" not in body
