"""Shared pytest setup for the ebay-best-offers test suite.

``run_ebay_best_offers.py`` is import-safe (its prompt and scheduler run only
under ``if __name__ == "__main__"``), so tests import it directly once the repo
root is on ``sys.path``.
"""
from __future__ import annotations

import pathlib
import sys

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# The real priced-differently accounts come from gitignored config, so tests would
# otherwise pass or fail depending on the machine. Pin fake ones for every test: it keeps
# the suite deterministic and keeps real account names out of this public repo.
FLAT_ACCOUNT = "AccountFlat"
FLAT_LABEL = "accountflat minimum profit margin"

# Mirrors production, where the discount-cap account also carries its own flat floor —
# that floor is the backstop the cap rule falls back to.
CAP_ACCOUNT = "AccountCap"
CAP_FLOOR_LABEL = "accountcap minimum profit margin"
CAP_LABEL = "accountcap maximum discount"


@pytest.fixture(autouse=True)
def _pin_account_pricing_rules(monkeypatch):
    import run_ebay_best_offers as script

    monkeypatch.setattr(script, "FLAT_MIN_PROFIT_ACCOUNTS",
                        {FLAT_ACCOUNT: FLAT_LABEL, CAP_ACCOUNT: CAP_FLOOR_LABEL})
    monkeypatch.setattr(script, "DISCOUNT_CAP_ACCOUNTS", {CAP_ACCOUNT: CAP_LABEL})
