"""Shared pytest fixtures for the ebay-best-offers test suite.

The main script (`run_ebay_best_offers.py`) executes side-effecting
top-level statements at import time (`ask_user`, `run_on_schedule`), so we
slice individual pure functions out of the source rather than importing the
module. This keeps tests fast and free of Excel/Selenium dependencies.
"""
from __future__ import annotations

import pathlib

import pytest


SCRIPT_PATH = pathlib.Path(__file__).resolve().parent.parent / "run_ebay_best_offers.py"


def _slice_function(src: str, name: str) -> str:
    """Return the source of one top-level function definition."""
    marker = f"def {name}"
    start = src.index(marker)
    # End at the next top-level `def ` after `start`, or EOF.
    nxt = src.find("\ndef ", start + 1)
    return src[start:nxt] if nxt != -1 else src[start:]


@pytest.fixture(scope="session")
def decide_offer():
    """Pure offer-decision function lifted from the main script."""
    src = SCRIPT_PATH.read_text(encoding="utf-8")
    ns: dict = {}
    exec(_slice_function(src, "_decide_offer"), ns)
    return ns["_decide_offer"]
