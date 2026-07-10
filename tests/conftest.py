"""Shared pytest setup for the ebay-best-offers test suite.

``run_ebay_best_offers.py`` is import-safe (its prompt and scheduler run only
under ``if __name__ == "__main__"``), so tests import it directly once the repo
root is on ``sys.path``.
"""
from __future__ import annotations

import pathlib
import sys

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
