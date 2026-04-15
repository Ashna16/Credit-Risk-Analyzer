"""Tests for fiscal-year orphan filtering in backend/financial_parser.py."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BACKEND = ROOT / "backend"
if str(BACKEND) not in sys.path:
    sys.path.insert(0, str(BACKEND))

from financial_parser import _filter_orphan_fiscal_year_rows  # noqa: E402


def test_filter_drops_sparse_year_below_main_band():
    rows = [
        {"statement_type": "income_statement", "selected_year": 2025, "revenue": 1.0, "ebit": 2.0},
        {"statement_type": "income_statement", "selected_year": 2024, "revenue": 1.0, "ebit": 2.0},
        {"statement_type": "income_statement", "selected_year": 2023, "revenue": 1.0, "ebit": 2.0},
        {"statement_type": "income_statement", "selected_year": 2021, "net_income": 99.0},
    ]
    out = _filter_orphan_fiscal_year_rows(rows)
    assert {int(r["selected_year"]) for r in out} == {2023, 2024, 2025}


def test_filter_noop_when_single_year():
    rows = [
        {"statement_type": "income_statement", "selected_year": 2024, "revenue": 1.0},
    ]
    out = _filter_orphan_fiscal_year_rows(rows)
    assert len(out) == 1
