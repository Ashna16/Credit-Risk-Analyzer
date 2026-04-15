import pandas as pd

from loan_doc_extractor_mvp.app import (
    _to_numeric_financial,
    _build_dynamic_credit_analysis,
    _should_invalidate_multi_bundle,
    _multi_block_reason,
    _should_block_selected_context,
    _lock_analysis_dataset,
)


def _base_tables():
    return {
        "financial_actuals": pd.DataFrame([
            {
                "revenue": 200_000_000,
                "ebitda": 40_000_000,
                "ebit": 24_000_000,
                "interest_expense": 8_000_000,
                "net_income": 10_000_000,
                "total_debt": 120_000_000,
                "cash": 20_000_000,
                "total_assets": 300_000_000,
                "current_assets": 90_000_000,
                "current_liabilities": 45_000_000,
                "operating_cash_flow": 35_000_000,
                "capital_expenditures": 5_000_000,
                "free_cash_flow": 30_000_000,
            }
        ]),
        "covenant_thresholds": pd.DataFrame([
            {
                "max_total_leverage_ratio": 4.0,
                "min_interest_coverage_ratio": 2.5,
                "min_fixed_charge_coverage_ratio": 1.2,
                "min_liquidity_requirement": 1.1,
            }
        ]),
        "deal_terms": pd.DataFrame([
            {
                "loan_amount": 100_000_000,
                "amortization_schedule": 6_000_000,
                "tenor": "5Y",
            }
        ]),
        "collateral_data": pd.DataFrame([
            {
                "collateral_value": 220_000_000,
                "lien_priority": "first",
                "guarantee_presence": True,
            }
        ]),
    }


def test_numeric_unit_parsing():
    assert _to_numeric_financial("8%") == 0.08
    assert _to_numeric_financial("75 bps") == 0.0075
    assert _to_numeric_financial("2.5x") == 2.5
    assert _to_numeric_financial("$10m") == 10_000_000


def test_recalc_changes_with_risk_drivers():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])

    base = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=10_000_000,
    )
    stressed = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Oil & Gas",
        geography="Emerging Market - High Volatility",
        business_stage="Startup",
        company_size="Small",
        years_in_operation=1,
        requested_amount=10_000_000,
    )

    assert stressed["final_score"] != base["final_score"]


def test_requested_amount_only_changes_limit():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])
    # Missing loan_amount should not couple requested amount to risk score.
    tables["deal_terms"] = pd.DataFrame([{"loan_amount": None, "amortization_schedule": 6_000_000, "tenor": "5Y"}])

    low_req = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=5_000_000,
    )
    high_req = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=25_000_000,
    )

    assert low_req["final_score"] == high_req["final_score"]
    assert low_req["grade"] == high_req["grade"]
    assert low_req["approved_limit"] != high_req["approved_limit"]


def test_interest_coverage_uses_ebit():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])
    model = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=10_000_000,
    )
    row = model["table"][model["table"]["Metric"] == "Interest Coverage"].iloc[0]
    # EBIT (24m) / Interest (8m) = 3.0, not EBITDA(40m)/Interest(8m)=5.0
    assert abs(float(row["Calculated Value"]) - 3.0) < 1e-9


def test_approved_limit_is_policy_capped_then_requested_capped():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])
    model_high_req = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=500_000_000,
    )
    model_low_req = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=5_000_000,
    )

    assert model_high_req["approved_limit"] <= model_high_req["policy_approved_limit"]
    assert model_low_req["approved_limit"] == 5_000_000


def test_geography_adjustment_changes_score_with_same_data():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])
    us = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=10_000_000,
    )
    emerging = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="Emerging Market - High Volatility",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=10_000_000,
    )
    assert emerging["final_score"] != us["final_score"]


def test_manual_override_inputs_change_score():
    tables = _base_tables()
    combined_df = pd.DataFrame([{"Revenue": 200_000_000}])
    conservative = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Oil & Gas",
        geography="Sanctioned / High Risk Region",
        business_stage="Startup",
        company_size="Small",
        years_in_operation=1,
        requested_amount=10_000_000,
    )
    favorable = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Healthcare",
        geography="Canada",
        business_stage="Mature",
        company_size="Large",
        years_in_operation=12,
        requested_amount=10_000_000,
    )
    assert conservative["final_score"] != favorable["final_score"]


def test_dataset_invalidation_multi_mode():
    bundle = {"source": "multi", "context_key": "A|B"}
    assert _should_invalidate_multi_bundle(bundle, "A|B", "multi") is False
    assert _should_invalidate_multi_bundle(bundle, "A|C", "multi") is True


def test_mixed_borrower_docs_blocked():
    bundle = {
        "source": "multi",
        "borrower_mismatch": True,
        "unresolved_borrower_docs": [],
        "completeness": {"missing_buckets": []},
    }
    reason = _multi_block_reason(bundle)
    assert reason is not None and "different borrowers" in reason.lower()


def test_unresolved_required_borrower_docs_blocked():
    bundle = {
        "source": "multi",
        "borrower_mismatch": False,
        "unresolved_required_borrower_docs": ["credit_agreement_sample_valid.pdf"],
        "unresolved_borrower_docs": ["credit_agreement_sample_valid.pdf"],
        "completeness": {"missing_buckets": []},
    }
    reason = _multi_block_reason(bundle)
    assert reason is not None
    assert "required docs" in reason.lower()


def test_stale_selected_file_blocking_helper():
    assert _should_block_selected_context(None, None, has_cache=False) is False
    assert _should_block_selected_context(None, "a.pdf", has_cache=False) is False
    # Different file + no cache => block
    from pathlib import Path
    assert _should_block_selected_context(Path("b.pdf"), "a.pdf", has_cache=False) is True
    # Different file + cache available => allow cache swap
    assert _should_block_selected_context(Path("b.pdf"), "a.pdf", has_cache=True) is False


def test_analysis_year_lock_latest_and_specific_year_modes():
    df = pd.DataFrame(
        [
            {"Sheet": "Income Statement", "Selected Year": 2023, "Revenue": 100},
            {"Sheet": "Income Statement", "Selected Year": 2024, "Revenue": 200},
            {"Sheet": "Income Statement", "Selected Year": 2025, "Revenue": 300},
            {"Sheet": "Balance Sheet", "Selected Year": 2023, "Total Assets": 500},
            {"Sheet": "Balance Sheet", "Selected Year": 2024, "Total Assets": 600},
            {"Sheet": "Balance Sheet", "Selected Year": 2025, "Total Assets": 700},
            {"Sheet": "Cash Flow", "Selected Year": 2023, "Operating Cash Flow": 50},
            {"Sheet": "Cash Flow", "Selected Year": 2024, "Operating Cash Flow": 60},
            {"Sheet": "Cash Flow", "Selected Year": 2025, "Operating Cash Flow": 70},
        ]
    )

    latest = _lock_analysis_dataset(df, mode="latest_available", specific_year=None)
    assert latest["locked_year"] == 2025
    assert latest["cross_year_error_flag"] is False
    assert set(latest["locked_df"]["Selected Year"].dropna().astype(int).tolist()) == {2025}

    y2024 = _lock_analysis_dataset(df, mode="specific_year", specific_year=2024)
    assert y2024["locked_year"] == 2024
    assert y2024["cross_year_error_flag"] is False
    assert set(y2024["locked_df"]["Selected Year"].dropna().astype(int).tolist()) == {2024}

    y2023 = _lock_analysis_dataset(df, mode="specific_year", specific_year=2023)
    assert y2023["locked_year"] == 2023
    assert y2023["cross_year_error_flag"] is False
    assert set(y2023["locked_df"]["Selected Year"].dropna().astype(int).tolist()) == {2023}


def test_partial_upload_income_statement_only_marks_incomplete_dependencies():
    tables = _base_tables()
    fa = tables["financial_actuals"].copy()
    # Income statement only
    fa.loc[0, "total_debt"] = None
    fa.loc[0, "cash"] = None
    fa.loc[0, "current_assets"] = None
    fa.loc[0, "current_liabilities"] = None
    fa.loc[0, "total_assets"] = None
    fa.loc[0, "equity"] = None
    fa.loc[0, "operating_cash_flow"] = None
    fa.loc[0, "capital_expenditures"] = None
    fa.loc[0, "free_cash_flow"] = None
    tables["financial_actuals"] = fa
    combined_df = pd.DataFrame([{"Detected Type": "Financial Statements", "Sheet": "Income Statement"}])
    model = _build_dynamic_credit_analysis(
        tables=tables,
        combined_df=combined_df,
        industry="Technology",
        geography="United States Tier 1",
        business_stage="Mature",
        company_size="Medium",
        years_in_operation=8,
        requested_amount=10_000_000,
    )
    risk_map = dict(zip(model["table"]["Metric"], model["table"]["Risk"]))
    assert risk_map["Debt to EBITDA"] == "Incomplete"
    assert risk_map["Current Ratio"] == "Incomplete"
    assert model["missing_data_warning"] is not None
