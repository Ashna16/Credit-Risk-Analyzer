# Credit Document Intelligence Console (MVP)

This app extracts bank-oriented fields from covenant credit documents, term documents, and credit agreements, with page-level evidence and confidence scores.

## Output principles
- Explicit field names (no generic `terms` bucket)
- `null` when a field is not present
- Every field includes: `value`, `page_number`, `source_snippet`, `confidence`, `pattern_used`
- Multiple facilities supported via `facility_overview[]`

## Bank-style schema sections
- `facility_overview`
- `parties`
- `pricing`
- `dates_tenor`
- `financial_covenants`
- `collateral_security`
- `fees`
- `events_of_default`

## Why these columns
The field dictionary aligns to commonly used bank credit abstraction and risk-reporting data points such as:
- Facility/commitment, obligor/agent party data, pricing spread/base rate, tenor/maturity, collateral/lien and covenant/event flags.
- Regulatory-style aliases (for example: `committed_exposure_global`, `outstanding_balance`, `origination_date`) are included to make downstream mapping easier.

Reference anchors used for schema alignment:
- Federal Reserve FR Y-14Q Corporate Loan data concepts (commitment/exposure/origination/maturity style reporting fields).
- ECB AnaCredit-style loan/obligor/collateral data dimensions (loan-level risk and counterparty fields).
- Standard syndicated credit agreement structures from public SEC exhibits (party roles, pricing, fees, covenants, defaults).

## Run
```bash
cd /Users/ashnaparekh/workspace/ashna_finance_project/loan_doc_extractor_mvp
python3 -m pip install -r requirements.txt
streamlit run app.py
```

## Exports
- Grid CSV
- Grid Excel (`.xlsx`)
- Structured JSON
