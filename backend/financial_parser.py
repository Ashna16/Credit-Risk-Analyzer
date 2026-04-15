from __future__ import annotations

import logging
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import pdfplumber

logger = logging.getLogger("financial_parser")

YEAR_PATTERN = re.compile(r"(19|20)\d{2}")
FY_PATTERN = re.compile(r"\bFY\s?(\d{2,4})\b", flags=re.IGNORECASE)

FIELD_SYNONYMS = {
    "revenue": ["total revenue", "net revenue", "net revenues", "revenue", "revenues", "net sales", "total sales", "totalrevenue"],
    "ebit": ["operating income", "income from operations", "operating profit", "ebit", "operatingincome"],
    "net_income": ["net income", "net income attributable", "net earnings", "profit for the year", "netincomeattributableto"],
    "interest_expense": [
        "interest expense",
        "interestexpense",
        "cash paid for interest",
        "interest paid",
    ],
    "total_assets": ["total assets", "totalassets", "totalcurrentassetsandotherassets"],
    "total_equity": [
        "total equity",
        "total stockholders equity",
        "total shareholders equity",
        "totalblackrock",
        "permanent equity",
        "total permanent equity",
    ],
    "current_assets": ["current assets", "total current assets", "totalcurrentassets"],
    "current_liabilities": ["current liabilities", "total current liabilities", "totalcurrentliabilities"],
    "inventory": ["inventory", "inventories"],
    "st_debt": ["short-term debt", "short term borrowings", "current portion of long-term debt"],
    "lt_debt": ["long-term debt", "long term debt", "long-term borrowings"],
    "operating_cf": [
        "net cash provided by operating activities",
        "cash from operations",
        "operating cash flow",
        "netcashprovidedbyoperatingactivities",
    ],
    "capex": ["capital expenditures", "purchases of property", "capex", "purchase of property and equipment"],
    "da": [
        "depreciation",
        "amortization",
        "depreciation and amortization",
        "amortization and impairment of intangible assets",
    ],
    "cogs": ["cost of revenue", "cost of revenues", "cost of goods sold", "cost of sales", "cogs"],
}

TEXT_TABLE_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "intersection_tolerance": 8,
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "edge_min_length": 3,
    "min_words_vertical": 2,
    "min_words_horizontal": 1,
}

INCOME_FIELDS = {"revenue", "cogs", "ebit", "net_income", "interest_expense", "da", "ebitda"}
BALANCE_FIELDS = {"total_assets", "total_equity", "current_assets", "current_liabilities", "inventory", "st_debt", "lt_debt"}
CASH_FIELDS = {"operating_cf", "capex"}


def _clean_text(value: str) -> str:
    text = value.lower().strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9\s&/\-]", "", text)
    return text


def _detect_unit_multiplier(page_text: str) -> Optional[int]:
    text = page_text.lower()
    if "in billions" in text or "amounts in billions" in text:
        return 1_000_000_000
    if "in millions" in text or "amounts in millions" in text:
        return 1_000_000
    if "in thousands" in text or "amounts in thousands" in text or "usd 000s" in text:
        return 1_000
    return None


def _parse_year(header: str) -> Optional[int]:
    if not header:
        return None
    header = header.strip()
    match = YEAR_PATTERN.search(header)
    if match:
        return int(match.group(0))
    fy_match = FY_PATTERN.search(header)
    if fy_match:
        year = fy_match.group(1)
        if len(year) == 2:
            return int(f"20{year}")
        return int(year)
    return None


def _parse_numeric(cell: str, unit_multiplier: int) -> Optional[float]:
    if cell is None:
        return None
    raw = str(cell).strip()
    if raw in {"", "—", "-", "–", "N/A", "NA"}:
        return None
    negative = False
    if raw.startswith("(") and raw.endswith(")"):
        negative = True
        raw = raw[1:-1]
    raw = raw.replace("$", "").replace(",", "").replace(" ", "")
    if raw == "":
        return None
    try:
        value = float(raw)
    except ValueError:
        return None
    value = -value if negative else value
    return value * unit_multiplier


def _detect_f_page(text: str) -> bool:
    return bool(re.search(r"\bF-\d+\b", text))


def _extract_year_columns(header_row: List[str]) -> Dict[int, int]:
    year_columns: Dict[int, int] = {}
    for idx, cell in enumerate(header_row):
        year = _parse_year(str(cell))
        if year:
            year_columns[year] = idx
    return year_columns


def _extract_year_columns_from_table(table: List[List[str]]) -> Dict[int, int]:
    # Search first rows for explicit year headers; pick the row with most matches.
    best: Dict[int, int] = {}
    for r in table[:4]:
        header = [str(cell or "") for cell in r]
        yc = _extract_year_columns(header)
        if len(yc) > len(best):
            best = yc
    return best


def _remap_year_columns_sec_left_is_latest(year_clean: Dict[int, int]) -> Dict[int, int]:
    """
    Many filers use newest-period-left, but some (e.g. Alphabet 10-K income) use
    oldest-first left-to-right (2023, 2024, 2025). Trust year headers when they
    are strictly ascending or descending by column index. Only remap when the
    header→column assignment looks scrambled—which can happen when pdfplumber
    merges cells oddly—by assigning descending fiscal years to ascending column order.
    """
    year_columns = {int(y): int(c) for y, c in year_clean.items()}
    if len(year_columns) < 2:
        return year_columns
    years_left_to_right = [y for y, _ in sorted(year_columns.items(), key=lambda kv: kv[1])]
    asc = years_left_to_right == sorted(years_left_to_right)
    desc = years_left_to_right == sorted(years_left_to_right, reverse=True)
    if asc or desc:
        return year_columns
    max_year = max(year_columns.keys())
    leftmost_col = min(year_columns.values())
    if year_columns.get(max_year) == leftmost_col:
        return year_columns
    cols_sorted = sorted(set(year_columns.values()))
    years_desc = sorted(year_columns.keys(), reverse=True)
    if len(cols_sorted) == len(years_desc):
        return {int(y): int(c) for y, c in zip(years_desc, cols_sorted)}
    return year_columns


def _max_reasonable_fiscal_year() -> int:
    return datetime.now().year + 1


def _ordered_fiscal_years_for_table(table: List[List[str]], year_columns: Dict[int, int], page_years: List[int]) -> List[int]:
    """
    Left-to-right fiscal years for this table only. Prefers embedded header row cells over
    whole-page year snippets (which mix income, balance, MD&A and cause column drift).
    """
    upper = _max_reasonable_fiscal_year()
    if year_columns:
        return [int(y) for y, _ in sorted(year_columns.items(), key=lambda kv: kv[1])]
    best: List[int] = []
    for r in table[:8]:
        if not r:
            continue
        seq: List[int] = []
        for cell in r:
            y = _parse_year(str(cell or ""))
            if y is not None and 1990 <= int(y) <= upper:
                seq.append(int(y))
        if len(seq) < 2:
            continue
        uniq = list(dict.fromkeys(seq))
        if len(uniq) < 2:
            continue
        diffs = [abs(uniq[i] - uniq[i + 1]) for i in range(len(uniq) - 1)]
        if any(d > 2 for d in diffs):
            continue
        asc = uniq == sorted(uniq)
        desc = uniq == sorted(uniq, reverse=True)
        if not (asc or desc):
            continue
        if len(uniq) > len(best):
            best = uniq
    if len(best) >= 2:
        return best
    if page_years:
        return list(dict.fromkeys(int(y) for y in page_years if 1990 <= int(y) <= upper))
    return []


def _is_primary_financial_statement_page(statement_hint: str, page_text: str, trust: str) -> bool:
    """
    Ignore MD&A / tax / supplemental tables that mention operations or cash but are not the
    face financials, so they don't pollute the canonical multi-year grid.
    """
    if trust == "high":
        return True
    u = (page_text or "").upper()
    if "INDEX TO CONSOLIDATED" in u and "FINANCIAL STATEMENTS" in u:
        return False
    if statement_hint == "income_statement":
        return any(
            x in u
            for x in (
                "CONSOLIDATED STATEMENTS OF INCOME",
                "CONSOLIDATED STATEMENT OF INCOME",
                "CONSOLIDATED STATEMENTS OF OPERATIONS",
                "CONSOLIDATED STATEMENT OF OPERATIONS",
                "CONDENSED CONSOLIDATED STATEMENTS OF OPERATIONS",
                "CONDENSED CONSOLIDATED STATEMENT OF OPERATIONS",
                "STATEMENTS OF OPERATIONS AND COMPREHENSIVE INCOME",
            )
        )
    if statement_hint == "balance_sheet":
        return any(
            x in u
            for x in (
                "CONSOLIDATED BALANCE SHEETS",
                "CONSOLIDATED BALANCE SHEET",
                "CONDENSED CONSOLIDATED BALANCE SHEETS",
                "CONDENSED CONSOLIDATED BALANCE SHEET",
            )
        )
    if statement_hint == "cash_flow":
        return any(
            x in u
            for x in (
                "CONSOLIDATED STATEMENTS OF CASH FLOWS",
                "CONSOLIDATED STATEMENT OF CASH FLOWS",
                "CONDENSED CONSOLIDATED STATEMENTS OF CASH FLOWS",
                "CONDENSED CONSOLIDATED STATEMENT OF CASH FLOWS",
            )
        )
    return False


def _statement_bucket_for_field(statement_hint: str, field_key: str) -> Optional[str]:
    if statement_hint == "income_statement" and field_key in INCOME_FIELDS:
        return "income"
    if statement_hint == "balance_sheet" and field_key in BALANCE_FIELDS:
        return "balance"
    if statement_hint == "cash_flow":
        if field_key in CASH_FIELDS:
            return "cash"
        if field_key == "net_income":
            return "income"
        # Supplemental cash-flow disclosures often carry interest paid; store with P&L for coverage ratios.
        if field_key == "interest_expense":
            return "income"
    return None


def _years_from_text(page_text: str) -> List[int]:
    current_year = datetime.now().year
    lines = [ln.strip() for ln in (page_text or "").splitlines() if ln and ln.strip()]
    context: List[str] = []
    for i, line in enumerate(lines):
        cl = _clean_text(line)
        if any(k in cl for k in ["year ended", "years ended", "fiscal year", "as of december", "as of june"]):
            context.append(line)
            if i + 1 < len(lines):
                context.append(lines[i + 1])
    search_text = " ".join(context) if context else (page_text or "")[:1200]
    years: List[int] = []
    for m in YEAR_PATTERN.finditer(search_text):
        try:
            y = int(m.group(0))
        except Exception:
            continue
        if 1990 <= y <= (current_year + 1):
            years.append(y)
    # Keep encounter order (do not sort), because filing tables usually list latest year first.
    seen: set[int] = set()
    ordered: List[int] = []
    for y in years:
        if y not in seen:
            seen.add(y)
            ordered.append(y)
    return ordered


def _statement_hint_from_text(page_text: str) -> Optional[str]:
    t = _clean_text(page_text or "")
    if "statements of income" in t or "income statement" in t or "statement of operations" in t:
        return "income_statement"
    if "balance sheets" in t or "statement of financial position" in t:
        return "balance_sheet"
    if "cash flows" in t or "cash flow statement" in t:
        return "cash_flow"
    return None


def _field_statement_type(field_key: str) -> Optional[str]:
    if field_key in INCOME_FIELDS:
        return "income_statement"
    if field_key in BALANCE_FIELDS:
        return "balance_sheet"
    if field_key in CASH_FIELDS:
        return "cash_flow"
    return None


def _field_priority(page_text: str, statement_hint: Optional[str], field_key: str, trust: str) -> int:
    target = _field_statement_type(field_key)
    if target is None:
        return 0
    score = 0
    if statement_hint == target:
        score += 2
    t = _clean_text(page_text or "")
    if target == "income_statement" and "consolidated statements of income" in t:
        score += 2
    if target == "balance_sheet" and "consolidated balance sheets" in t:
        score += 2
    if target == "cash_flow" and "consolidated statements of cash flows" in t:
        score += 2
    if trust == "high":
        score += 1
    return score


def _row_numeric_values(row: List[str], unit_multiplier: int) -> List[float]:
    vals: List[float] = []
    for cell in row[1:]:
        v = _parse_numeric(cell, unit_multiplier)
        if v is not None:
            vals.append(v)
    return vals


def _row_label_with_index(row: List[str]) -> tuple[str, int]:
    for idx, cell in enumerate(row[:5]):
        txt = str(cell or "").strip()
        if not txt or txt == "$":
            continue
        if _parse_numeric(txt, 1) is not None:
            continue
        return txt, idx
    return str(row[0] or "").strip(), 0


def _match_field(label: str) -> Optional[str]:
    cleaned = _clean_text(label)
    if not cleaned:
        return None
    for key, synonyms in FIELD_SYNONYMS.items():
        if key == "revenue" and ("cost of revenue" in cleaned or "cost of revenues" in cleaned):
            continue
        for synonym in synonyms:
            s = _clean_text(synonym)
            if cleaned == s or cleaned.startswith(f"{s} ") or cleaned.startswith(f"{s}:"):
                return key
    return None


def _filter_orphan_fiscal_year_rows(rows: List[Dict]) -> List[Dict]:
    """
    Drop rows for years that sit outside the dense fiscal band (e.g. a lone net_income
    pick from a note table labeled with an old year) so clients don't see spurious FYs.
    """
    if not rows:
        return rows
    exclude = {
        "statement_type",
        "selected_year",
        "extraction_confidence",
        "__page_map",
        "__pages",
    }
    per_year_peak: Dict[int, int] = {}
    for r in rows:
        y = r.get("selected_year")
        if y is None:
            continue
        yi = int(y)
        n = sum(1 for k, v in r.items() if k not in exclude and v is not None)
        per_year_peak[yi] = max(per_year_peak.get(yi, 0), n)
    if len(per_year_peak) < 2:
        return rows
    substantive_years = {y for y, d in per_year_peak.items() if d >= 2}
    if not substantive_years:
        return rows
    band_lo = min(substantive_years)
    band_hi = max(substantive_years)
    return [r for r in rows if r.get("selected_year") is not None and band_lo <= int(r["selected_year"]) <= band_hi]


def parse_financial_statements(pdf_path: str) -> List[Dict]:
    income_by_year: Dict[int, Dict[str, Any]] = {}
    balance_by_year: Dict[int, Dict[str, Any]] = {}
    cash_by_year: Dict[int, Dict[str, Any]] = {}
    field_priority_income: Dict[int, Dict[str, int]] = {}
    field_priority_balance: Dict[int, Dict[str, int]] = {}
    field_priority_cash: Dict[int, Dict[str, int]] = {}
    field_pages_income: Dict[int, Dict[str, int]] = {}
    field_pages_balance: Dict[int, Dict[str, int]] = {}
    field_pages_cash: Dict[int, Dict[str, int]] = {}
    unit_multiplier = 1

    with pdfplumber.open(pdf_path) as pdf:
        for pass_idx in (0, 1):
            # Pass 0: default extractor (fast, precise when line boundaries are present)
            # Pass 1: text-table fallback (denser filings where table borders are weak/missing)
            if pass_idx == 1 and (income_by_year or balance_by_year or cash_by_year):
                break
            table_settings = None if pass_idx == 0 else TEXT_TABLE_SETTINGS
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                detected_unit = _detect_unit_multiplier(page_text)
                if detected_unit:
                    unit_multiplier = detected_unit

                trust = "high" if _detect_f_page(page_text) else "normal"
                statement_hint = _statement_hint_from_text(page_text)
                page_years = _years_from_text(page_text)
                tables = page.extract_tables(table_settings=table_settings) if table_settings else page.extract_tables()
                tables = tables or []
                for table in tables:
                    # Many SEC PDFs extract as one-row tables; treat them as valid.
                    if not table or len(table) < 1:
                        continue
                    if statement_hint not in {"income_statement", "balance_sheet", "cash_flow"}:
                        continue
                    if not _is_primary_financial_statement_page(statement_hint, page_text, trust):
                        continue
                    year_columns = _extract_year_columns_from_table(table)
                    if year_columns:
                        year_columns = _remap_year_columns_sec_left_is_latest(year_columns)
                    table_years = _ordered_fiscal_years_for_table(table, year_columns, page_years)
                    if not table_years:
                        continue

                    row_iter = table[1:] if (year_columns and len(table) > 1) else table
                    for row in row_iter:
                        if not row:
                            continue
                        label, label_idx = _row_label_with_index(row)
                        field_key = _match_field(label)
                        if not field_key:
                            continue
                        bucket = _statement_bucket_for_field(statement_hint, field_key)
                        if bucket is None:
                            continue

                        row_values = _row_numeric_values(row[label_idx:], unit_multiplier)
                        fallback_vals: Dict[int, float] = {}
                        if not year_columns and row_values and table_years:
                            uniq = list(dict.fromkeys(int(y) for y in table_years))
                            asc = uniq == sorted(uniq)
                            desc = uniq == sorted(uniq, reverse=True)
                            if asc or desc:
                                use_years = uniq
                            else:
                                use_years = sorted(uniq, reverse=True)
                            n = min(len(row_values), len(use_years))
                            picked = row_values[:n]
                            use_years = use_years[:n]
                            fallback_vals = {int(y): float(v) for y, v in zip(use_years, picked)}

                        if bucket == "income":
                            store = income_by_year
                            field_priority = field_priority_income
                            field_pages = field_pages_income
                        elif bucket == "balance":
                            store = balance_by_year
                            field_priority = field_priority_balance
                            field_pages = field_pages_balance
                        else:
                            store = cash_by_year
                            field_priority = field_priority_cash
                            field_pages = field_pages_cash

                        for year in table_years:
                            value = None
                            if year_columns:
                                col_idx = year_columns.get(year)
                                if col_idx is not None and col_idx < len(row):
                                    value = _parse_numeric(row[col_idx], unit_multiplier)
                            else:
                                value = fallback_vals.get(int(year))

                            if value is None:
                                continue
                            yi = int(year)
                            if yi not in store:
                                store[yi] = {}
                            new_priority = _field_priority(page_text, statement_hint, field_key, trust)
                            prev_priority = field_priority.get(yi, {}).get(field_key, -1)
                            existing = store[yi].get(field_key)
                            if existing is None or new_priority > prev_priority:
                                store[yi][field_key] = value
                                field_priority.setdefault(yi, {})[field_key] = new_priority
                                try:
                                    field_pages.setdefault(yi, {})[field_key] = int(getattr(page, "page_number", 0) or 0)
                                except Exception:
                                    pass

    all_years = sorted(set(income_by_year) | set(balance_by_year) | set(cash_by_year), reverse=True)
    final_rows: List[Dict] = []
    # Newest fiscal year first: helps APIs and clients that scan rows in insertion order.
    for year in all_years:
        inc = income_by_year.get(year, {})
        bal = balance_by_year.get(year, {})
        cf = cash_by_year.get(year, {})
        payload = {
            "selected_year": year,
            "revenue": inc.get("revenue"),
            "cogs": inc.get("cogs"),
            "ebit": inc.get("ebit"),
            "net_income": inc.get("net_income"),
            "interest_expense": inc.get("interest_expense"),
            "da": inc.get("da"),
            "ebitda": inc.get("ebitda"),
            "total_assets": bal.get("total_assets"),
            "total_equity": bal.get("total_equity"),
            "current_assets": bal.get("current_assets"),
            "current_liabilities": bal.get("current_liabilities"),
            "inventory": bal.get("inventory"),
            "st_debt": bal.get("st_debt"),
            "lt_debt": bal.get("lt_debt"),
            "operating_cf": cf.get("operating_cf"),
            "capex": cf.get("capex"),
            "extraction_confidence": None,
        }
        if payload.get("ebitda") is None and payload.get("ebit") is not None and payload.get("da") is not None:
            payload["ebitda"] = payload["ebit"] + payload["da"]
        if payload.get("ebit") is not None and payload.get("revenue") is not None:
            if payload["revenue"] > 1000 and -10 < payload["ebit"] < 10:
                logger.warning("EBIT sanity check failed for year %s", year)
                payload["ebit"] = None

        income_fields = {
            "statement_type": "income_statement",
            "selected_year": year,
            "revenue": payload.get("revenue"),
            "cogs": payload.get("cogs"),
            "ebit": payload.get("ebit"),
            "net_income": payload.get("net_income"),
            "interest_expense": payload.get("interest_expense"),
            "da": payload.get("da"),
            "ebitda": payload.get("ebitda"),
            "__page_map": {},
            "__pages": [],
            "extraction_confidence": None,
        }
        balance_fields = {
            "statement_type": "balance_sheet",
            "selected_year": year,
            "total_assets": payload.get("total_assets"),
            "total_equity": payload.get("total_equity"),
            "current_assets": payload.get("current_assets"),
            "current_liabilities": payload.get("current_liabilities"),
            "inventory": payload.get("inventory"),
            "st_debt": payload.get("st_debt"),
            "lt_debt": payload.get("lt_debt"),
            "__page_map": {},
            "__pages": [],
            "extraction_confidence": None,
        }
        cash_fields = {
            "statement_type": "cash_flow",
            "selected_year": year,
            "operating_cf": payload.get("operating_cf"),
            "capex": payload.get("capex"),
            "__page_map": {},
            "__pages": [],
            "extraction_confidence": None,
        }

        def apply_confidence(row: Dict, exclude: set[str]) -> Optional[Dict]:
            non_null = sum(
                1
                for key, value in row.items()
                if key not in exclude and value is not None
            )
            if non_null == 0:
                return None
            if non_null >= 8:
                row["extraction_confidence"] = "high"
            elif non_null >= 3:
                row["extraction_confidence"] = "basic"
            else:
                row["extraction_confidence"] = "low"
            return row

        for field in INCOME_FIELDS:
            p = field_pages_income.get(year, {}).get(field)
            if p:
                income_fields["__page_map"][field] = p
        for field in BALANCE_FIELDS:
            p = field_pages_balance.get(year, {}).get(field)
            if p:
                balance_fields["__page_map"][field] = p
        for field in CASH_FIELDS:
            p = field_pages_cash.get(year, {}).get(field)
            if p:
                cash_fields["__page_map"][field] = p
        income_fields["__pages"] = sorted(set(income_fields["__page_map"].values()))
        balance_fields["__pages"] = sorted(set(balance_fields["__page_map"].values()))
        cash_fields["__pages"] = sorted(set(cash_fields["__page_map"].values()))

        income_row = apply_confidence(income_fields, {"statement_type", "selected_year", "extraction_confidence", "__page_map", "__pages"})
        balance_row = apply_confidence(balance_fields, {"statement_type", "selected_year", "extraction_confidence", "__page_map", "__pages"})
        cash_row = apply_confidence(cash_fields, {"statement_type", "selected_year", "extraction_confidence", "__page_map", "__pages"})

        for row in (income_row, balance_row, cash_row):
            if row:
                final_rows.append(row)

    final_rows = [row for row in final_rows if row.get("selected_year") is not None]
    final_rows = _filter_orphan_fiscal_year_rows(final_rows)
    return final_rows
