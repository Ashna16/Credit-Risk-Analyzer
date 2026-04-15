from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import fitz
import requests
from dotenv import load_dotenv

# Load .env from project root and module folder.
load_dotenv()
load_dotenv(Path(__file__).with_name('.env'))

FieldObj = Dict[str, Any]

MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.0-flash")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/{MODEL_NAME}:generateContent"
GEMINI_TIMEOUT_SECONDS = int(os.getenv("GEMINI_TIMEOUT_SECONDS", "240"))


def _null_field() -> FieldObj:
    return {
        "value": None,
        "page_number": None,
        "source_snippet": None,
        "confidence": None,
        "pattern_used": None,
    }


def _normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def _to_field(obj: Any) -> FieldObj:
    if obj is None:
        return _null_field()
    if isinstance(obj, (str, int, float, bool)):
        return {
            "value": str(obj),
            "page_number": None,
            "source_snippet": None,
            "confidence": 0.6,
            "pattern_used": "gemini-flash",
        }
    if not isinstance(obj, dict):
        return _null_field()

    conf = obj.get("confidence")
    try:
        conf = float(conf) if conf is not None else None
    except (TypeError, ValueError):
        conf = None
    if conf is not None:
        conf = max(0.0, min(1.0, conf))

    page = obj.get("page_number")
    try:
        page = int(page) if page is not None else None
    except (TypeError, ValueError):
        page = None

    return {
        "value": obj.get("value"),
        "page_number": page,
        "source_snippet": _normalize_whitespace(str(obj.get("source_snippet") or "")) or None,
        "confidence": round(conf, 2) if conf is not None else None,
        "pattern_used": "gemini-flash",
    }


def _read_pdf_pages(pdf_path: Path) -> List[str]:
    doc = fitz.open(pdf_path)
    pages: List[str] = []
    for page in doc:
        pages.append(_normalize_whitespace(page.get_text("text")))
    return pages


def _empty_schema() -> Dict[str, Any]:
    return {
        "facility_overview": [
            {
                "facility_id": _null_field(),
                "facility_type": _null_field(),
                "facility_amount_total_commitment": _null_field(),
                "committed_exposure_global": _null_field(),
                "outstanding_balance": _null_field(),
                "unused_commitment_amount": _null_field(),
                "currency": _null_field(),
                "purpose_of_loan": _null_field(),
                "syndicated": _null_field(),
                "agent_bank": _null_field(),
                "origination_date": _null_field(),
                "governing_law": _null_field(),
                "status": _null_field(),
            }
        ],
        "parties": {
            "borrower_name": _null_field(),
            "parent_company": _null_field(),
            "guarantors": _null_field(),
            "lenders": _null_field(),
            "administrative_agent": _null_field(),
            "collateral_agent": _null_field(),
        },
        "pricing": {
            "base_rate_type": _null_field(),
            "margin": _null_field(),
            "spread_grid": _null_field(),
            "default_interest_rate": _null_field(),
            "interest_payment_frequency": _null_field(),
        },
        "dates_tenor": {
            "closing_date": _null_field(),
            "maturity_date": _null_field(),
            "availability_period": _null_field(),
            "amortization_schedule": _null_field(),
            "final_payment_date": _null_field(),
        },
        "financial_covenants": [],
        "collateral_security": {
            "secured_or_unsecured": _null_field(),
            "collateral_type": _null_field(),
            "lien_priority": _null_field(),
            "guarantees_provided": _null_field(),
        },
        "fees": {
            "upfront_fee_percent": _null_field(),
            "commitment_fee_percent": _null_field(),
            "letter_of_credit_fee": _null_field(),
            "agency_fee": _null_field(),
            "prepayment_premium": _null_field(),
        },
        "events_of_default": {
            "payment_default": _null_field(),
            "covenant_breach": _null_field(),
            "cross_default": _null_field(),
            "bankruptcy": _null_field(),
            "change_of_control": _null_field(),
        },
        "amendment_terms": {
            "amendment_number": _null_field(),
            "amendment_effective_date": _null_field(),
            "prior_maturity_date": _null_field(),
            "revised_maturity_date": _null_field(),
            "prior_commitment_amount": _null_field(),
            "revised_commitment_amount": _null_field(),
            "spread_change": _null_field(),
            "consent_fee": _null_field(),
            "waiver_included": _null_field(),
            "covenant_terms_amended": _null_field(),
            "parties_consenting": _null_field(),
            "agent_consent_required": _null_field(),
        },
    }


def _flatten_schema(schema: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    def walk(node: Any, path: str) -> None:
        if isinstance(node, dict):
            if set(node.keys()) >= {"value", "page_number", "source_snippet", "confidence", "pattern_used"}:
                rows.append(
                    {
                        "field_path": path,
                        "value": node.get("value"),
                        "page_number": node.get("page_number"),
                        "source_snippet": node.get("source_snippet"),
                        "confidence": node.get("confidence"),
                    }
                )
            else:
                for k, v in node.items():
                    walk(v, f"{path}.{k}" if path else k)
        elif isinstance(node, list):
            for idx, item in enumerate(node):
                walk(item, f"{path}[{idx}]")

    walk(schema, "")
    return rows


def _extract_json(text: str) -> Dict[str, Any]:
    text = (text or "").strip()
    if not text:
        return {}

    # Try direct JSON first.
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try fenced block.
    m = re.search(r"```json\s*(\{.*\})\s*```", text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    # Last resort: first {...} block.
    m2 = re.search(r"(\{.*\})", text, flags=re.DOTALL)
    if m2:
        try:
            return json.loads(m2.group(1))
        except json.JSONDecodeError:
            return {}
    return {}


def _local_field(value: Optional[str], page_number: Optional[int], source_snippet: Optional[str], confidence: float = 0.45) -> FieldObj:
    return {
        "value": value,
        "page_number": page_number,
        "source_snippet": _normalize_whitespace(source_snippet or "") or None,
        "confidence": confidence if value is not None else None,
        "pattern_used": "local-regex-fallback",
    }


def _local_match(pages: List[str], patterns: List[str], flags: int = re.IGNORECASE) -> FieldObj:
    for page_idx, text in enumerate(pages, start=1):
        for pattern in patterns:
            m = re.search(pattern, text, flags=flags)
            if not m:
                continue
            value = m.group(1).strip() if m.lastindex else m.group(0).strip()
            snippet = text[max(0, m.start() - 60): m.end() + 80]
            return _local_field(value, page_idx, snippet, confidence=0.45)
    return _null_field()


def _rx_field(value: Any, page_number: int, source_snippet: str, confidence: float, pattern_used: str) -> FieldObj:
    return {
        "value": value,
        "page_number": page_number,
        "source_snippet": source_snippet,
        "confidence": round(confidence, 2),
        "pattern_used": pattern_used,
    }


def _rx_snippet(text: str, start: int, end: int, window: int = 180) -> str:
    s = max(0, start - window)
    e = min(len(text), end + window)
    return _normalize_whitespace(text[s:e])


def _rx_clean_entity(text: str) -> str:
    v = _normalize_whitespace(text).strip(" ,.;")
    v = re.sub(r"^(among|and)\s+", "", v, flags=re.IGNORECASE)
    cut_markers = [
        " commitment fee",
        " upfront fee",
        " covenant",
        " shall ",
        " will ",
        ". ",
        "; ",
    ]
    lower = v.lower()
    cut_pos = len(v)
    for marker in cut_markers:
        idx = lower.find(marker)
        if idx != -1:
            cut_pos = min(cut_pos, idx)
    v = v[:cut_pos].strip(" ,.;")
    if len(v.split()) > 14:
        v = " ".join(v.split()[:14]).strip(" ,.;")
    return v.strip(" ,.;")


def _rx_clean_agent_name(text: str) -> str:
    v = _rx_clean_entity(text)
    na_hits = re.findall(r"([A-Z][A-Z0-9.&'\-\s]{2,90},\s*N\.A\.?)", v)
    if na_hits:
        return na_hits[-1].strip(" ,.;")
    if " and " in v.lower():
        v = re.split(r"\band\b", v, flags=re.IGNORECASE)[-1].strip(" ,.;")
    return v


def _valid_party_name(value: str) -> bool:
    v = _normalize_whitespace(value or "")
    if not v or len(v) > 140:
        return False
    bad_fragments = [
        "of which has",
        "shall",
        "business days",
        "failed to confirm",
        "may not pay dividends",
        "during the period",
    ]
    lower = v.lower()
    if any(frag in lower for frag in bad_fragments):
        return False
    token_count = len(v.split())
    if token_count > 18:
        return False
    corp_suffixes = ["inc", "llc", "ltd", "plc", "corp", "corporation", "company", "bank", "n.a", "reit"]
    if any(sfx in lower for sfx in corp_suffixes):
        return True
    letters = [c for c in v if c.isalpha()]
    if not letters:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return upper_ratio >= 0.55


def _rx_find_best(
    pages: List[str],
    patterns: List[tuple[str, float]],
    search_page_limit: Optional[int] = None,
    cleaner=None,
    validator=None,
) -> FieldObj:
    page_cap = search_page_limit or len(pages)
    best: Optional[FieldObj] = None
    for i, page_text in enumerate(pages[:page_cap], start=1):
        if not page_text:
            continue
        for pattern, conf in patterns:
            m = re.search(pattern, page_text, flags=re.IGNORECASE | re.DOTALL)
            if not m:
                continue
            raw = m.group(1) if m.lastindex else m.group(0)
            val = cleaner(raw) if cleaner else _normalize_whitespace(raw)
            if validator is not None and not validator(val):
                continue
            candidate = _rx_field(val, i, _rx_snippet(page_text, m.start(), m.end()), conf, pattern)
            if best is None:
                best = candidate
                continue
            if (candidate.get("confidence") or 0) > (best.get("confidence") or 0):
                best = candidate
                continue
            if (candidate.get("confidence") or 0) == (best.get("confidence") or 0):
                if (candidate.get("page_number") or 9999) < (best.get("page_number") or 9999):
                    best = candidate
    return best or _null_field()


def _rx_find_boolean_field(
    pages: List[str],
    positive_patterns: List[tuple[str, float]],
    negative_patterns: Optional[List[tuple[str, float]]] = None,
    search_page_limit: Optional[int] = None,
) -> FieldObj:
    neg = _rx_find_best(pages, negative_patterns or [], search_page_limit=search_page_limit)
    pos = _rx_find_best(pages, positive_patterns, search_page_limit=search_page_limit)
    if not _field_has_value(pos) and not _field_has_value(neg):
        return _null_field()
    if not _field_has_value(pos):
        neg["value"] = "No"
        return neg
    if not _field_has_value(neg):
        pos["value"] = "Yes"
        return pos
    pos_conf = float(pos.get("confidence") or 0)
    neg_conf = float(neg.get("confidence") or 0)
    if neg_conf > pos_conf:
        neg["value"] = "No"
        return neg
    pos["value"] = "Yes"
    return pos


def _rx_extract_financial_covenants(pages: List[str]) -> List[Dict[str, FieldObj]]:
    cov_patterns = [
        (r"(Maximum\s+Leverage\s+Ratio)[^\.;:\n]{0,120}?([0-9]+(?:\.[0-9]+)?x)", 0.92),
        (r"(Minimum\s+Interest\s+Coverage\s+Ratio)[^\.;:\n]{0,120}?([0-9]+(?:\.[0-9]+)?x)", 0.92),
        (r"(Debt\s+Service\s+Coverage\s+Ratio)[^\.;:\n]{0,120}?([0-9]+(?:\.[0-9]+)?x)", 0.9),
        (r"(Minimum\s+Liquidity)[^\.;:\n]{0,140}?((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)", 0.88),
        (r"(Maximum\s+Senior\s+Secured\s+Leverage\s+Ratio)[^\.;:\n]{0,120}?([0-9]+(?:\.[0-9]+)?x)", 0.91),
    ]
    type_only_patterns = [
        (r"\b(Financial\s+Covenants?)\b", 0.72),
        (r"\b(Net\s+Leverage\s+Ratio)\b", 0.75),
        (r"\b(Interest\s+Coverage\s+Ratio)\b", 0.75),
        (r"\b(Fixed\s+Charge\s+Coverage\s+Ratio)\b", 0.75),
    ]
    freq_patterns = [
        (r"\b(quarterly)\b", 0.8),
        (r"\b(monthly)\b", 0.8),
        (r"\b(semi-annual|semiannual)\b", 0.78),
        (r"\b(annually|annual)\b", 0.78),
    ]
    eq_cure = _rx_find_boolean_field(
        pages,
        positive_patterns=[(r"\b(equity\s+cure\s+right|equity\s+cure\s+is\s+permitted|equity\s+cure)\b", 0.86)],
        negative_patterns=[(r"\b(no\s+equity\s+cure)\b", 0.86)],
        search_page_limit=180,
    )

    results: List[Dict[str, FieldObj]] = []
    seen_types = set()
    for i, page in enumerate(pages[:200], start=1):
        for pat, conf in cov_patterns:
            for m in re.finditer(pat, page, flags=re.IGNORECASE):
                ctype = _normalize_whitespace(m.group(1)).title()
                key = ctype.lower()
                if key in seen_types:
                    continue
                seen_types.add(key)
                threshold = _normalize_whitespace(m.group(2))
                start = max(0, m.start() - 140)
                end = min(len(page), m.end() + 140)
                local = page[start:end]
                tf = _null_field()
                for fpat, fconf in freq_patterns:
                    fm = re.search(fpat, local, flags=re.IGNORECASE)
                    if fm:
                        tf = _rx_field(fm.group(1).title(), i, _rx_snippet(page, start + fm.start(), start + fm.end()), fconf, fpat)
                        break
                src = _rx_field(i, i, _rx_snippet(page, m.start(), m.end()), conf, pat)
                results.append(
                    {
                        "covenant_type": _rx_field(ctype, i, _rx_snippet(page, m.start(1), m.end(1)), conf, pat),
                        "threshold_value": _rx_field(threshold, i, _rx_snippet(page, m.start(2), m.end(2)), conf, pat),
                        "test_frequency": tf,
                        "calculation_definition_page": src,
                        "equity_cure_allowed": eq_cure,
                    }
                )

    # If no ratio thresholds found, still emit covenant type labels when present.
    if not results:
        for i, page in enumerate(pages[:200], start=1):
            for pat, conf in type_only_patterns:
                for m in re.finditer(pat, page, flags=re.IGNORECASE):
                    ctype = _normalize_whitespace(m.group(1)).title()
                    key = ctype.lower()
                    if key in seen_types:
                        continue
                    seen_types.add(key)
                    results.append(
                        {
                            "covenant_type": _rx_field(ctype, i, _rx_snippet(page, m.start(), m.end()), conf, pat),
                            "threshold_value": _null_field(),
                            "test_frequency": _null_field(),
                            "calculation_definition_page": _rx_field(i, i, _rx_snippet(page, m.start(), m.end()), conf, pat),
                            "equity_cure_allowed": eq_cure,
                        }
                    )
    return results


def _filename_doc_type(pdf_path: Path) -> Optional[str]:
    n = pdf_path.name.lower()
    if "term" in n and "sheet" in n:
        return "Term Sheet"
    if "covenant" in n or "compliance" in n:
        return "Covenant Compliance Certificate"
    if "fee" in n:
        return "Fee Letter"
    if "security" in n:
        return "Security Agreement"
    if "amend" in n:
        return "Amendment"
    if "credit" in n:
        return "Credit Agreement"
    return None


def _field_has_value(field: Any) -> bool:
    if not isinstance(field, dict):
        return False
    v = field.get("value")
    return v is not None and str(v).strip() != ""


def _choose_better_field(current: FieldObj, candidate: FieldObj) -> FieldObj:
    if not _field_has_value(candidate):
        return current
    if not _field_has_value(current):
        return candidate

    current_conf = float(current.get("confidence") or 0)
    cand_conf = float(candidate.get("confidence") or 0)
    current_val = str(current.get("value") or "")
    cand_val = str(candidate.get("value") or "")

    if cand_conf >= current_conf + 0.15:
        return candidate
    if current_conf < 0.5 and cand_conf >= current_conf:
        return candidate
    if len(current_val) > 150 and len(cand_val) < len(current_val):
        return candidate
    return current


def _regex_candidates_for_core_fields(pages: List[str]) -> Dict[str, FieldObj]:
    return {
        "parties.borrower_name": _rx_find_best(
            pages,
            [
                (r"among\s+([A-Z][A-Z0-9.,&'\-\s]{2,140}?)\s*,\s+as\s+the\s+Borrower", 0.96),
                (r"among\s+([A-Z][A-Z0-9.,&'\-\s]{2,140}?)\s*,\s+The\s+Several\s+Lenders", 0.95),
                (r"among\s+([A-Z][A-Z0-9.,&'\-\s]{2,140}?)\s*,\s+The\s+Lenders", 0.95),
                (r"borrower\s*[:\-]\s*([A-Z][A-Z0-9.,&'\-\s]{2,140})", 0.9),
            ],
            search_page_limit=24,
            cleaner=_rx_clean_entity,
            validator=_valid_party_name,
        ),
        "parties.administrative_agent": _rx_find_best(
            pages,
            [
                (r"(?:and\s+)?([A-Z][A-Z0-9.,&'\-\s]{2,120}?)\s*,\s+as\s+Administrative\s+Agent", 0.94),
                (r"(?:and\s+)?([A-Z][A-Z0-9.,&'\-\s]{2,120}?)\s*,\s+as\s+the\s+Administrative\s+Agent", 0.94),
            ],
            search_page_limit=24,
            cleaner=_rx_clean_agent_name,
        ),
        "parties.guarantors": _rx_find_best(
            pages,
            [
                (r"([A-Z][A-Z0-9.,&'\-\s]{2,180}?)\s*,\s+as\s+Guarantors", 0.9),
                (r"certain\s+subsidiaries\s+of\s+([A-Z][A-Z0-9.,&'\-\s]{2,160})\s+from\s+time\s+to\s+time\s+party\s+hereto,\s+as\s+Guarantors", 0.92),
            ],
            search_page_limit=40,
            cleaner=_rx_clean_entity,
            validator=_valid_party_name,
        ),
        "parties.lenders": _rx_find_best(
            pages,
            [
                (r"(The\s+Lenders\s+and\s+Issuing\s+Banks\s+from\s+time\s+to\s+time\s+Party\s+Hereto)", 0.9),
                (r"(The\s+Several\s+Lenders\s+from\s+Time\s+to\s+Time\s+Parties\s+Thereto)", 0.9),
                (r"(The\s+Lenders\s+and\s+L/C\s+Issuers\s+Party\s+Hereto)", 0.88),
            ],
            search_page_limit=35,
            cleaner=lambda x: _normalize_whitespace(x),
        ),
        "facility_overview.0.facility_type": _rx_find_best(
            pages,
            [
                (r"\b(Revolving\s+Credit\s+Facility|Revolving\s+Facility|Revolver)\b", 0.92),
                (r"\b(Term\s+Loan\s+A|Term\s+Loan\s+B|Term\s+Loan)\b", 0.9),
            ],
            search_page_limit=60,
            cleaner=lambda x: _normalize_whitespace(x).title(),
        ),
        "facility_overview.0.facility_amount_total_commitment": _rx_find_best(
            pages,
            [
                (r"(?:aggregate\s+)?(?:commitments?|facility(?:\s+amount)?)\s+(?:of|in\s+an\s+amount\s+of)?\s*((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)", 0.9),
                (r"principal\s+amount\s+of\s*((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)", 0.86),
            ],
            search_page_limit=80,
        ),
        "facility_overview.0.agent_bank": _rx_find_best(
            pages,
            [
                (r"(?:and\s+)?([A-Z][A-Z0-9.,&'\-\s]{2,120}?)\s*,\s+as\s+Administrative\s+Agent", 0.93),
                (r"(?:and\s+)?([A-Z][A-Z0-9.,&'\-\s]{2,120}?)\s*,\s+as\s+Agent", 0.82),
            ],
            search_page_limit=20,
            cleaner=lambda x: _rx_clean_agent_name(x),
        ),
        "pricing.base_rate_type": _rx_find_best(
            pages,
            [
                (r"\b(SOFR)\b", 0.9),
                (r"\b(LIBOR)\b", 0.88),
                (r"\b(Prime\s+Rate|Prime)\b", 0.86),
                (r"\b(Fixed\s+Rate|Fixed)\b", 0.84),
            ],
            search_page_limit=100,
            cleaner=lambda x: _normalize_whitespace(x).upper().replace(" RATE", ""),
        ),
        "pricing.margin": _rx_find_best(
            pages,
            [
                (r"(?:SOFR|LIBOR|Base\s+Rate|Prime)\s*(?:\+|plus)\s*([0-9]+(?:\.[0-9]+)?\s*(?:%|bps|basis\s+points))", 0.9),
                (r"Applicable\s+Rate\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?\s*(?:%|bps|basis\s+points))", 0.84),
            ],
            search_page_limit=120,
        ),
    }


def _set_field_by_path(schema: Dict[str, Any], path: str, field: FieldObj) -> None:
    parts = path.split(".")
    node: Any = schema
    for p in parts[:-1]:
        if p.isdigit():
            node = node[int(p)]
        else:
            node = node[p]
    key = parts[-1]
    if key.isdigit():
        node[int(key)] = field
    else:
        node[key] = field


def _get_field_by_path(schema: Dict[str, Any], path: str) -> FieldObj:
    parts = path.split(".")
    node: Any = schema
    for p in parts:
        if p.isdigit():
            node = node[int(p)]
        else:
            node = node[p]
    return node


def _enrich_schema_with_regex(schema: Dict[str, Any], pages: List[str]) -> None:
    candidates = _regex_candidates_for_core_fields(pages)
    for path, candidate in candidates.items():
        # normalize "facility_overview.0.*" access into existing schema object
        if path.startswith("facility_overview.0.") and not schema.get("facility_overview"):
            continue
        current = _get_field_by_path(schema, path)
        merged = _choose_better_field(current, candidate)
        if merged is not current:
            _set_field_by_path(schema, path, merged)

    for key in ["borrower_name", "parent_company", "guarantors", "lenders", "administrative_agent", "collateral_agent"]:
        f = schema["parties"].get(key) or _null_field()
        if _field_has_value(f) and not _valid_party_name(str(f.get("value"))):
            schema["parties"][key] = _null_field()


def _local_document_type(pages: List[str]) -> str:
    text = " ".join(pages[:5]).lower()
    if "term sheet" in text:
        return "Term Sheet"
    if "covenant compliance certificate" in text or ("covenant" in text and "certificate" in text):
        return "Covenant Compliance Certificate"
    if "fee letter" in text:
        return "Fee Letter"
    if "security agreement" in text:
        return "Security Agreement"
    if "amendment" in text:
        return "Amendment"
    if any(k in text for k in ["credit agreement", "facility", "borrower", "lender"]):
        return "Credit Agreement"
    return "Unknown"


def _local_regex_parse(pages: List[str]) -> Dict[str, Any]:
    raw: Dict[str, Any] = {
        "document_type": _local_document_type(pages),
        "facility_overview": [
            {
                "facility_id": None,
                "facility_type": _local_match(pages, [r"\b(revolving credit facility|term loan [a-z]?|credit facility)\b"]),
                "facility_amount_total_commitment": _local_match(
                    pages,
                    [
                        r"(?:aggregate\s+)?(?:commitments?|facility(?:\s+amount)?)\s+(?:of|in\s+an\s+amount\s+of)?\s*((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)",
                        r"principal\s+amount\s+of\s*((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)",
                        r"(?:total\s+)?(?:commitment|facility amount|principal amount)[:\s$]{0,12}([$]?[0-9][0-9,]*(?:\.[0-9]+)?(?:\s*(?:million|billion|m|bn))?)",
                    ],
                ),
                "committed_exposure_global": None,
                "outstanding_balance": None,
                "unused_commitment_amount": None,
                "currency": _local_match(pages, [r"\b(USD|EUR|GBP|CAD|AUD|JPY)\b"]),
                "purpose_of_loan": _local_match(pages, [r"purpose[:\s]{1,6}([^\n]{3,120})"]),
                "syndicated": _local_match(pages, [r"\b(syndicated)\b"]),
                "agent_bank": _local_match(
                    pages,
                    [
                        r"(?:administrative|facility|collateral)?\s*agent[:\s]{1,6}([A-Za-z0-9&,\.\- ]{3,80})",
                        r"agent bank[:\s]{1,6}([A-Za-z0-9&,\.\- ]{3,80})",
                    ],
                ),
                "origination_date": _local_match(
                    pages,
                    [r"(?:closing|effective|origination)\s+date[:\s]{1,6}([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})"],
                ),
                "governing_law": _local_match(pages, [r"governing law[:\s]{1,6}([A-Za-z ,\-]{3,80})"]),
                "status": _local_match(pages, [r"\b(active|amended|terminated|closed)\b"]),
            }
        ],
        "parties": {
            "borrower_name": _local_match(
                pages,
                [
                    r"among\s+([A-Z][A-Z0-9.,&'\-\s]{2,140}?)\s*,\s+as\s+the\s+Borrower",
                    r"borrower\s*[:\-]\s*([A-Z][A-Z0-9.,&'\-\s]{2,140})",
                    r"between\s+([A-Z][A-Z0-9&,\.\- ]{3,120})\s+as borrower",
                ],
            ),
            "parent_company": None,
            "guarantors": _local_match(pages, [r"guarantor[s]?[:\s]{1,6}([A-Za-z0-9&,\.\- ;]{3,120})"]),
            "lenders": _local_match(pages, [r"lender[s]?[:\s]{1,6}([A-Za-z0-9&,\.\- ;]{3,160})"]),
            "administrative_agent": _local_match(pages, [r"administrative agent[:\s]{1,6}([A-Za-z0-9&,\.\- ]{3,120})"]),
            "collateral_agent": _local_match(pages, [r"collateral agent[:\s]{1,6}([A-Za-z0-9&,\.\- ]{3,120})"]),
        },
        "pricing": {
            "base_rate_type": _local_match(pages, [r"\b(SOFR|LIBOR|Prime Rate|Base Rate)\b"]),
            "margin": _local_match(pages, [r"(?:margin|spread)[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
            "spread_grid": None,
            "default_interest_rate": _local_match(pages, [r"default interest(?: rate)?[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
            "interest_payment_frequency": _local_match(pages, [r"(monthly|quarterly|semi-annual|annual)"]),
        },
        "dates_tenor": {
            "closing_date": _local_match(
                pages, [r"closing date[:\s]{1,6}([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})"]
            ),
            "maturity_date": _local_match(
                pages, [r"maturity date[:\s]{1,6}([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})"]
            ),
            "availability_period": None,
            "amortization_schedule": None,
            "final_payment_date": None,
        },
        "financial_covenants": [],
        "collateral_security": {
            "secured_or_unsecured": _local_match(pages, [r"\b(secured|unsecured)\b"]),
            "collateral_type": _local_match(pages, [r"collateral(?: type)?[:\s]{1,6}([A-Za-z0-9&,\.\- ;]{3,120})"]),
            "lien_priority": _local_match(pages, [r"lien priority[:\s]{1,6}([A-Za-z0-9&,\.\- ]{3,120})"]),
            "guarantees_provided": _local_match(pages, [r"guarantee[s]?[:\s]{1,6}([A-Za-z0-9&,\.\- ;]{3,120})"]),
        },
        "fees": {
            "upfront_fee_percent": _local_match(pages, [r"upfront fee[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
            "commitment_fee_percent": _local_match(
                pages,
                [
                    r"commitment\s+fee\s*(?:will\s+be\s+set\s+at|[:\s]{1,8})\s*([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))",
                    r"unused\s+commitment\s+fee\s*(?:will\s+be\s+set\s+at|[:\s]{1,8})\s*([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))",
                ],
            ),
            "letter_of_credit_fee": _local_match(pages, [r"letter of credit fee[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
            "agency_fee": _local_match(pages, [r"agency fee[:\s]{1,6}([$]?[0-9][0-9,]*(?:\.[0-9]+)?)"]),
            "prepayment_premium": _local_match(pages, [r"prepayment premium[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
        },
        "events_of_default": {
            "payment_default": _local_match(pages, [r"\b(payment default)\b"]),
            "covenant_breach": _local_match(pages, [r"\b(covenant breach)\b"]),
            "cross_default": _local_match(pages, [r"\b(cross[- ]default)\b"]),
            "bankruptcy": _local_match(pages, [r"\b(bankruptcy|insolvency)\b"]),
            "change_of_control": _local_match(pages, [r"\b(change of control)\b"]),
        },
        "amendment_terms": {
            "amendment_number": _local_match(pages, [r"amendment (?:no\.?|number)[:\s]{1,6}([A-Za-z0-9\-]+)"]),
            "amendment_effective_date": _local_match(
                pages, [r"amendment effective date[:\s]{1,6}([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})"]
            ),
            "prior_maturity_date": None,
            "revised_maturity_date": None,
            "prior_commitment_amount": None,
            "revised_commitment_amount": None,
            "spread_change": _local_match(pages, [r"spread change[:\s]{1,6}([0-9]+(?:\.[0-9]+)?\s*(?:%|bps))"]),
            "consent_fee": _local_match(pages, [r"consent fee[:\s]{1,6}([$]?[0-9][0-9,]*(?:\.[0-9]+)?)"]),
            "waiver_included": _local_match(pages, [r"\b(waiver)\b"]),
            "covenant_terms_amended": _local_match(pages, [r"\b(covenant terms amended)\b"]),
            "parties_consenting": _local_match(pages, [r"parties consenting[:\s]{1,6}([A-Za-z0-9&,\.\- ;]{3,120})"]),
            "agent_consent_required": _local_match(pages, [r"agent consent required[:\s]{1,6}([A-Za-z]+)"]),
        },
    }
    for party_key in ["borrower_name", "parent_company", "guarantors", "lenders", "administrative_agent", "collateral_agent"]:
        f = raw.get("parties", {}).get(party_key)
        if isinstance(f, dict) and f.get("value") not in (None, "", "null", "None"):
            cleaned = _rx_clean_entity(str(f.get("value")))
            if _valid_party_name(cleaned):
                f["value"] = cleaned
            else:
                raw["parties"][party_key] = _null_field()
    raw["financial_covenants"] = _rx_extract_financial_covenants(pages)
    return raw


def _gemini_parse(pdf_path: Path) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        raise RuntimeError("GEMINI_API_KEY is missing. Set it in .env")

    prompt = """
You are a banking credit-document extraction engine.
Extract structured fields from the provided PDF and return JSON only.

Rules:
- For EVERY field, return: value, page_number, source_snippet, confidence.
- confidence must be between 0 and 1.
- page_number is 1-based.
- If field not present, return value=null, page_number=null, source_snippet=null, confidence=null.
- Use explicit yes/no values for boolean-like fields where possible.
- If multiple facilities exist, include all in facility_overview array.
- If multiple covenants exist, include all in financial_covenants array.

Return JSON with this structure:
{
  "document_type": "...",
  "facility_overview": [
    {
      "facility_id": {"value":...,"page_number":...,"source_snippet":...,"confidence":...},
      "facility_type": {...},
      "facility_amount_total_commitment": {...},
      "committed_exposure_global": {...},
      "outstanding_balance": {...},
      "unused_commitment_amount": {...},
      "currency": {...},
      "purpose_of_loan": {...},
      "syndicated": {...},
      "agent_bank": {...},
      "origination_date": {...},
      "governing_law": {...},
      "status": {...}
    }
  ],
  "parties": {
    "borrower_name": {...},
    "parent_company": {...},
    "guarantors": {...},
    "lenders": {...},
    "administrative_agent": {...},
    "collateral_agent": {...}
  },
  "pricing": {
    "base_rate_type": {...},
    "margin": {...},
    "spread_grid": {...},
    "default_interest_rate": {...},
    "interest_payment_frequency": {...}
  },
  "dates_tenor": {
    "closing_date": {...},
    "maturity_date": {...},
    "availability_period": {...},
    "amortization_schedule": {...},
    "final_payment_date": {...}
  },
  "financial_covenants": [
    {
      "covenant_type": {...},
      "threshold_value": {...},
      "test_frequency": {...},
      "calculation_definition_page": {...},
      "equity_cure_allowed": {...}
    }
  ],
  "collateral_security": {
    "secured_or_unsecured": {...},
    "collateral_type": {...},
    "lien_priority": {...},
    "guarantees_provided": {...}
  },
  "fees": {
    "upfront_fee_percent": {...},
    "commitment_fee_percent": {...},
    "letter_of_credit_fee": {...},
    "agency_fee": {...},
    "prepayment_premium": {...}
  },
  "events_of_default": {
    "payment_default": {...},
    "covenant_breach": {...},
    "cross_default": {...},
    "bankruptcy": {...},
    "change_of_control": {...}
  },
  "amendment_terms": {
    "amendment_number": {...},
    "amendment_effective_date": {...},
    "prior_maturity_date": {...},
    "revised_maturity_date": {...},
    "prior_commitment_amount": {...},
    "revised_commitment_amount": {...},
    "spread_change": {...},
    "consent_fee": {...},
    "waiver_included": {...},
    "covenant_terms_amended": {...},
    "parties_consenting": {...},
    "agent_consent_required": {...}
  }
}
""".strip()

    pdf_bytes = pdf_path.read_bytes()
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": "application/pdf",
                            "data": base64.b64encode(pdf_bytes).decode("utf-8"),
                        }
                    },
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }

    resp = requests.post(
        GEMINI_ENDPOINT,
        params={"key": GEMINI_API_KEY},
        json=payload,
        timeout=GEMINI_TIMEOUT_SECONDS,
    )
    if resp.status_code >= 400:
        try:
            err_msg = (resp.json().get("error") or {}).get("message") or resp.text
        except Exception:
            err_msg = resp.text
        if resp.status_code == 429:
            raise RuntimeError(
                "Gemini quota exceeded or rate-limited (HTTP 429). "
                "Check billing/quotas for this API key and retry."
            )
        raise RuntimeError(f"Gemini API error {resp.status_code}: {_normalize_whitespace(str(err_msg))[:400]}")
    data = resp.json()

    parts = (
        data.get("candidates", [{}])[0]
        .get("content", {})
        .get("parts", [])
    )
    text = "\n".join(str(p.get("text", "")) for p in parts)
    parsed = _extract_json(text)
    if not parsed:
        raise RuntimeError("Gemini response did not include parseable JSON")
    return parsed


def _normalize_schema(raw: Dict[str, Any]) -> Dict[str, Any]:
    schema = _empty_schema()

    fac_in = raw.get("facility_overview") or []
    facilities: List[Dict[str, FieldObj]] = []
    for idx, fac in enumerate(fac_in):
        if not isinstance(fac, dict):
            continue
        facilities.append(
            {
                "facility_id": _to_field(fac.get("facility_id") or {"value": f"facility_{idx + 1}"}),
                "facility_type": _to_field(fac.get("facility_type")),
                "facility_amount_total_commitment": _to_field(fac.get("facility_amount_total_commitment")),
                "committed_exposure_global": _to_field(fac.get("committed_exposure_global")),
                "outstanding_balance": _to_field(fac.get("outstanding_balance")),
                "unused_commitment_amount": _to_field(fac.get("unused_commitment_amount")),
                "currency": _to_field(fac.get("currency")),
                "purpose_of_loan": _to_field(fac.get("purpose_of_loan")),
                "syndicated": _to_field(fac.get("syndicated")),
                "agent_bank": _to_field(fac.get("agent_bank")),
                "origination_date": _to_field(fac.get("origination_date")),
                "governing_law": _to_field(fac.get("governing_law")),
                "status": _to_field(fac.get("status")),
            }
        )
    if facilities:
        schema["facility_overview"] = facilities

    for key in schema["parties"].keys():
        schema["parties"][key] = _to_field((raw.get("parties") or {}).get(key))
    for key in schema["pricing"].keys():
        schema["pricing"][key] = _to_field((raw.get("pricing") or {}).get(key))
    for key in schema["dates_tenor"].keys():
        schema["dates_tenor"][key] = _to_field((raw.get("dates_tenor") or {}).get(key))
    for key in schema["collateral_security"].keys():
        schema["collateral_security"][key] = _to_field((raw.get("collateral_security") or {}).get(key))
    for key in schema["fees"].keys():
        schema["fees"][key] = _to_field((raw.get("fees") or {}).get(key))
    for key in schema["events_of_default"].keys():
        schema["events_of_default"][key] = _to_field((raw.get("events_of_default") or {}).get(key))
    for key in schema["amendment_terms"].keys():
        schema["amendment_terms"][key] = _to_field((raw.get("amendment_terms") or {}).get(key))

    covs_in = raw.get("financial_covenants") or []
    covs: List[Dict[str, FieldObj]] = []
    for cov in covs_in:
        if not isinstance(cov, dict):
            continue
        covs.append(
            {
                "covenant_type": _to_field(cov.get("covenant_type")),
                "threshold_value": _to_field(cov.get("threshold_value")),
                "test_frequency": _to_field(cov.get("test_frequency")),
                "calculation_definition_page": _to_field(cov.get("calculation_definition_page")),
                "equity_cure_allowed": _to_field(cov.get("equity_cure_allowed")),
            }
        )
    schema["financial_covenants"] = covs

    return schema


def _fallback_covenants_for_certificate(pages: List[str]) -> List[Dict[str, FieldObj]]:
    ratio_patterns = [
        (r"(Net\s+Leverage\s+Ratio)[^\.;:\n]{0,90}?([0-9]+(?:\.[0-9]+)?x)", 0.78),
        (r"(Interest\s+Coverage\s+Ratio)[^\.;:\n]{0,90}?([0-9]+(?:\.[0-9]+)?x)", 0.78),
        (r"(Fixed\s+Charge\s+Coverage\s+Ratio)[^\.;:\n]{0,90}?([0-9]+(?:\.[0-9]+)?x)", 0.78),
        (r"(Debt\s+Service\s+Coverage\s+Ratio)[^\.;:\n]{0,90}?([0-9]+(?:\.[0-9]+)?x)", 0.78),
        (r"(Maximum\s+Leverage\s+Ratio)[^\.;:\n]{0,90}?([0-9]+(?:\.[0-9]+)?x)", 0.78),
    ]
    for i, page in enumerate(pages[:200], start=1):
        for pat, conf in ratio_patterns:
            m = re.search(pat, page, flags=re.IGNORECASE)
            if m:
                return [
                    {
                        "covenant_type": _rx_field(_normalize_whitespace(m.group(1)).title(), i, _rx_snippet(page, m.start(1), m.end(1)), conf, pat),
                        "threshold_value": _rx_field(_normalize_whitespace(m.group(2)), i, _rx_snippet(page, m.start(2), m.end(2)), conf, pat),
                        "test_frequency": _null_field(),
                        "calculation_definition_page": _rx_field(i, i, _rx_snippet(page, m.start(), m.end()), conf, pat),
                        "equity_cure_allowed": _null_field(),
                    }
                ]
    # Last resort: preserve the sheet with a meaningful covenant type label.
    return [
        {
            "covenant_type": _local_field("Covenant Compliance", 1, "Fallback covenant type from document classification.", 0.55),
            "threshold_value": _null_field(),
            "test_frequency": _null_field(),
            "calculation_definition_page": _null_field(),
            "equity_cure_allowed": _null_field(),
        }
    ]


def _fallback_fill_core_financial(schema: Dict[str, Any], pages: List[str]) -> None:
    facility = schema.get("facility_overview", [{}])[0]
    if facility and not _field_has_value(facility.get("facility_amount_total_commitment")):
        amt = _rx_find_best(
            pages,
            [
                (r"(?:aggregate\s+)?(?:commitments?|facility(?:\s+amount)?|loan\s+amount)\s+(?:of|in\s+an\s+amount\s+of)?\s*((?:USD|US\$|\$)\s?[0-9][0-9,\.]+)", 0.82),
                (r"((?:USD|US\$|\$)\s?[0-9][0-9,\.]+\s*(?:million|billion|m|bn)?)", 0.68),
            ],
            search_page_limit=220,
        )
        if _field_has_value(amt):
            facility["facility_amount_total_commitment"] = amt

    if facility and not _field_has_value(facility.get("governing_law")):
        gov = _rx_find_best(
            pages,
            [
                (r"governing\s+law[^A-Za-z]{0,20}(New\s+York|Delaware|California|England\s+and\s+Wales|Ontario)", 0.8),
                (r"laws\s+of\s+the\s+State\s+of\s+([A-Za-z ]{3,40})", 0.75),
            ],
            search_page_limit=240,
            cleaner=lambda x: _normalize_whitespace(x).title(),
        )
        if _field_has_value(gov):
            facility["governing_law"] = gov

    dates = schema.get("dates_tenor", {})
    if dates and not _field_has_value(dates.get("maturity_date")):
        mat = _rx_find_best(
            pages,
            [
                (r"maturity\s+date\s*[:\-]?\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})", 0.84),
                (r"extend,\s+to\s+([A-Za-z]+\s+\d{1,2},\s+\d{4}),\s+the\s+maturity", 0.8),
                (r"maturity\s+of\s+their\s+existing[^\.;]{0,80}to\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", 0.78),
            ],
            search_page_limit=260,
        )
        if _field_has_value(mat):
            dates["maturity_date"] = mat

    if dates and not _field_has_value(dates.get("closing_date")):
        clo = _rx_find_best(
            pages,
            [
                (r"dated\s+as\s+of\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", 0.78),
                (r"closing\s+date\s*[:\-]?\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})", 0.78),
            ],
            search_page_limit=80,
        )
        if _field_has_value(clo):
            dates["closing_date"] = clo


def extract_fields(pdf_path: str | Path) -> Dict[str, Any]:
    pdf_path = Path(pdf_path)
    pages = _read_pdf_pages(pdf_path)

    raw: Dict[str, Any] = {}
    error: Optional[str] = None
    warning: Optional[str] = None
    parser_used = f"gemini:{MODEL_NAME}"
    try:
        raw = _gemini_parse(pdf_path)
    except Exception as e:
        err_msg = str(e)
        # Always fall back so the app keeps working without a key, on quota, or on API errors.
        raw = _local_regex_parse(pages)
        parser_used = "local-regex-fallback"
        error = None
        if not GEMINI_API_KEY or "GEMINI_API_KEY is missing" in err_msg:
            warning = (
                "GEMINI_API_KEY is not set. Using local pattern-based extraction only. "
                "Set GEMINI_API_KEY in loan_doc_extractor_mvp/.env for full Gemini extraction."
            )
        elif "429" in err_msg or "quota" in err_msg.lower() or "rate" in err_msg.lower() and "limit" in err_msg.lower():
            warning = "Gemini quota or rate limit hit. Using local fallback extraction (reduced accuracy)."
        else:
            warning = f"Gemini call failed ({err_msg[:240]}). Using local fallback extraction (reduced accuracy)."

    filename_type = _filename_doc_type(pdf_path)
    if filename_type:
        raw = dict(raw or {})
        raw["document_type"] = filename_type

    schema = _normalize_schema(raw) if raw else _empty_schema()
    _enrich_schema_with_regex(schema, pages)
    _fallback_fill_core_financial(schema, pages)
    if raw.get("document_type") == "Covenant Compliance Certificate" and not schema.get("financial_covenants"):
        schema["financial_covenants"] = _fallback_covenants_for_certificate(pages)
    if raw.get("document_type") in {"Credit Agreement", "Security Agreement", "Term Sheet"} and not schema.get("financial_covenants"):
        schema["financial_covenants"] = _fallback_covenants_for_certificate(pages)
    flat_fields = _flatten_schema(schema)

    document_type = (raw.get("document_type") if isinstance(raw, dict) else None) or None

    summary = {
        "filename": pdf_path.name,
        "document_type": document_type,
        "total_pages": len(pages),
        "total_field_entries": len(flat_fields),
        "non_null_field_entries": sum(1 for r in flat_fields if r.get("value") is not None),
        "facility_count": len(schema.get("facility_overview", [])),
        "financial_covenant_count": len(schema.get("financial_covenants", [])),
        "parser": parser_used,
        "error": error,
        "warning": warning,
    }

    return {
        "schema_version": "2.0",
        "summary": summary,
        "extraction": schema,
        "flat_fields": flat_fields,
    }
