from datetime import datetime
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, ConfigDict


class DocumentBase(BaseModel):
    file_name: str
    file_url: str
    file_size: float
    doc_type: str
    company_name: Optional[str] = None
    ticker: Optional[str] = None
    fiscal_year: Optional[int] = None
    period_type: Optional[str] = None
    reporting_unit: Optional[str] = None
    detected_unit: Optional[str] = None


class DocumentRead(DocumentBase):
    id: int
    uploaded_at: datetime

    model_config = ConfigDict(from_attributes=True)


class ExtractedFinancialBase(BaseModel):
    statement_type: str
    revenue: Optional[float] = None
    cogs: Optional[float] = None
    ebitda: Optional[float] = None
    ebit: Optional[float] = None
    net_income: Optional[float] = None
    interest_expense: Optional[float] = None
    total_assets: Optional[float] = None
    total_equity: Optional[float] = None
    current_assets: Optional[float] = None
    current_liabilities: Optional[float] = None
    inventory: Optional[float] = None
    st_debt: Optional[float] = None
    lt_debt: Optional[float] = None
    operating_cf: Optional[float] = None
    capex: Optional[float] = None
    da: Optional[float] = None
    selected_year: Optional[int] = None
    extraction_confidence: Optional[str] = None
    raw_fields: Optional[Dict[str, Any]] = None

    model_config = ConfigDict(extra='allow')


class ExtractedFinancialCreate(ExtractedFinancialBase):
    pass


class ExtractedFinancialRead(ExtractedFinancialBase):
    id: int
    document_id: int
    created_at: datetime

    model_config = ConfigDict(from_attributes=True)


class MetricScoreBase(BaseModel):
    metric_name: str
    calculated_value: Optional[float] = None
    industry_threshold: Optional[float] = None
    base_score: Optional[float] = None
    adjusted_score: Optional[float] = None
    status: Optional[str] = None
    risk_level: Optional[str] = None


class MetricScoreCreate(MetricScoreBase):
    pass


class MetricScoreRead(MetricScoreBase):
    id: int
    analysis_id: int

    model_config = ConfigDict(from_attributes=True)


class CreditAnalysisBase(BaseModel):
    document_id: int
    industry: Optional[str] = None
    geographic_risk: Optional[str] = None
    business_stage: Optional[str] = None
    company_size: Optional[str] = None
    loan_type: Optional[str] = None
    years_in_operation: Optional[int] = None
    requested_amount: Optional[float] = None
    currency_scale: Optional[str] = None
    risk_score: Optional[float] = None
    risk_band: Optional[str] = None
    policy_limit: Optional[float] = None
    approval_status: Optional[str] = None
    weighted_score: Optional[float] = None


class CreditAnalysisCreate(CreditAnalysisBase):
    metric_scores: List[MetricScoreCreate]


class CreditAnalysisRead(CreditAnalysisBase):
    id: int
    created_at: datetime
    metric_scores: List[MetricScoreRead] = []

    model_config = ConfigDict(from_attributes=True)


class DocumentWithFinancials(BaseModel):
    document: DocumentRead
    extracted_financials: List[ExtractedFinancialRead]


class AnalysisWithMetrics(CreditAnalysisRead):
    pass
