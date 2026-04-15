from sqlalchemy import JSON, Column, DateTime, Float, ForeignKey, Integer, String, func
from sqlalchemy.orm import relationship

from database import Base


class Document(Base):
    __tablename__ = 'documents'

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    file_name = Column(String, nullable=False)
    file_url = Column(String, nullable=False)
    file_size = Column(Float, nullable=False)
    doc_type = Column(String, nullable=False)
    company_name = Column(String, nullable=True)
    ticker = Column(String, nullable=True)
    fiscal_year = Column(Integer, nullable=True)
    period_type = Column(String, nullable=True)
    reporting_unit = Column(String, nullable=True)
    detected_unit = Column(String, nullable=True)
    uploaded_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    extracted_financials = relationship('ExtractedFinancial', back_populates='document', cascade='all, delete-orphan')
    analyses = relationship('CreditAnalysis', back_populates='document', cascade='all, delete-orphan')


class ExtractedFinancial(Base):
    __tablename__ = 'extracted_financials'

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    document_id = Column(Integer, ForeignKey('documents.id', ondelete='CASCADE'), nullable=False, index=True)
    statement_type = Column(String, nullable=False)

    revenue = Column(Float, nullable=True)
    cogs = Column(Float, nullable=True)
    ebitda = Column(Float, nullable=True)
    ebit = Column(Float, nullable=True)
    net_income = Column(Float, nullable=True)
    interest_expense = Column(Float, nullable=True)
    total_assets = Column(Float, nullable=True)
    total_equity = Column(Float, nullable=True)
    current_assets = Column(Float, nullable=True)
    current_liabilities = Column(Float, nullable=True)
    inventory = Column(Float, nullable=True)
    st_debt = Column(Float, nullable=True)
    lt_debt = Column(Float, nullable=True)
    operating_cf = Column(Float, nullable=True)
    capex = Column(Float, nullable=True)
    da = Column(Float, nullable=True)
    raw_fields = Column(JSON, nullable=True)

    selected_year = Column(Integer, nullable=True)
    extraction_confidence = Column(String, nullable=True)
    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    document = relationship('Document', back_populates='extracted_financials')


class CreditAnalysis(Base):
    __tablename__ = 'credit_analyses'

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    document_id = Column(Integer, ForeignKey('documents.id', ondelete='CASCADE'), nullable=False, index=True)

    industry = Column(String, nullable=True)
    geographic_risk = Column(String, nullable=True)
    business_stage = Column(String, nullable=True)
    company_size = Column(String, nullable=True)
    loan_type = Column(String, nullable=True)
    years_in_operation = Column(Integer, nullable=True)
    requested_amount = Column(Float, nullable=True)
    currency_scale = Column(String, nullable=True)

    risk_score = Column(Float, nullable=True)
    risk_band = Column(String, nullable=True)
    policy_limit = Column(Float, nullable=True)
    approval_status = Column(String, nullable=True)
    weighted_score = Column(Float, nullable=True)

    created_at = Column(DateTime(timezone=True), server_default=func.now(), nullable=False)

    document = relationship('Document', back_populates='analyses')
    metric_scores = relationship('MetricScore', back_populates='analysis', cascade='all, delete-orphan')


class MetricScore(Base):
    __tablename__ = 'metric_scores'

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    analysis_id = Column(Integer, ForeignKey('credit_analyses.id', ondelete='CASCADE'), nullable=False, index=True)

    metric_name = Column(String, nullable=False)
    calculated_value = Column(Float, nullable=True)
    industry_threshold = Column(Float, nullable=True)
    base_score = Column(Float, nullable=True)
    adjusted_score = Column(Float, nullable=True)
    status = Column(String, nullable=True)
    risk_level = Column(String, nullable=True)

    analysis = relationship('CreditAnalysis', back_populates='metric_scores')
