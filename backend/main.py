import logging
import os
import shutil
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from sqlalchemy.orm import Session, selectinload

import models
import schemas
from database import SessionLocal, engine

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / '.env')
PROJECT_ROOT = BASE_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.append(str(PROJECT_ROOT))

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', './uploads')
UPLOAD_DIR = (BASE_DIR / UPLOAD_FOLDER).resolve() if not Path(UPLOAD_FOLDER).is_absolute() else Path(UPLOAD_FOLDER)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

models.Base.metadata.create_all(bind=engine)

_OPENAPI_TAGS = [
    {
        'name': 'Credit API',
        'description': '**Use these first.** Stable paths under `/api/credit/...` (upload, financial rows, documents list, health).',
    },
    {'name': 'System', 'description': 'Health and liveness probes.'},
    {'name': 'Extraction', 'description': 'Legacy alias: upload + parse (same handler as Credit API upload).'},
    {'name': 'Persist', 'description': 'Legacy alias: insert financial rows and credit analyses.'},
    {'name': 'Fetch', 'description': 'Legacy alias: read documents and analyses.'},
]

app = FastAPI(
    title='Credit & Extraction API',
    openapi_tags=_OPENAPI_TAGS,
    description=(
        'Document upload, financial extraction, and credit-analysis persistence. '
        'Prefer **Credit API** paths (`/api/credit/...`); `/api/extract`, `/api/persist`, and `/api/fetch` are backward-compatible aliases.\n\n'
        '**Swagger tip:** For `{document_id}` path routes, click **Try it out**, then enter the integer **id** from '
        '`POST /api/credit/documents/upload-and-parse` or `GET /api/credit/documents`. '
        'Alternatively use **`GET /api/credit/query/...`** endpoints below — they take **`document_id` as a query parameter** '
        '(often easier to fill in Swagger).'
    ),
)

ALLOWED_ORIGINS = [
    'http://localhost:9000',
    'http://127.0.0.1:9000',
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*'],
)

# Serve uploaded PDFs at /uploads/<filename>
app.mount('/uploads', StaticFiles(directory=str(UPLOAD_DIR)), name='uploads')

ALLOWED_DOC_TYPES = {'10-K', '10-Q', 'balance_sheet', 'credit_agreement', 'other'}
ALLOWED_PERIOD_TYPES = {'annual', 'quarterly'}
ALLOWED_REPORTING_UNITS = {'millions', 'thousands', 'whole_dollars'}
ALLOWED_STATEMENT_TYPES = {'income_statement', 'balance_sheet', 'cash_flow'}
ALLOWED_RISK_BANDS = {'Low', 'Moderate', 'Elevated', 'High'}
ALLOWED_METRIC_STATUS = {'Calculated', 'Incomplete'}
ALLOWED_METRIC_RISK = {'Low', 'High'}

logger = logging.getLogger('credit_analyzer')
logging.basicConfig(level=logging.INFO)


def _parse_numeric(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(',', '')
    text = text.replace('$', '')
    negative = False
    if text.startswith('(') and text.endswith(')'):
        negative = True
        text = text[1:-1].strip()
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@app.get('/api/actions/health-check', tags=['System'], summary='Health check (liveness)')
@app.get('/api/credit/system/health', tags=['Credit API'], summary='[Credit] Health check (same as /health)')
@app.get('/health', include_in_schema=False)
def health_check():
    return {'status': 'ok'}


# Uploads a PDF file, stores it in /uploads, runs financial extraction, saves document + extracted rows.
@app.post(
    '/api/credit/documents/upload-and-parse',
    response_model=schemas.DocumentRead,
    tags=['Credit API'],
    summary='[Credit] Upload PDF, parse statements, persist document + extracted rows',
)
@app.post(
    '/api/extract/upload-pdf-and-parse-financials',
    response_model=schemas.DocumentRead,
    tags=['Extraction'],
    summary='Upload PDF, parse financial statements, persist document and ExtractedFinancial rows',
)
@app.post('/api/documents/upload', response_model=schemas.DocumentRead, include_in_schema=False)
def upload_document(
    file: UploadFile = File(...),
    doc_type: str = Form('other'),
    company_name: str | None = Form(None),
    ticker: str | None = Form(None),
    fiscal_year: int | None = Form(None),
    period_type: str | None = Form(None),
    reporting_unit: str | None = Form(None),
    detected_unit: str | None = Form(None),
    db: Session = Depends(get_db),
):
    if file.content_type not in {'application/pdf', 'application/octet-stream'}:
        raise HTTPException(status_code=400, detail='Only PDF files are supported.')
    if doc_type not in ALLOWED_DOC_TYPES:
        raise HTTPException(status_code=400, detail='Invalid doc_type.')
    if period_type is not None and period_type not in ALLOWED_PERIOD_TYPES:
        raise HTTPException(status_code=400, detail='Invalid period_type.')
    if reporting_unit is not None and reporting_unit not in ALLOWED_REPORTING_UNITS:
        raise HTTPException(status_code=400, detail='Invalid reporting_unit.')

    safe_name = f"{uuid.uuid4().hex}_{Path(file.filename or 'document.pdf').name}"
    destination = UPLOAD_DIR / safe_name

    with destination.open('wb') as buffer:
        shutil.copyfileobj(file.file, buffer)

    file_size = float(destination.stat().st_size)
    file_url = f'/uploads/{safe_name}'

    record = models.Document(
        file_name=file.filename or safe_name,
        file_url=file_url,
        file_size=file_size,
        doc_type=doc_type,
        company_name=company_name,
        ticker=ticker,
        fiscal_year=fiscal_year,
        period_type=period_type,
        reporting_unit=reporting_unit,
        detected_unit=detected_unit,
    )
    db.add(record)
    db.commit()
    db.refresh(record)
    try:
        from financial_parser import parse_financial_statements
        extracted_rows = parse_financial_statements(str(destination))
        for row in extracted_rows:
            extracted = schemas.ExtractedFinancialCreate(**row)
            save_extracted_financials(record.id, extracted, db)
    except Exception as exc:
        logger.exception("Extraction failed for %s: %s", destination, exc)
    return record


# Saves one extracted financial statement row for a document.
@app.post(
    '/api/credit/documents/{document_id}/financial-rows',
    response_model=schemas.ExtractedFinancialRead,
    tags=['Credit API'],
    summary='[Credit] Insert one ExtractedFinancial row for a document',
)
@app.post(
    '/api/persist/extracted-financial-row/{document_id}',
    response_model=schemas.ExtractedFinancialRead,
    tags=['Persist'],
    summary='Insert one ExtractedFinancial row (income / balance / cash) for a document',
)
@app.post('/api/documents/{document_id}/financials', response_model=schemas.ExtractedFinancialRead, include_in_schema=False)
def save_extracted_financials(document_id: int, payload: schemas.ExtractedFinancialCreate, db: Session = Depends(get_db)):
    document = db.query(models.Document).filter(models.Document.id == document_id).first()
    if not document:
        raise HTTPException(status_code=404, detail='Document not found.')
    if payload.statement_type not in ALLOWED_STATEMENT_TYPES:
        raise HTTPException(status_code=400, detail='Invalid statement_type.')

    payload_data = payload.model_dump()
    extra_fields: Dict[str, Any] = {}
    if getattr(payload, "__pydantic_extra__", None):
        extra_fields.update(payload.__pydantic_extra__)
    raw_from_payload = payload_data.pop("raw_fields", None)
    if isinstance(raw_from_payload, dict):
        extra_fields.update(raw_from_payload)

    allowed_keys = {
        "statement_type",
        "revenue",
        "cogs",
        "ebitda",
        "ebit",
        "net_income",
        "interest_expense",
        "total_assets",
        "total_equity",
        "current_assets",
        "current_liabilities",
        "inventory",
        "st_debt",
        "lt_debt",
        "operating_cf",
        "capex",
        "da",
        "selected_year",
        "extraction_confidence",
        "raw_fields",
    }
    filtered = {k: v for k, v in payload_data.items() if k in allowed_keys}
    filtered["raw_fields"] = extra_fields or None
    record = models.ExtractedFinancial(document_id=document_id, **filtered)
    db.add(record)
    db.commit()
    db.refresh(record)
    return record


@app.get(
    '/api/credit/documents/{document_id}/financial-rows',
    tags=['Credit API'],
    summary='[Credit] List ExtractedFinancial rows for a document',
)
@app.get(
    '/api/fetch/extracted-financial-rows/{document_id}',
    tags=['Fetch'],
    summary='List all ExtractedFinancial rows for a document (newest fiscal year first)',
)
@app.get('/api/documents/{document_id}/financials', include_in_schema=False)
def get_document_financials(document_id: int, db: Session = Depends(get_db)):
    document = db.query(models.Document).filter(models.Document.id == document_id).first()
    if not document:
        raise HTTPException(status_code=404, detail='Document not found.')
    rows = (
        db.query(models.ExtractedFinancial)
        .filter(models.ExtractedFinancial.document_id == document_id)
        .order_by(models.ExtractedFinancial.selected_year.desc().nulls_last(), models.ExtractedFinancial.created_at.desc())
        .all()
    )
    output: List[Dict[str, Any]] = []
    for row in rows:
        base = schemas.ExtractedFinancialRead.model_validate(row).model_dump()
        raw = base.pop("raw_fields", None)
        merged = dict(base)
        if isinstance(raw, dict):
            for k, v in raw.items():
                if k not in merged:
                    merged[k] = v
        output.append(merged)
    return output


# Saves one credit analysis and all 12 metric score rows linked to that analysis.
@app.post(
    '/api/persist/credit-analysis-with-metrics',
    response_model=schemas.AnalysisWithMetrics,
    tags=['Persist'],
    summary='Create CreditAnalysis plus exactly 12 MetricScore rows',
)
@app.post('/api/analyses', response_model=schemas.AnalysisWithMetrics, include_in_schema=False)
def create_analysis(payload: schemas.CreditAnalysisCreate, db: Session = Depends(get_db)):
    document = db.query(models.Document).filter(models.Document.id == payload.document_id).first()
    if not document:
        raise HTTPException(status_code=404, detail='Document not found.')
    if payload.risk_band is not None and payload.risk_band not in ALLOWED_RISK_BANDS:
        raise HTTPException(status_code=400, detail='Invalid risk_band.')
    if len(payload.metric_scores) != 12:
        raise HTTPException(status_code=400, detail='Exactly 12 metric_scores are required.')

    analysis_data = payload.model_dump(exclude={'metric_scores'})
    analysis = models.CreditAnalysis(**analysis_data)
    db.add(analysis)
    db.flush()

    metric_rows = []
    for metric in payload.metric_scores:
        if metric.status is not None and metric.status not in ALLOWED_METRIC_STATUS:
            raise HTTPException(status_code=400, detail=f'Invalid metric status: {metric.status}')
        if metric.risk_level is not None and metric.risk_level not in ALLOWED_METRIC_RISK:
            raise HTTPException(status_code=400, detail=f'Invalid metric risk_level: {metric.risk_level}')
        metric_rows.append(models.MetricScore(analysis_id=analysis.id, **metric.model_dump()))

    db.add_all(metric_rows)
    db.commit()
    out = (
        db.query(models.CreditAnalysis)
        .options(selectinload(models.CreditAnalysis.metric_scores))
        .filter(models.CreditAnalysis.id == analysis.id)
        .first()
    )
    return out


# Returns all uploaded documents sorted by newest first.
@app.get(
    '/api/credit/documents',
    response_model=List[schemas.DocumentRead],
    tags=['Credit API'],
    summary='[Credit] List all uploaded documents (newest first)',
)
@app.get('/api/fetch/documents-list', response_model=List[schemas.DocumentRead], tags=['Fetch'], summary='List all uploaded documents (newest first)')
@app.get('/api/documents', response_model=List[schemas.DocumentRead], include_in_schema=False)
def list_documents(db: Session = Depends(get_db)):
    return db.query(models.Document).order_by(models.Document.uploaded_at.desc()).all()


# Returns one document and all extracted financial rows linked to it.
@app.get(
    '/api/fetch/document-with-financials/{document_id}',
    response_model=schemas.DocumentWithFinancials,
    tags=['Fetch'],
    summary='Get Document plus ExtractedFinancial rows (newest fiscal year first)',
)
@app.get('/api/documents/{document_id}', response_model=schemas.DocumentWithFinancials, include_in_schema=False)
def get_document(document_id: int, db: Session = Depends(get_db)):
    document = db.query(models.Document).filter(models.Document.id == document_id).first()
    if not document:
        raise HTTPException(status_code=404, detail='Document not found.')

    financials = (
        db.query(models.ExtractedFinancial)
        .filter(models.ExtractedFinancial.document_id == document_id)
        .order_by(models.ExtractedFinancial.selected_year.desc().nulls_last(), models.ExtractedFinancial.created_at.desc())
        .all()
    )
    return {'document': document, 'extracted_financials': financials}


# Returns one analysis and all of its metric score rows.
@app.get(
    '/api/fetch/credit-analysis/{analysis_id}',
    response_model=schemas.AnalysisWithMetrics,
    tags=['Fetch'],
    summary='Get one CreditAnalysis with MetricScore rows',
)
@app.get('/api/analyses/{analysis_id}', response_model=schemas.AnalysisWithMetrics, include_in_schema=False)
def get_analysis(analysis_id: int, db: Session = Depends(get_db)):
    analysis = (
        db.query(models.CreditAnalysis)
        .options(selectinload(models.CreditAnalysis.metric_scores))
        .filter(models.CreditAnalysis.id == analysis_id)
        .first()
    )
    if not analysis:
        raise HTTPException(status_code=404, detail='Analysis not found.')
    return analysis


# Returns analysis history for one document, including metric scores for each run.
@app.get(
    '/api/fetch/credit-analyses-for-document/{document_id}',
    response_model=List[schemas.AnalysisWithMetrics],
    tags=['Fetch'],
    summary='List CreditAnalysis runs for a document (newest first)',
)
@app.get('/api/documents/{document_id}/analyses', response_model=List[schemas.AnalysisWithMetrics], include_in_schema=False)
def get_document_analyses(document_id: int, db: Session = Depends(get_db)):
    document = db.query(models.Document).filter(models.Document.id == document_id).first()
    if not document:
        raise HTTPException(status_code=404, detail='Document not found.')

    return (
        db.query(models.CreditAnalysis)
        .options(selectinload(models.CreditAnalysis.metric_scores))
        .filter(models.CreditAnalysis.document_id == document_id)
        .order_by(models.CreditAnalysis.created_at.desc())
        .all()
    )


# --- Query-parameter aliases (easier to use in Swagger UI than path `{document_id}`) ---


@app.get(
    '/api/credit/query/financial-rows',
    tags=['Credit API'],
    summary='[Credit] List financial rows (?document_id=) — Swagger-friendly',
)
def get_document_financials_by_query(
    document_id: int = Query(
        ...,
        description='Primary key of the document row (`id` from upload or GET /api/credit/documents).',
        ge=1,
        examples=[1],
    ),
    db: Session = Depends(get_db),
):
    return get_document_financials(document_id, db)


@app.get(
    '/api/credit/query/document-with-financials',
    response_model=schemas.DocumentWithFinancials,
    tags=['Credit API'],
    summary='[Credit] Document + financial rows (?document_id=) — Swagger-friendly',
)
def get_document_by_query(
    document_id: int = Query(
        ...,
        description='Primary key of the document row (`id` from upload or GET /api/credit/documents).',
        ge=1,
        examples=[1],
    ),
    db: Session = Depends(get_db),
):
    return get_document(document_id, db)


@app.get(
    '/api/credit/query/credit-analyses-for-document',
    response_model=List[schemas.AnalysisWithMetrics],
    tags=['Credit API'],
    summary='[Credit] Credit analyses for document (?document_id=) — Swagger-friendly',
)
def get_document_analyses_by_query(
    document_id: int = Query(
        ...,
        description='Primary key of the document row (`id` from upload or GET /api/credit/documents).',
        ge=1,
        examples=[1],
    ),
    db: Session = Depends(get_db),
):
    return get_document_analyses(document_id, db)
