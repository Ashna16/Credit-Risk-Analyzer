# Credit Risk Analyzer

**Repository:** [github.com/Ashna16/Credit-Risk-Analyzer](https://github.com/Ashna16/Credit-Risk-Analyzer)

A **credit document intelligence** workspace: upload PDFs (loan agreements, covenants, 10‑K/10‑Q–style filings), **extract** structured fields with optional AI (Google Gemini) and rules-based fallback, **visualize** financials, and run a **credit risk** dashboard. A **FastAPI** backend persists documents, extracted financial rows, and credit analyses to **PostgreSQL** (or SQLite for local dev).

---

## Security (read this first)

- **Do not commit** `.env` files, API keys, PEM keys, or `credentials.json`. They are listed in `.gitignore`.
- Copy **`loan_doc_extractor_mvp/.env.example`** and **`backend/.env.example`** to `.env` locally and fill in real values **only on your machine**.
- **`GEMINI_API_KEY`** and **`DATABASE_URL`** are examples in docs only—never paste live secrets into the repo.

---

## What’s in this repo

| Piece | Role |
|--------|------|
| **`loan_doc_extractor_mvp/`** | Streamlit UI: upload, extraction, data viz, credit risk, exports |
| **`backend/`** | FastAPI app: upload/parse, persist financial rows & credit analyses, Swagger at `/docs` |
| **`extractor.py`** | PDF → Gemini + regex/schema extraction (optional Gemini) |
| **`backend/financial_parser.py`** | PDF financial statement parsing (uses `pdfplumber`) |

---

## Quick start (UI)

```bash
cd loan_doc_extractor_mvp
python3 -m pip install -r requirements.txt
cp .env.example .env   # add GEMINI_API_KEY if you use AI extraction
streamlit run app.py --server.port 9000 --server.address localhost
```

Open **http://localhost:9000** (or omit `--server.port` to use Streamlit’s default **8501**).

---

## Backend API (read vs write)

**Base URL (local):** `http://127.0.0.1:8000` — interactive docs: **http://127.0.0.1:8000/docs**

| Action | Method | Path (examples) |
|--------|--------|------------------|
| Upload PDF & parse on server | **POST** | `/api/credit/documents/upload-and-parse` |
| Insert one financial row | **POST** | `/api/credit/documents/{document_id}/financial-rows` |
| Save credit analysis + metrics | **POST** | `/api/persist/credit-analysis-with-metrics` |
| List documents | **GET** | `/api/credit/documents` |
| Read extracted financial rows | **GET** | `/api/credit/documents/{document_id}/financial-rows` or `/api/credit/query/financial-rows?document_id=` |
| Read document + rows | **GET** | `/api/credit/query/document-with-financials?document_id=` |
| Read credit analyses for a document | **GET** | `/api/credit/query/credit-analyses-for-document?document_id=` |

Use **GET** to **read** what is stored; **POST** uploads or creates rows.

---

## Backend dev

```bash
cd backend
python3 -m pip install -r requirements.txt
cp .env.example .env   # set DATABASE_URL, etc.
uvicorn main:app --reload --port 8000
```

Point the Streamlit app at the API with **`BACKEND_BASE_URL=http://localhost:8000`** (default in code).

---

## Tests

```bash
# MVP tests
PYTHONPATH=. python3 -m pytest loan_doc_extractor_mvp/tests/ -v

# Backend tests
python3 -m pytest backend/tests/ -v
```

---

## Tech stack (summary)

- Python 3, Streamlit, FastAPI, SQLAlchemy, Pandas  
- PDF: PyMuPDF, pdfplumber (financial parser path), optional Gemini REST API  
- Charts: Altair / Streamlit native charts in the Data Visualization module  

---

## What’s in version control

This repository is intentionally **source-only**: `backend/` (API + parser + tests) and `loan_doc_extractor_mvp/` (Streamlit UI, extractor, tests), plus root `README.md`, `.gitignore`, and `*.env.example`. It does **not** include uploaded PDFs, manual test corpora, portfolio diagrams, or local run logs—those stay on your machine (see `.gitignore`).

## Project layout (abbreviated)

```
Credit-Risk-Analyzer/
├── backend/                 # FastAPI + financial_parser + DB models + scripts + tests
├── loan_doc_extractor_mvp/  # Streamlit app, extractor, tests, .env.example
├── README.md
└── .gitignore               # .env, uploads, keys, caches, assets
```

---

## License & contact

Portfolio / educational project. For licensing or collaboration, contact the repository owner (**[@Ashna16](https://github.com/Ashna16)**).
