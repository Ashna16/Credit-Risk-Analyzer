# Credit Risk Analyzer

![Credit Risk Analyzer — financial document automation](assets/readme-hero.png)

Every month or quarter, after a bank lends money to a company, a team of analysts manually pulls their financial documents — 10-Ks, income statements, balance sheets, cash flow statements, covenant reports — opens each one, copies the numbers into Excel by hand, and calculates whether that borrower is still safe.

For a portfolio of hundreds of clients, that's days of work every single cycle. One missed number means a missed risk. One missed risk means a default nobody saw coming. 📄

I built **Credit Risk Analyzer** to automate the entire flow.

Upload any financial document → AI extracts every field automatically → assigns a **confidence score** to each extracted value so you know exactly what to trust → calculates **12 credit risk ratios** across liquidity, leverage, coverage, profitability and cash flow → flags covenant breaches → outputs a credit decision with a **policy-approved loan limit**.

The confidence scoring is the part I'm most proud of. A bank can't act on data it can't trust. Every single extracted field tells you how confident the system is — so analysts aren't flying blind, they're verifying the right things.

**Built with:** #FastAPI #Python #pdfplumber #SQLAlchemy #Streamlit #Docker #OpenAICodex #Cursor #TablePlus #Figma #Swagger #Product #BusinessAnalyst #PM #QualityAnalyst #Banks #Automation

---

**Repository:** [github.com/Ashna16/Credit-Risk-Analyzer](https://github.com/Ashna16/Credit-Risk-Analyzer)

---

## Developers

### Security

- Do **not** commit `.env` files or API keys (see `.gitignore`). Copy `loan_doc_extractor_mvp/.env.example` and `backend/.env.example` to `.env` locally.

### Quick start (UI)

```bash
cd loan_doc_extractor_mvp
python3 -m pip install -r requirements.txt
cp .env.example .env   # optional: GEMINI_API_KEY for AI extraction
streamlit run app.py --server.port 9000 --server.address localhost
```

Open **http://localhost:9000**.

### Backend API

```bash
cd backend && python3 -m pip install -r requirements.txt && cp .env.example .env
uvicorn main:app --reload --port 8000
```

Swagger: **http://127.0.0.1:8000/docs** — `GET` endpoints read stored data; `POST /api/credit/documents/upload-and-parse` uploads and parses; `POST /api/persist/credit-analysis-with-metrics` saves credit analysis + metrics.

### Tests

```bash
PYTHONPATH=. python3 -m pytest loan_doc_extractor_mvp/tests/ -v
python3 -m pytest backend/tests/ -v
```

---

## License & contact

Portfolio / educational project. **[Ashna16](https://github.com/Ashna16)**
