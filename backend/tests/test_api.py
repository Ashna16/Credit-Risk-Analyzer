"""API smoke tests using SQLite (:memory:). Run from backend/: pytest."""

from fastapi.testclient import TestClient


def test_health(client: TestClient) -> None:
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json().get("status") == "ok"
    rc = client.get("/api/credit/system/health")
    assert rc.status_code == 200


def test_upload_persist_and_fetch_financials(client: TestClient) -> None:
    files = {
        "file": (
            "minimal.pdf",
            b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n",
            "application/pdf",
        )
    }
    r = client.post(
        "/api/extract/upload-pdf-and-parse-financials",
        files=files,
        data={"doc_type": "other", "company_name": "TestCo"},
    )
    assert r.status_code == 200
    doc_id = r.json()["id"]

    payload = {
        "statement_type": "income_statement",
        "revenue": 1_000_000.0,
        "net_income": 100_000.0,
        "ebit": 150_000.0,
    }
    r2 = client.post(f"/api/persist/extracted-financial-row/{doc_id}", json=payload)
    assert r2.status_code == 200
    assert r2.json()["revenue"] == 1_000_000.0

    r3 = client.get(f"/api/fetch/extracted-financial-rows/{doc_id}")
    assert r3.status_code == 200
    rows = r3.json()
    assert isinstance(rows, list)
    assert len(rows) >= 1
    assert any(row.get("revenue") == 1_000_000.0 for row in rows)


def test_list_documents(client: TestClient) -> None:
    r = client.get("/api/fetch/documents-list")
    assert r.status_code == 200
    assert isinstance(r.json(), list)


def test_create_and_get_credit_analysis(client: TestClient) -> None:
    files = {
        "file": (
            "minimal2.pdf",
            b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n",
            "application/pdf",
        )
    }
    r = client.post(
        "/api/extract/upload-pdf-and-parse-financials",
        files=files,
        data={"doc_type": "other"},
    )
    assert r.status_code == 200
    doc_id = r.json()["id"]

    metric_scores = [
        {
            "metric_name": f"m{i}",
            "calculated_value": float(i),
            "industry_threshold": 1.0,
            "base_score": 1.0,
            "adjusted_score": 1.0,
            "status": "Calculated",
            "risk_level": "Low",
        }
        for i in range(12)
    ]
    body = {
        "document_id": doc_id,
        "industry": "Technology",
        "risk_band": "Low",
        "metric_scores": metric_scores,
    }
    r2 = client.post("/api/persist/credit-analysis-with-metrics", json=body)
    assert r2.status_code == 200
    out = r2.json()
    assert out["document_id"] == doc_id
    assert len(out["metric_scores"]) == 12

    aid = out["id"]
    r3 = client.get(f"/api/fetch/credit-analysis/{aid}")
    assert r3.status_code == 200
    assert len(r3.json()["metric_scores"]) == 12

    r4 = client.get(f"/api/fetch/credit-analyses-for-document/{doc_id}")
    assert r4.status_code == 200
    assert len(r4.json()) >= 1


def test_credit_api_paths_upload_persist_fetch(client: TestClient) -> None:
    """Same behavior as legacy routes via /api/credit/... aliases."""
    files = {
        "file": (
            "credit_alias.pdf",
            b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n",
            "application/pdf",
        )
    }
    r = client.post(
        "/api/credit/documents/upload-and-parse",
        files=files,
        data={"doc_type": "other", "company_name": "AliasCo"},
    )
    assert r.status_code == 200
    doc_id = r.json()["id"]
    r2 = client.post(
        f"/api/credit/documents/{doc_id}/financial-rows",
        json={"statement_type": "income_statement", "revenue": 500.0},
    )
    assert r2.status_code == 200
    r3 = client.get(f"/api/credit/documents/{doc_id}/financial-rows")
    assert r3.status_code == 200
    assert any(row.get("revenue") == 500.0 for row in r3.json())
    r3q = client.get("/api/credit/query/financial-rows", params={"document_id": doc_id})
    assert r3q.status_code == 200
    assert r3q.json() == r3.json()
    r4 = client.get("/api/credit/documents")
    assert r4.status_code == 200
    assert any(d.get("company_name") == "AliasCo" for d in r4.json())
