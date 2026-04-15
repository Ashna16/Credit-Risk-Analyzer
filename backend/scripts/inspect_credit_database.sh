#!/usr/bin/env bash
# Lists tables and row counts for the Credit Analyzer DB (Docker Postgres by default).
#
# Postgres.app: use "Connect" → host 127.0.0.1, port 5432, user postgres,
# password from backend/.env (POSTGRES_PASSWORD), database credit_analyzer.
# If Postgres.app and Docker both use port 5432, only one can run—stop one or map Docker to 5433.

set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! docker info >/dev/null 2>&1; then
  echo "Docker is not running. Start Docker Desktop, then: cd backend && docker compose up -d"
  exit 1
fi

if ! docker exec credit_analyzer_db pg_isready -U postgres >/dev/null 2>&1; then
  echo "Container credit_analyzer_db is not up. Run: cd \"$ROOT\" && docker compose up -d"
  exit 1
fi

echo "=== credit_analyzer @ credit_analyzer_db (Docker) ==="
docker exec credit_analyzer_db psql -U postgres -d credit_analyzer -c "\dt"
echo ""
echo "=== row counts ==="
docker exec credit_analyzer_db psql -U postgres -d credit_analyzer -c \
  "SELECT 'documents' AS tbl, COUNT(*)::int AS n FROM documents
   UNION ALL SELECT 'extracted_financials', COUNT(*)::int FROM extracted_financials
   UNION ALL SELECT 'credit_analyses', COUNT(*)::int FROM credit_analyses
   UNION ALL SELECT 'metric_scores', COUNT(*)::int FROM metric_scores;"
echo ""
echo "=== sample: latest documents ==="
docker exec credit_analyzer_db psql -U postgres -d credit_analyzer -c \
  "SELECT id, file_name, company_name, uploaded_at FROM documents ORDER BY id DESC LIMIT 5;"
