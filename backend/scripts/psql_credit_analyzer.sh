#!/usr/bin/env bash
# Opens interactive psql inside the Docker Postgres container (credit_analyzer database).

set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! docker exec credit_analyzer_db pg_isready -U postgres >/dev/null 2>&1; then
  echo "Start the DB first: cd \"$ROOT\" && docker compose up -d"
  exit 1
fi

exec docker exec -it credit_analyzer_db psql -U postgres -d credit_analyzer
