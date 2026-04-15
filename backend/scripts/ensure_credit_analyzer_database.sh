#!/usr/bin/env bash
# Creates database "credit_analyzer" inside the Docker Postgres container if missing (rare).
set -euo pipefail
if ! docker exec credit_analyzer_db pg_isready -U postgres >/dev/null 2>&1; then
  echo "Start Docker first: cd \"$(cd "$(dirname "$0")/.." && pwd)\" && docker compose up -d"
  exit 1
fi
EXISTS=$(docker exec credit_analyzer_db psql -U postgres -d postgres -Atc "SELECT 1 FROM pg_database WHERE datname='credit_analyzer';")
if [ "$EXISTS" = "1" ]; then
  echo "OK: database credit_analyzer already exists in Docker."
else
  echo "Creating database credit_analyzer in Docker..."
  docker exec credit_analyzer_db psql -U postgres -d postgres -c "CREATE DATABASE credit_analyzer;"
  echo "Done. Start the API once (docker compose) so SQLAlchemy creates tables."
fi
