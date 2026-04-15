#!/usr/bin/env bash
# Opens TablePlus with the Docker credit_analyzer database (default host port PGPORT from .env, usually 5433).
# Requires: Docker Postgres container running, TablePlus installed.

set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
if [ -f "$ROOT/.env" ]; then
  # shellcheck disable=SC1090
  set -a && . "$ROOT/.env" && set +a
fi

USER="${POSTGRES_USER:-postgres}"
PASS="${POSTGRES_PASSWORD:-change_me}"
DB="${POSTGRES_DB:-credit_analyzer}"
HOST="${PGHOST:-127.0.0.1}"
PORT="${PGPORT:-5433}"

# TablePlus expects a postgres:// or postgresql:// URL (password URL-encoded if needed)
URL="postgresql://${USER}:${PASS}@${HOST}:${PORT}/${DB}"

if ! command -v open >/dev/null 2>&1; then
  echo "Paste this into TablePlus → Import from URL:"
  echo "$URL"
  exit 0
fi

echo "Opening TablePlus → ${DB} @ ${HOST}:${PORT}"
open -a TablePlus "$URL"
