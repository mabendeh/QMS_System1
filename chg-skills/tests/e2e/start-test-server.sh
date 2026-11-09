#!/usr/bin/env bash
# Starts the app on port 3100 against a brand-new, empty database created for this test run only.
# Nothing existing is reset or deleted; tests/e2e/teardown.ts drops this run's database afterwards.
set -euo pipefail
cd "$(dirname "$0")/../.."
set -a; source .env; set +a
if [ -z "${TEST_DATABASE_URL:-}" ]; then echo "TEST_DATABASE_URL is not set" >&2; exit 1; fi
NAME="chg_e2e_$(date +%s)"
BASE="${TEST_DATABASE_URL%%\?*}"
ADMIN_URL="${BASE%/*}/postgres"
psql "$ADMIN_URL" -qc "CREATE DATABASE $NAME"
export DATABASE_URL="${BASE%/*}/$NAME?schema=public"
mkdir -p var
echo "$DATABASE_URL" > var/e2e-db-url
export UPLOAD_DIR="./var/test-uploads"
mkdir -p "$UPLOAD_DIR"
npx prisma migrate deploy >/dev/null
npx tsx prisma/seed.ts >/dev/null
[ -d .next ] || npx next build >/dev/null
exec npx next start -p 3100
