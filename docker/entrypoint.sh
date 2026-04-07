#!/bin/sh
set -eu

echo "Starting Sardoba bot container..."

# Best-effort wait for PostgreSQL (avoid race on first start).
python - <<'PY'
import os, sys, time
import psycopg2

database_url = os.getenv("DATABASE_URL", "")
if database_url.startswith("postgresql+asyncpg://"):
    database_url = "postgresql://" + database_url.split("://", 1)[1]

deadline = time.time() + 60
last_err = None
while time.time() < deadline:
    try:
        if database_url:
            conn = psycopg2.connect(database_url)
        else:
            conn = psycopg2.connect(
                host=os.getenv("DB_HOST", "db"),
                port=int(os.getenv("DB_PORT", "5432")),
                user=os.getenv("DB_USER", "sardoba"),
                password=os.getenv("DB_PASSWORD", ""),
                dbname=os.getenv("DB_NAME", "sardoba_bot"),
            )
        conn.close()
        print("PostgreSQL is ready.")
        break
    except Exception as e:
        last_err = e
        time.sleep(2)
else:
    print(f"PostgreSQL is not ready after 60s: {last_err}")
    sys.exit(1)
PY

exec python -m sardoba_bot
