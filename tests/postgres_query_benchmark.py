import asyncio
import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
import sys
from zoneinfo import ZoneInfo

import psycopg2

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from db_config import DatabaseConnection


DEFAULT_DB_URL = "postgresql://sardoba:sardoba@127.0.0.1:55432/sardoba_bench"
SCHEMA_PATH = Path(__file__).resolve().parents[1] / "sql" / "postgres" / "schema.sql"


def connect():
    return psycopg2.connect(os.getenv("POSTGRES_BENCH_URL", DEFAULT_DB_URL))


def dt(value):
    return datetime.fromisoformat(value).replace(tzinfo=ZoneInfo("Asia/Tashkent"))


def reset_schema(conn):
    schema_sql = SCHEMA_PATH.read_text(encoding="utf-8-sig")
    with conn.cursor() as cur:
        cur.execute("DROP SCHEMA IF EXISTS public CASCADE;")
        cur.execute("CREATE SCHEMA public;")
        cur.execute(schema_sql)
    conn.commit()


def seed_data(conn, cashier_count=200, shifts_per_cashier=60):
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO users (
                telegram_id,
                first_name,
                last_name,
                phone_number,
                role,
                password_hash,
                is_active
            )
            SELECT
                900000 + gs,
                'Cashier' || gs,
                'Bench',
                '+99890' || LPAD(gs::text, 7, '0'),
                'cashier',
                'hash',
                TRUE
            FROM generate_series(1, %s) AS gs
            """,
            (cashier_count,),
        )
        cur.execute(
            """
            INSERT INTO approval_requests (
                telegram_id,
                first_name,
                last_name,
                phone_number,
                role,
                password_hash,
                status,
                requested_at
            )
            SELECT
                800000 + gs,
                'Req' || gs,
                'Bench',
                '+99891' || LPAD(gs::text, 7, '0'),
                'cashier',
                'hash',
                CASE WHEN gs % 3 = 0 THEN 'approved' ELSE 'pending' END,
                TIMESTAMPTZ '2026-03-01 08:00:00+05' + (gs * interval '10 minutes')
            FROM generate_series(1, 5000) AS gs
            """
        )
        cur.execute(
            """
            INSERT INTO shifts (
                user_id,
                location_id,
                opened_at,
                closed_at,
                opening_amount,
                closing_amount,
                is_open
            )
            SELECT
                u.id,
                ((g.shift_no - 1) %% 4) + 1,
                TIMESTAMPTZ '2026-03-01 06:00:00+05'
                    + ((g.shift_no - 1) * interval '6 hours')
                    + ((u.id %% 7) * interval '15 minutes'),
                TIMESTAMPTZ '2026-03-01 16:00:00+05'
                    + ((g.shift_no - 1) * interval '6 hours')
                    + ((u.id %% 7) * interval '15 minutes'),
                50000,
                150000,
                FALSE
            FROM users u
            CROSS JOIN generate_series(1, %s) AS g(shift_no)
            WHERE u.role = 'cashier'
            """,
            (shifts_per_cashier,),
        )
        cur.execute(
            """
            INSERT INTO reports (
                shift_id,
                report_type,
                sales_amount,
                debt_received,
                expenses,
                uzcard_amount,
                humo_amount,
                p2p_amount,
                uzcard_refund,
                humo_refund,
                other_payments,
                debt_payments,
                debt_refunds
            )
            SELECT
                s.id,
                'daily_report',
                100000 + (s.id % 100) * 10000,
                (s.id % 10) * 1000,
                (s.id % 7) * 700,
                (s.id % 5) * 2500,
                (s.id % 4) * 2000,
                (s.id % 3) * 1500,
                (s.id % 3) * 500,
                (s.id % 2) * 400,
                (s.id % 6) * 600,
                (s.id % 5) * 300,
                (s.id % 4) * 200
            FROM shifts s
            """
        )
        cur.execute(
            """
            INSERT INTO reports (
                shift_id,
                report_type,
                sales_amount,
                debt_received,
                expenses,
                uzcard_amount,
                humo_amount,
                p2p_amount,
                uzcard_refund,
                humo_refund,
                other_payments,
                debt_payments,
                debt_refunds
            )
            SELECT
                s.id,
                'daily_report',
                120000 + (s.id % 100) * 12000,
                (s.id % 10) * 1200,
                (s.id % 7) * 750,
                (s.id % 5) * 2700,
                (s.id % 4) * 2200,
                (s.id % 3) * 1600,
                (s.id % 3) * 520,
                (s.id % 2) * 420,
                (s.id % 6) * 650,
                (s.id % 5) * 320,
                (s.id % 4) * 240
            FROM shifts s
            WHERE s.id % 3 = 0
            """
        )
        cur.execute("ANALYZE;")
    conn.commit()


def benchmark_query(conn, query, params, iterations=25):
    durations = []
    row_count = 0
    with conn.cursor() as cur:
        for _ in range(iterations):
            started_at = time.perf_counter()
            cur.execute(query, params)
            rows = cur.fetchall()
            durations.append((time.perf_counter() - started_at) * 1000)
            row_count = len(rows)
    return {
        "avg_ms": round(sum(durations) / len(durations), 3),
        "min_ms": round(min(durations), 3),
        "max_ms": round(max(durations), 3),
        "rows": row_count,
    }


def explain_query(conn, query, params):
    with conn.cursor() as cur:
        cur.execute(f"EXPLAIN (ANALYZE, BUFFERS) {query}", params)
        lines = [row[0] for row in cur.fetchall()]

    execution_time_ms = None
    planning_time_ms = None
    for line in lines:
        execution_match = re.search(r"Execution Time: ([0-9.]+) ms", line)
        if execution_match:
            execution_time_ms = float(execution_match.group(1))
        planning_match = re.search(r"Planning Time: ([0-9.]+) ms", line)
        if planning_match:
            planning_time_ms = float(planning_match.group(1))

    return {
        "planning_ms": round(planning_time_ms or 0.0, 3),
        "execution_ms": round(execution_time_ms or 0.0, 3),
    }


async def benchmark_runtime_db_query(db, query, params, mode="one", iterations=25):
    durations = []
    row_count = 0
    for _ in range(iterations):
        started_at = time.perf_counter()
        if mode == "all":
            rows = await db.fetch_all(query, params)
            row_count = len(rows)
        else:
            row = await db.fetch_one(query, params)
            row_count = 1 if row else 0
        durations.append((time.perf_counter() - started_at) * 1000)
    return {
        "avg_ms": round(sum(durations) / len(durations), 3),
        "min_ms": round(min(durations), 3),
        "max_ms": round(max(durations), 3),
        "rows": row_count,
    }


async def benchmark_runtime_db(database_url):
    db = DatabaseConnection()
    db.database_url = database_url
    await db.connect()
    try:
        today_shift_query = """
            SELECT id, is_open, opened_at, closed_at
            FROM shifts
            WHERE user_id = %s AND opened_at >= %s AND opened_at < %s
            ORDER BY opened_at DESC, id DESC
            LIMIT 1
        """
        latest_report_query = """
            SELECT id
            FROM reports
            WHERE shift_id = %s AND report_type = 'daily_report'
            ORDER BY id DESC
            LIMIT 1
        """
        export_range_query = """
            SELECT
                s.id,
                u.first_name,
                u.last_name,
                l.name AS location,
                s.opened_at,
                s.closed_at,
                COALESCE(r.sales_amount, 0) AS sales_amount,
                COALESCE(r.debt_received, 0) AS debt_received,
                COALESCE(r.expenses, 0) AS expenses,
                COALESCE(r.uzcard_amount, 0) AS uzcard_amount,
                COALESCE(r.humo_amount, 0) AS humo_amount,
                COALESCE(r.p2p_amount, 0) AS p2p_amount,
                COALESCE(r.uzcard_refund, 0) AS uzcard_refund,
                COALESCE(r.humo_refund, 0) AS humo_refund,
                COALESCE(r.other_payments, 0) AS other_payments,
                COALESCE(r.debt_payments, 0) AS debt_payments,
                COALESCE(r.debt_refunds, 0) AS debt_refunds
            FROM shifts s
            JOIN users u ON s.user_id = u.id
            JOIN locations l ON s.location_id = l.id
            JOIN LATERAL (
                SELECT
                    sales_amount,
                    debt_received,
                    expenses,
                    uzcard_amount,
                    humo_amount,
                    p2p_amount,
                    uzcard_refund,
                    humo_refund,
                    other_payments,
                    debt_payments,
                    debt_refunds
                FROM reports
                WHERE shift_id = s.id AND report_type = 'daily_report'
                ORDER BY id DESC
                LIMIT 1
            ) r ON TRUE
            WHERE s.opened_at >= %s AND s.opened_at < %s
            ORDER BY s.opened_at DESC
        """
        approval_query = """
            SELECT id
            FROM approval_requests
            WHERE telegram_id = %s AND status = %s
            ORDER BY requested_at DESC
            LIMIT 1
        """
        return {
            "today_shift_query": await benchmark_runtime_db_query(
                db,
                today_shift_query,
                (10, dt("2026-03-15T00:00:00"), dt("2026-03-16T00:00:00")),
            ),
            "latest_daily_report_query": await benchmark_runtime_db_query(
                db,
                latest_report_query,
                (9000,),
            ),
            "export_range_query": await benchmark_runtime_db_query(
                db,
                export_range_query,
                (dt("2026-03-10T00:00:00"), dt("2026-03-17T00:00:00")),
                mode="all",
                iterations=10,
            ),
            "approval_lookup_query": await benchmark_runtime_db_query(
                db,
                approval_query,
                (800120, "approved"),
            ),
        }
    finally:
        await db.disconnect()


def main():
    conn = connect()
    try:
        reset_schema(conn)
        seed_data(conn)

        today_shift_query = """
            SELECT id, is_open, opened_at, closed_at
            FROM shifts
            WHERE user_id = %s AND opened_at >= %s AND opened_at < %s
            ORDER BY opened_at DESC, id DESC
            LIMIT 1
        """
        latest_report_query = """
            SELECT id
            FROM reports
            WHERE shift_id = %s AND report_type = 'daily_report'
            ORDER BY id DESC
            LIMIT 1
        """
        export_range_query = """
            SELECT
                s.id,
                u.first_name,
                u.last_name,
                l.name AS location,
                s.opened_at,
                s.closed_at,
                COALESCE(r.sales_amount, 0) AS sales_amount,
                COALESCE(r.debt_received, 0) AS debt_received,
                COALESCE(r.expenses, 0) AS expenses,
                COALESCE(r.uzcard_amount, 0) AS uzcard_amount,
                COALESCE(r.humo_amount, 0) AS humo_amount,
                COALESCE(r.p2p_amount, 0) AS p2p_amount,
                COALESCE(r.uzcard_refund, 0) AS uzcard_refund,
                COALESCE(r.humo_refund, 0) AS humo_refund,
                COALESCE(r.other_payments, 0) AS other_payments,
                COALESCE(r.debt_payments, 0) AS debt_payments,
                COALESCE(r.debt_refunds, 0) AS debt_refunds
            FROM shifts s
            JOIN users u ON s.user_id = u.id
            JOIN locations l ON s.location_id = l.id
            JOIN LATERAL (
                SELECT
                    sales_amount,
                    debt_received,
                    expenses,
                    uzcard_amount,
                    humo_amount,
                    p2p_amount,
                    uzcard_refund,
                    humo_refund,
                    other_payments,
                    debt_payments,
                    debt_refunds
                FROM reports
                WHERE shift_id = s.id AND report_type = 'daily_report'
                ORDER BY id DESC
                LIMIT 1
            ) r ON TRUE
            WHERE s.opened_at >= %s AND s.opened_at < %s
            ORDER BY s.opened_at DESC
        """
        approval_query = """
            SELECT id
            FROM approval_requests
            WHERE telegram_id = %s AND status = %s
            ORDER BY requested_at DESC
            LIMIT 1
        """

        result = {
            "dataset": {
                "cashiers": 200,
                "shifts": 12000,
                "daily_reports": 16000,
                "approval_requests": 5000,
            },
            "today_shift_query": benchmark_query(
                conn,
                today_shift_query,
                (10, dt("2026-03-15T00:00:00"), dt("2026-03-16T00:00:00")),
            ),
            "latest_daily_report_query": benchmark_query(conn, latest_report_query, (9000,)),
            "export_range_query": benchmark_query(
                conn,
                export_range_query,
                (dt("2026-03-10T00:00:00"), dt("2026-03-17T00:00:00")),
                iterations=10,
            ),
            "approval_lookup_query": benchmark_query(
                conn,
                approval_query,
                (800120, "approved"),
            ),
        }

        result["explain"] = {
            "today_shift_query": explain_query(
                conn,
                today_shift_query,
                (10, dt("2026-03-15T00:00:00"), dt("2026-03-16T00:00:00")),
            ),
            "latest_daily_report_query": explain_query(conn, latest_report_query, (9000,)),
            "export_range_query": explain_query(
                conn,
                export_range_query,
                (dt("2026-03-10T00:00:00"), dt("2026-03-17T00:00:00")),
            ),
        }

        result["runtime_asyncpg"] = asyncio.run(
            benchmark_runtime_db(os.getenv("POSTGRES_BENCH_URL", DEFAULT_DB_URL))
        )

        print(json.dumps(result, indent=2))
    finally:
        conn.close()


if __name__ == "__main__":
    main()
