import asyncio
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from bot import SardobaBot
from export_utils import ExportUtils
from utils import hash_password, verify_password


class FakeDB:
    def __init__(self, rows):
        self.rows = rows

    async def fetch_all(self, query, params=None):
        return self.rows

    async def fetch_one(self, query, params=None):
        return None

    async def disconnect(self):
        return None


def avg_ms(fn, iterations=1):
    started = time.perf_counter()
    for _ in range(iterations):
        fn()
    return round(((time.perf_counter() - started) * 1000) / iterations, 3)


async def avg_ms_async(fn, iterations=1):
    started = time.perf_counter()
    for _ in range(iterations):
        await fn()
    return round(((time.perf_counter() - started) * 1000) / iterations, 3)


def build_bot():
    bot = SardobaBot.__new__(SardobaBot)
    bot.db = None
    bot.export_utils = None
    bot._group_chat_id_cache = None
    bot._locations_cache = None
    return bot


def build_export_utils(rows):
    export_utils = ExportUtils.__new__(ExportUtils)
    export_utils.db = FakeDB(rows)
    export_utils._owns_db = False
    return export_utils


async def main():
    hashed = hash_password("secret123")
    sample_daily_rows = [
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-04-01 09:00:00",
            "closed_at": "2026-04-01 22:00:00",
            "sales_amount": 1000,
            "debt_received": 100,
            "expenses": 50,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "total_balance": 1050,
        }
        for _ in range(100)
    ]
    sample_pdf_rows = [
        {
            "cashier_name": "Ali Valiyev",
            "location": "Sardoba",
            "date": "2026-04-01",
            "open_time": "09:00:00",
            "sales_amount": 1000,
            "debt_received": 100,
            "expenses": 50,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "total_balance": 1050,
        }
        for _ in range(50)
    ]

    bot = build_bot()
    export_excel = build_export_utils(sample_daily_rows)
    export_pdf = build_export_utils(sample_pdf_rows)

    metrics = {
        "hash_password_avg_ms": avg_ms(lambda: hash_password("secret123"), iterations=3),
        "verify_password_avg_ms": avg_ms(lambda: verify_password(hashed, "secret123"), iterations=10),
        "build_shift_summary_avg_ms": avg_ms(
            lambda: bot._build_shift_summary_message(
                {
                    "first_name": "Nilufar",
                    "last_name": "G'afurova",
                    "location": "Sardoba (Severniy)",
                    "opened_at": "2026-03-28 09:27:21",
                    "closed_at": "2026-03-28 09:49:16",
                    "opening_amount": 3000,
                    "closing_amount": 100000,
                    "sales_amount": 10_000_000,
                    "debt_received": 0,
                    "expenses": 200_000,
                    "uzcard_amount": 0,
                    "humo_amount": 0,
                    "uzcard_refund": 0,
                    "humo_refund": 0,
                    "other_payments": 0,
                    "debt_payments": 0,
                    "debt_refunds": 0,
                }
            ),
            iterations=2000,
        ),
        "generate_excel_report_avg_ms": await avg_ms_async(
            lambda: export_excel.generate_excel_report("daily"),
            iterations=1,
        ),
        "generate_pdf_report_avg_ms": await avg_ms_async(
            lambda: export_pdf.generate_pdf_report("daily"),
            iterations=1,
        ),
    }
    print(json.dumps(metrics, indent=2))


if __name__ == "__main__":
    asyncio.run(main())
