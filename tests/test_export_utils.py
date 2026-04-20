from io import BytesIO

import pytest
from openpyxl import load_workbook

from export_utils import ExportUtils


class FakeDB:
    def __init__(self, rows):
        self.rows = list(rows)
        self.calls = []

    async def fetch_all(self, query, params=None):
        self.calls.append((query, params))
        return list(self.rows)

    async def disconnect(self):
        return None


def make_export_utils(rows):
    export_utils = ExportUtils.__new__(ExportUtils)
    export_utils.db = FakeDB(rows)
    export_utils._owns_db = False
    return export_utils


def test_datetime_bounds():
    export_utils = make_export_utils([])
    start, end = export_utils._datetime_bounds("2026-04-01", "2026-04-03")
    assert start.isoformat() == "2026-04-01T00:00:00+05:00"
    assert end.isoformat() == "2026-04-04T00:00:00+05:00"


@pytest.mark.asyncio
async def test_fetch_excel_dataset_daily():
    rows = [
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
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "total_balance": 1050,
        }
    ]
    export_utils = make_export_utils(rows)

    headers, values = await export_utils._fetch_excel_dataset("daily", "2026-04-01", "2026-04-01")

    assert headers[0] == "Ism"
    assert values[0][0] == "Ali"
    query, params = export_utils.db.calls[0]
    assert "JOIN LATERAL" in query
    assert params[0].isoformat() == "2026-04-01T00:00:00+05:00"
    assert params[1].isoformat() == "2026-04-02T00:00:00+05:00"


@pytest.mark.asyncio
async def test_generate_excel_report():
    rows = [
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
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "total_balance": 1050,
        }
    ]
    export_utils = make_export_utils(rows)

    file_obj = await export_utils.generate_excel_report("daily")
    workbook = load_workbook(filename=BytesIO(file_obj.getvalue()))
    worksheet = workbook.active

    assert worksheet["A1"].value == "Ism"
    assert worksheet["A2"].value == "Ali"


@pytest.mark.asyncio
async def test_generate_pdf_report():
    rows = [
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
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "total_balance": 1050,
        }
    ]
    export_utils = make_export_utils(rows)

    pdf = await export_utils.generate_pdf_report("daily")

    assert pdf.getvalue().startswith(b"%PDF")
    assert len(pdf.getvalue()) > 100


@pytest.mark.asyncio
async def test_close_connection_delegates():
    export_utils = make_export_utils([])
    await export_utils.close_connection()
    assert True
