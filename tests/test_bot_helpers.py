import json
from collections import defaultdict
from datetime import timedelta
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import AsyncMock
from unittest.mock import Mock
from zipfile import ZipFile

import pytest
from PIL import Image as PILImage
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import ConversationHandler

from bot import (
    CLOSE_SHIFT,
    CLOSE_SHIFT_NOTE,
    MAIN_MENU,
    OPEN_SHIFT_AMOUNT,
    REPORT_DEBT_PAYMENTS,
    REGISTER_PHONE,
    REPORT_DEBT_RECEIVED,
    REPORT_EXPENSES,
    REPORT_OTHER_PAYMENTS,
    REPORT_P2P,
    REPORT_SALES,
    SUBMIT_DAILY_REPORT,
    UPLOAD_PAYMENT_IMAGE,
    UPLOAD_RECEIPT_ROLL,
    UPLOAD_WORKPLACE_STATUS,
    UPLOAD_ZERO_REPORT,
    SardobaBot,
)
from bot_constants import ADMIN_MENU_ROWS, CASHIER_MENU_ROWS, EXPORT_MENU_ROWS
from sardoba_bot.common.utils import hash_password


class FakeDB:
    def __init__(self, fetch_one_results=None, fetch_all_result=None):
        self.fetch_one_results = list(fetch_one_results or [])
        self.fetch_all_result = list(fetch_all_result or [])
        self.connect_calls = 0
        self.disconnect_calls = 0
        self.fetch_one_calls = []
        self.fetch_all_calls = []
        self.execute_calls = []

    async def connect(self):
        self.connect_calls += 1
        return True

    async def disconnect(self):
        self.disconnect_calls += 1

    async def fetch_one(self, query, params=None):
        self.fetch_one_calls.append((query, params))
        if self.fetch_one_results:
            return self.fetch_one_results.pop(0)
        return None

    async def fetch_all(self, query, params=None):
        self.fetch_all_calls.append((query, params))
        return list(self.fetch_all_result)

    async def execute_query(self, query, params=None):
        self.execute_calls.append((query, params))
        return True


def make_bot(fake_db=None):
    bot = SardobaBot.__new__(SardobaBot)
    bot.db = fake_db or FakeDB()
    bot.export_utils = None
    bot._group_chat_id_cache = None
    bot._locations_cache = None
    return bot


class FakeMessage:
    def __init__(self, text=None, contact=None):
        self.text = text
        self.contact = contact
        self.replies = []
        self.documents = []

    async def reply_text(self, text, reply_markup=None):
        self.replies.append({"text": text, "reply_markup": reply_markup})

    async def reply_document(self, document=None, caption=None, **kwargs):
        self.documents.append({"document": document, "caption": caption, **kwargs})


def make_text_update(text, user_id=42, first_name="Ali"):
    message = FakeMessage(text=text)
    update = SimpleNamespace(message=message, effective_user=SimpleNamespace(id=user_id, first_name=first_name))
    return update, message


@pytest.mark.asyncio
async def test_group_chat_id_is_cached():
    bot = make_bot(FakeDB(fetch_one_results=[{"group_chat_id": -100123}]))
    assert await bot._get_group_chat_id() == -100123
    assert await bot._get_group_chat_id() == -100123
    assert len(bot.db.fetch_one_calls) == 1


@pytest.mark.asyncio
async def test_locations_are_cached_and_resolved():
    locations = [{"id": 1, "name": "Sardoba (Severniy)", "address": "", "is_active": True}]
    bot = make_bot(FakeDB(fetch_all_result=locations))
    assert await bot._get_location_name(1) == "Sardoba (Severniy)"
    assert await bot._get_locations() == locations
    assert len(bot.db.fetch_all_calls) == 1


@pytest.mark.asyncio
async def test_initialize_applies_schema_file(monkeypatch, tmp_path):
    schema_path = tmp_path / "schema.sql"
    schema_path.write_text("CREATE TABLE IF NOT EXISTS users (id BIGINT PRIMARY KEY);", encoding="utf-8")
    monkeypatch.setattr("sardoba_bot.telegram.bot.POSTGRES_SCHEMA_PATH", schema_path)

    fake_db = FakeDB()
    bot = make_bot(fake_db)

    await bot.initialize()

    assert fake_db.connect_calls == 1
    assert len(fake_db.execute_calls) == 1
    assert "CREATE TABLE IF NOT EXISTS users" in fake_db.execute_calls[0][0]


def test_day_bounds_and_total_balance():
    bot = make_bot()
    start, end = bot._day_bounds("2026-04-01", "2026-04-02")
    assert start.isoformat() == "2026-04-01T00:00:00+05:00"
    assert end.isoformat() == "2026-04-03T00:00:00+05:00"
    total = bot._calculate_total_balance(
        {
            "sales_amount": 1000,
            "debt_received": 200,
            "expenses": 50,
            "uzcard_amount": 100,
            "humo_amount": 300,
            "p2p_amount": 60,
            "uzcard_refund": 20,
            "humo_refund": 10,
            "other_payments": 40,
            "debt_payments": 30,
            "debt_refunds": 5,
        }
    )
    assert total == 1595


def test_close_shift_progress_text_includes_percent_and_remaining():
    bot = make_bot()
    text = bot._build_close_shift_progress_text(3, 10, "Excel tayyorlanmoqda")
    assert "3/10 (30%)" in text
    assert "Excel tayyorlanmoqda" in text
    assert "Qoldi: 7 ta bosqich" in text


@pytest.mark.asyncio
async def test_after_sverka_step_waits_for_finish_button_when_all_done():
    bot = make_bot()
    bot.show_sverka_menu = AsyncMock()
    bot._finalize_sverka = AsyncMock(return_value=MAIN_MENU)
    update = make_text_update("x")[0]
    context = SimpleNamespace(
        user_data={
            "sverka_status": {
                "sales_amount": True,
                "debt_received": True,
                "expenses": True,
                "uzcard_amount": True,
                "humo_amount": True,
                "p2p_amount": True,
                "uzcard_refund": True,
                "humo_refund": True,
                "other_payments": True,
                "debt_payments": True,
                "debt_refunds": True,
            }
        }
    )

    state = await bot._after_sverka_step(update, context)

    assert state == SUBMIT_DAILY_REPORT
    bot.show_sverka_menu.assert_awaited_once()
    bot._finalize_sverka.assert_not_awaited()


def test_shift_full_xlsx_workbook_embeds_images():
    bot = make_bot()
    shift = {
        "first_name": "Ali",
        "last_name": "Valiyev",
        "phone_number": "+998901234567",
        "location": "Sardoba",
        "opened_at": "2026-04-07 05:41:00",
        "closed_at": "2026-04-07 10:10:00",
        "opening_amount": 1000,
        "closing_amount": 2500,
    }
    report = {"sales_amount": 1500}
    images = [{"image_type": "uzcard_payment", "image_url": "file_1", "uploaded_at": "2026-04-07 05:45:21"}]

    img = PILImage.new("RGB", (4, 4), (255, 0, 0))
    raw = BytesIO()
    img.save(raw, format="PNG")

    workbook = bot._build_shift_full_xlsx_workbook(shift, report, images, {"file_1": raw.getvalue()})

    with ZipFile(BytesIO(workbook.getvalue())) as archive:
        names = archive.namelist()
        assert any(name.startswith("xl/media/") for name in names)
        xml_payload = b"".join(archive.read(name) for name in names if name.endswith(".xml"))
        assert b"file_1" not in xml_payload


def test_build_shift_summary_message_is_readable():
    bot = make_bot()
    text = bot._build_shift_summary_message(
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
            "p2p_amount": 75_000,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
        }
    )
    assert "📊 Kunlik umumiy hisobot" in text
    assert "👤 Kassir: Nilufar G'afurova" in text
    assert "🧾 Sverka" in text
    assert "💳 P2P: 75 000" in text


def test_build_shift_summary_message_includes_expense_detail():
    bot = make_bot()
    text = bot._build_shift_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "closed_at": "2026-03-28 09:49:16",
            "opening_amount": 3000,
            "closing_amount": 100000,
            "sales_amount": 1_000_000,
            "debt_received": 0,
            "expenses": 200_000,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "report_data": {
                "expense_detail": {
                    "payment_type": "P2P",
                    "paid_to": "Yetkazib beruvchi",
                    "recipient_name": "Abror",
                    "recipient_phone": "+998901234567",
                    "reason": "Muz uchun to'lov",
                }
            },
        }
    )
    assert "📌 Chiqim tafsiloti" in text
    assert "💳 To'lov turi: P2P" in text
    assert "📝 Sabab: Muz uchun to'lov" in text


def test_build_shift_summary_message_includes_debt_received_detail():
    bot = make_bot()
    text = bot._build_shift_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "closed_at": "2026-03-28 09:49:16",
            "opening_amount": 3000,
            "closing_amount": 100000,
            "sales_amount": 1_000_000,
            "debt_received": 150_000,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "report_data": {
                "debt_received_detail": {
                    "payment_type": "Naqd",
                }
            },
        }
    )
    assert "📌 Kelgan qarz tafsiloti" in text
    assert "💳 To'lov turi: Naqd" in text


def test_build_shift_summary_message_includes_debt_payments_detail():
    bot = make_bot()
    text = bot._build_shift_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "closed_at": "2026-03-28 09:49:16",
            "opening_amount": 3000,
            "closing_amount": 100000,
            "sales_amount": 1_000_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 150_000,
            "debt_refunds": 0,
            "report_data": {
                "debt_payments_detail": {
                    "counterparty_name": "Rustam",
                    "counterparty_phone": "+998901112233",
                    "payment_type": "P2P",
                }
            },
        }
    )
    assert "📌 Qarz to'lovi tafsiloti" in text
    assert "👤 Kimga: Rustam" in text
    assert "💳 To'lov turi: P2P" in text


def test_parse_amount_accepts_zero_and_rejects_letters():
    bot = make_bot()
    assert bot._parse_amount("0") == 0
    assert bot._parse_amount("12 300") == 12300

    with pytest.raises(ValueError):
        bot._parse_amount("12000 so'm")

    with pytest.raises(ValueError):
        bot._parse_amount("12ming")

    with pytest.raises(ValueError):
        bot._parse_amount("12,000")


@pytest.mark.asyncio
async def test_today_shift_query_uses_range_filter():
    bot = make_bot(FakeDB(fetch_one_results=[{"id": 1}]))
    await bot._today_shift_for_user(7)
    query, params = bot.db.fetch_one_calls[0]
    assert "DATE(opened_at)" not in query
    assert params[0] == 7
    assert params[1].isoformat().endswith("00:00:00+05:00")


@pytest.mark.asyncio
async def test_save_daily_report_inserts_when_missing():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 100,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
        }
    )

    await bot.save_daily_report(None, context)

    assert "INSERT INTO reports" in fake_db.execute_calls[0][0]


@pytest.mark.asyncio
async def test_save_daily_report_updates_when_existing():
    fake_db = FakeDB(fetch_one_results=[{"id": 11}])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 100,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
        }
    )

    await bot.save_daily_report(None, context)

    assert "UPDATE reports" in fake_db.execute_calls[0][0]


@pytest.mark.asyncio
async def test_save_daily_report_persists_expense_detail_json():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 100,
            "debt_received": 0,
            "expenses": 120_000,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "expense_payment_type": "P2P",
            "expense_paid_to": "Yetkazib beruvchi",
            "expense_recipient_name": "Abror",
            "expense_recipient_phone": "+998901234567",
            "expense_reason": "Muz uchun to'lov",
        }
    )

    await bot.save_daily_report(None, context)

    query, params = fake_db.execute_calls[0]
    assert "report_data" in query
    payload = json.loads(params["report_data"])
    assert payload["expense_detail"]["payment_type"] == "P2P"
    assert payload["expense_detail"]["recipient_phone"] == "+998901234567"


@pytest.mark.asyncio
async def test_save_daily_report_persists_debt_received_payment_type_json():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 100,
            "debt_received": 50_000,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "debt_received_payment_type": "P2P",
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    payload = json.loads(params["report_data"])
    assert payload["debt_received_detail"]["payment_type"] == "P2P"


@pytest.mark.asyncio
async def test_save_daily_report_persists_p2p_amount_column():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 120_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 35_000,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    assert params["p2p_amount"] == 35_000


@pytest.mark.asyncio
async def test_save_daily_report_persists_generic_payment_methods_json():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 120_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 10_000,
            "debt_payments": 0,
            "debt_refunds": 0,
            "other_payments_payment_type": "Naqd",
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    payload = json.loads(params["report_data"])
    assert payload["payment_methods"]["other_payments"] == "Naqd"


@pytest.mark.asyncio
async def test_report_debt_received_positive_amount_requests_payment_type():
    bot = make_bot()
    update, message = make_text_update("50000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_received", "pending_sverka_state": REPORT_DEBT_RECEIVED}
    )

    state = await bot.report_debt_received(update, context)

    assert state == REPORT_DEBT_RECEIVED
    assert context.user_data["debt_received"] == 50000
    assert context.user_data["debt_received_detail_stage"] == "counterparty_name"
    assert "kimdan keldi" in message.replies[-1]["text"].lower()


@pytest.mark.asyncio
async def test_report_sales_positive_amount_completes_without_payment_type():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("50000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "sales_amount", "pending_sverka_state": REPORT_SALES}
    )

    state = await bot.report_sales(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["sales_amount"] == 50000
    assert "sales_amount_detail_stage" not in context.user_data
    assert "sales_amount_payment_type" not in context.user_data
    assert not message.replies
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_debt_received_collects_payment_type_and_returns_to_sverka():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_received", "pending_sverka_state": REPORT_DEBT_RECEIVED}
    )

    await bot.report_debt_received(make_text_update("50000")[0], context)
    await bot.report_debt_received(make_text_update("Jamshid")[0], context)
    await bot.report_debt_received(make_text_update("+998901234567")[0], context)
    state = await bot.report_debt_received(make_text_update("P2P")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_received_payment_type"] == "P2P"
    assert context.user_data["debt_received_counterparty_name"] == "Jamshid"
    assert context.user_data["debt_received_counterparty_phone"] == "+998901234567"
    assert "debt_received_detail_stage" not in context.user_data
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_other_payments_positive_amount_requests_payment_type():
    bot = make_bot()
    update, message = make_text_update("25000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "other_payments", "pending_sverka_state": REPORT_OTHER_PAYMENTS}
    )

    state = await bot.report_other_payments(update, context)

    assert state == REPORT_OTHER_PAYMENTS
    assert context.user_data["other_payments"] == 25000
    assert context.user_data["other_payments_detail_stage"] == "payment_type"
    assert isinstance(message.replies[-1]["reply_markup"], ReplyKeyboardMarkup)


@pytest.mark.asyncio
async def test_report_p2p_saves_amount_and_returns_to_sverka():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "p2p_amount", "pending_sverka_state": REPORT_P2P}
    )

    state = await bot.report_p2p(make_text_update("120000")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["p2p_amount"] == 120000
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_debt_payments_positive_amount_completes_without_payment_type():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_payments", "pending_sverka_state": REPORT_DEBT_PAYMENTS}
    )

    state = await bot.report_debt_payments(make_text_update("120000")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_payments"] == 120000
    assert "debt_payments_detail_stage" not in context.user_data
    assert "debt_payments_payment_type" not in context.user_data
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_debt_received_zero_skips_payment_type():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_received", "pending_sverka_state": REPORT_DEBT_RECEIVED}
    )

    state = await bot.report_debt_received(make_text_update("0")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_received"] == 0
    assert "debt_received_detail_stage" not in context.user_data
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_expenses_positive_amount_requests_detail():
    bot = make_bot()
    update, message = make_text_update("120000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "expenses", "pending_sverka_state": REPORT_EXPENSES}
    )

    state = await bot.report_expenses(update, context)

    assert state == REPORT_EXPENSES
    assert context.user_data["expenses"] == 120000
    assert context.user_data["expense_detail_stage"] == "payment_type"
    assert isinstance(message.replies[-1]["reply_markup"], ReplyKeyboardMarkup)
    assert "To'lov turini tanlang" in message.replies[-1]["text"]


@pytest.mark.asyncio
async def test_report_expenses_collects_detail_and_returns_to_sverka():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "expenses", "pending_sverka_state": REPORT_EXPENSES}
    )

    await bot.report_expenses(make_text_update("120000")[0], context)
    await bot.report_expenses(make_text_update("P2P")[0], context)
    await bot.report_expenses(make_text_update("Yetkazib beruvchi")[0], context)
    await bot.report_expenses(make_text_update("Abror")[0], context)
    await bot.report_expenses(make_text_update("+998901234567")[0], context)
    state = await bot.report_expenses(make_text_update("Muz uchun to'lov")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["expense_payment_type"] == "P2P"
    assert context.user_data["expense_reason"] == "Muz uchun to'lov"
    assert "expense_detail_stage" not in context.user_data
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_finalize_sverka_sends_group_summary():
    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_shift_summary = AsyncMock(
        return_value={
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-04-14 10:00:00",
            "sales_amount": 1_000,
            "debt_received": 0,
            "expenses": 100,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "report_data": {},
        }
    )
    bot._send_group_message = AsyncMock()
    bot.show_cashier_menu = AsyncMock()
    context = SimpleNamespace(
        user_data={"current_shift_id": 5, "flow": "sverka", "sverka_entrypoint": "standalone"},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    state = await bot._finalize_sverka(update, context)

    assert state == MAIN_MENU
    bot.save_daily_report.assert_awaited_once()
    bot._send_group_message.assert_awaited_once()
    bot.show_cashier_menu.assert_awaited_once_with(update, context)


@pytest.mark.asyncio
async def test_finalize_sverka_from_closing_continues_to_close_amount_step():
    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_shift_summary = AsyncMock(return_value={"location": "Sardoba", "report_data": {}})
    bot._send_group_message = AsyncMock()
    bot._prompt_close_shift_amount = AsyncMock(return_value=CLOSE_SHIFT)
    bot.show_cashier_menu = AsyncMock()
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "flow": "sverka",
            "sverka_entrypoint": "closing",
            "pending_sverka_key": "sales_amount",
            "pending_sverka_state": REPORT_SALES,
            "expense_detail_stage": "payment_type",
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    state = await bot._finalize_sverka(update, context)

    assert state == CLOSE_SHIFT
    bot.save_daily_report.assert_awaited_once()
    bot._send_group_message.assert_awaited_once()
    bot._prompt_close_shift_amount.assert_awaited_once()
    bot.show_cashier_menu.assert_not_awaited()
    assert "pending_sverka_key" not in context.user_data
    assert "expense_detail_stage" not in context.user_data


@pytest.mark.asyncio
async def test_show_sverka_menu_includes_cancel_button():
    bot = make_bot()
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99))

    await bot.show_sverka_menu(update, context)

    kwargs = context.bot.send_message.await_args.kwargs
    markup = kwargs["reply_markup"]
    labels = [button.text for row in markup.inline_keyboard for button in row]
    assert "🟢 Yakunlash" in labels
    assert "❌ Bekor qilish" in labels


@pytest.mark.asyncio
async def test_sverka_cancel_button_exits_flow_and_clears_state():
    bot = make_bot()
    bot.show_cashier_menu = AsyncMock()
    query = SimpleNamespace(data="sv:cancel", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={
            "flow": "sverka",
            "pending_sverka_key": "sales_amount",
            "pending_sverka_state": REPORT_SALES,
            "sverka_status": {"sales_amount": False},
            "expense_detail_stage": "payment_type",
            "debt_refunds_detail_stage": "payment_type",
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == MAIN_MENU
    assert context.user_data.get("flow") is None
    assert "pending_sverka_key" not in context.user_data
    assert "sverka_status" not in context.user_data
    assert "expense_detail_stage" not in context.user_data
    assert "debt_refunds_detail_stage" not in context.user_data
    bot.show_cashier_menu.assert_awaited_once_with(update, context)


@pytest.mark.asyncio
async def test_send_group_shift_photo_sends_photo_with_caption():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"group_chat_id": -100123},
                {
                    "first_name": "Ali",
                    "last_name": "Valiyev",
                    "location": "Sardoba",
                    "opened_at": "2026-04-14 10:00:00",
                },
            ]
        )
    )
    context = SimpleNamespace(bot=SimpleNamespace(send_photo=AsyncMock(), send_document=AsyncMock()))

    sent = await bot._send_group_shift_photo(context, 5, "file_photo", "Ish joyi holati rasmi")

    assert sent is True
    context.bot.send_photo.assert_awaited_once()
    kwargs = context.bot.send_photo.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert kwargs["photo"] == "file_photo"
    assert "Ish joyi holati rasmi" in kwargs["caption"]
    assert "Kassir: Ali Valiyev" in kwargs["caption"]
    context.bot.send_document.assert_not_awaited()


@pytest.mark.asyncio
async def test_send_group_shift_photo_falls_back_to_document():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"group_chat_id": -100123},
                {
                    "first_name": "Ali",
                    "last_name": "Valiyev",
                    "location": "Sardoba",
                    "opened_at": "2026-04-14 10:00:00",
                },
            ]
        )
    )
    context = SimpleNamespace(
        bot=SimpleNamespace(
            send_photo=AsyncMock(side_effect=RuntimeError("photo failed")),
            send_document=AsyncMock(),
        )
    )

    sent = await bot._send_group_shift_photo(context, 5, "file_doc", "Zaxira chek lenta rasmi")

    assert sent is True
    context.bot.send_photo.assert_awaited_once()
    context.bot.send_document.assert_awaited_once()
    kwargs = context.bot.send_document.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert kwargs["document"] == "file_doc"
    assert "Zaxira chek lenta rasmi" in kwargs["caption"]


@pytest.mark.asyncio
async def test_flush_opening_group_photos_sends_media_album_with_all_items():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {
                    "first_name": "Ali",
                    "last_name": "Valiyev",
                    "location": "Sardoba",
                    "opened_at": "2026-04-14 10:00:00",
                },
                {"group_chat_id": -100123},
            ]
        )
    )
    context = SimpleNamespace(
        bot=SimpleNamespace(send_media_group=AsyncMock()),
        user_data={
            "pending_opening_group_photos": [
                {
                    "file_id": "file_1",
                    "image_title": "Ish joyi holati rasmi",
                    "event_time": "2026-04-14 10:01:00",
                    "media_kind": "photo",
                },
                {
                    "file_id": "file_2",
                    "image_title": "Ish joyi holati rasmi",
                    "event_time": "2026-04-14 10:01:05",
                    "media_kind": "photo",
                },
            ]
        },
    )

    await bot._flush_opening_group_photos(context, 5)

    context.bot.send_media_group.assert_awaited_once()
    kwargs = context.bot.send_media_group.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert len(kwargs["media"]) == 2
    assert "pending_opening_group_photos" not in context.user_data


@pytest.mark.asyncio
async def test_open_shift_amount_rejects_letters():
    bot = make_bot()
    message = FakeMessage(text="120000 so'm")
    update = SimpleNamespace(message=message, effective_user=SimpleNamespace(id=42))
    context = SimpleNamespace(user_data={"location_id": 1})

    state = await bot.open_shift_amount(update, context)

    assert state == OPEN_SHIFT_AMOUNT
    assert "faqat raqam kiriting" in message.replies[-1]["text"].lower()


@pytest.mark.asyncio
async def test_workplace_status_media_group_keeps_all_images_and_duplicates():
    bot = make_bot()
    bot._save_shift_image = AsyncMock()
    bot._send_group_shift_photo = AsyncMock(return_value=True)
    bot._get_image_file_id = lambda update: getattr(update.message, "_file_id", None)

    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "opening_stage": "workplace_status",
            "workplace_status_uploaded_ids": [],
        }
    )

    def make_image_update(file_id, media_group_id=None):
        message = FakeMessage()
        message._file_id = file_id
        message.media_group_id = media_group_id
        return SimpleNamespace(
            message=message,
            effective_user=SimpleNamespace(id=42),
            effective_chat=SimpleNamespace(id=99),
        ), message

    update1, _ = make_image_update("file_1", "grp_1")
    update2, _ = make_image_update("file_2", "grp_1")
    update3, _ = make_image_update("file_1", "grp_1")  # duplicate in same album

    state1 = await bot.upload_workplace_status(update1, context)
    state2 = await bot.upload_workplace_status(update2, context)
    state3 = await bot.upload_workplace_status(update3, context)

    assert state1 == UPLOAD_WORKPLACE_STATUS
    assert state2 == UPLOAD_WORKPLACE_STATUS
    assert state3 == UPLOAD_WORKPLACE_STATUS
    assert context.user_data["opening_stage"] == "workplace_status"
    assert context.user_data["opening_stage_locked_media_group_id"] == "grp_1"
    assert len(context.user_data["workplace_status_uploaded_ids"]) == 3
    assert len(context.user_data["pending_opening_group_photos"]) == 3
    assert bot._save_shift_image.await_count == 3
    assert bot._send_group_shift_photo.await_count == 0

    update4, _ = make_image_update("file_terminal", None)
    state4 = await bot.upload_workplace_status(update4, context)

    assert state4 == UPLOAD_ZERO_REPORT
    assert context.user_data["opening_stage"] == "zero_report"
    bot._save_shift_image.assert_any_await(5, "terminal_power", "file_terminal")


@pytest.mark.asyncio
async def test_upload_receipt_roll_sends_group_opening_message():
    bot = make_bot()
    bot._save_shift_image = AsyncMock()
    bot._finalize_shift_opening_flow = AsyncMock()
    bot._get_image_file_id = lambda update: "file_1"
    message = FakeMessage()
    update = SimpleNamespace(
        message=message,
        effective_chat=SimpleNamespace(id=99),
        effective_user=SimpleNamespace(id=42, first_name="Ali", last_name="Valiyev"),
    )
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "location_id": 1,
            "opening_amount": 120000,
            "opening_amount_time": "2026-04-14 14:44:00",
            "pending_opening_group_photos": [
                {"file_id": "file_prev", "image_title": "Ish joyi holati rasmi", "event_time": None}
            ],
        },
        bot=SimpleNamespace(),
    )

    state = await bot.upload_receipt_roll(update, context)

    assert state == MAIN_MENU
    bot._save_shift_image.assert_awaited_once_with(5, "receipt_roll", "file_1")
    bot._finalize_shift_opening_flow.assert_awaited_once()


@pytest.mark.asyncio
async def test_upload_receipt_roll_media_group_waits_for_debounce_finalize():
    bot = make_bot()
    bot._save_shift_image = AsyncMock()
    bot._schedule_receipt_roll_finalize = Mock()
    bot._finalize_shift_opening_flow = AsyncMock()
    bot._get_image_file_id = lambda update: "file_1"
    message = FakeMessage()
    message.media_group_id = "grp_1"
    update = SimpleNamespace(
        message=message,
        effective_chat=SimpleNamespace(id=99),
        effective_user=SimpleNamespace(id=42, first_name="Ali", last_name="Valiyev"),
    )
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "location_id": 1,
            "opening_amount": 120000,
            "opening_amount_time": "2026-04-14 14:44:00",
        },
        bot=SimpleNamespace(),
    )

    state = await bot.upload_receipt_roll(update, context)

    assert state == UPLOAD_RECEIPT_ROLL
    bot._save_shift_image.assert_awaited_once_with(5, "receipt_roll", "file_1")
    bot._schedule_receipt_roll_finalize.assert_called_once()
    bot._finalize_shift_opening_flow.assert_not_awaited()


@pytest.mark.asyncio
async def test_register_phone_requires_contact_share_button():
    bot = make_bot()
    message = FakeMessage(text="+998901234567")
    update = SimpleNamespace(message=message, effective_user=SimpleNamespace(id=42))
    context = SimpleNamespace(user_data={})

    state = await bot.register_phone(update, context)

    assert state == REGISTER_PHONE
    assert "oddiy matn emas" in message.replies[-1]["text"]
    assert isinstance(message.replies[-1]["reply_markup"], ReplyKeyboardMarkup)
    assert message.replies[-1]["reply_markup"].keyboard[0][0].request_contact is True


@pytest.mark.asyncio
async def test_register_phone_accepts_shared_contact():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    bot.notify_admins_new_request = AsyncMock()
    message = FakeMessage(contact=SimpleNamespace(phone_number="+998901234567", user_id=42))
    update = SimpleNamespace(message=message, effective_user=SimpleNamespace(id=42))
    context = SimpleNamespace(user_data={"first_name": "Ali", "last_name": "Valiyev"})

    state = await bot.register_phone(update, context)

    assert state == ConversationHandler.END
    assert context.user_data["phone"] == "+998901234567"
    assert len(fake_db.execute_calls) == 1
    assert isinstance(message.replies[-1]["reply_markup"], ReplyKeyboardRemove)


@pytest.mark.asyncio
async def test_unknown_user_is_asked_to_start_instead_of_restarting_role_prompt():
    fake_db = FakeDB(fetch_one_results=[None, None, None])
    bot = make_bot(fake_db)
    bot.start = AsyncMock(side_effect=AssertionError("start should not be called"))
    message = FakeMessage(text="Hisobotlar")
    update = SimpleNamespace(message=message, effective_user=SimpleNamespace(id=77, first_name="Ali"))
    context = SimpleNamespace(user_data={})

    state = await bot.handle_message(update, context)

    assert state == ConversationHandler.END
    assert message.replies[-1]["text"] == "Avval /start buyrug'ini yuboring."


@pytest.mark.asyncio
async def test_handle_message_closing_flow_routes_to_note_step_when_amount_already_saved():
    bot = make_bot()
    bot.close_shift = AsyncMock()
    bot.close_shift_note = AsyncMock()
    update, _ = make_text_update("har qanday izoh", user_id=42)
    context = SimpleNamespace(user_data={"flow": "closing", "pending_close_amount": 600000})

    await bot.handle_message(update, context)

    bot.close_shift_note.assert_awaited_once_with(update, context)
    bot.close_shift.assert_not_awaited()


@pytest.mark.asyncio
async def test_handle_message_closing_flow_routes_to_amount_step_when_amount_missing():
    bot = make_bot()
    bot.close_shift = AsyncMock()
    bot.close_shift_note = AsyncMock()
    update, _ = make_text_update("600000", user_id=42)
    context = SimpleNamespace(user_data={"flow": "closing"})

    await bot.handle_message(update, context)

    bot.close_shift.assert_awaited_once_with(update, context)
    bot.close_shift_note.assert_not_awaited()


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("text", "method_name"),
    [
        ("Hisobotlar", "show_admin_reports_menu"),
        ("Orqaga", "show_admin_menu"),
        ("Barcha kassirlar", "send_all_cashiers"),
        ("Kassir so'rovlari", "handle_approval_requests"),
        ("Ma'lumotlarni o'zgartirish", "modify_user_data"),
        ("Excel/PDF yuklab olish", "export_data"),
    ],
)
async def test_admin_buttons_dispatch_to_expected_handlers(text, method_name):
    bot = make_bot()
    update, _ = make_text_update(text)
    context = SimpleNamespace(user_data={})
    user = {"role": "admin", "first_name": "Admin"}

    for name in (
        "show_admin_reports_menu",
        "show_admin_menu",
        "send_all_cashiers",
        "handle_approval_requests",
        "modify_user_data",
        "export_data",
    ):
        setattr(bot, name, AsyncMock())

    await bot.handle_admin_command(update, context, user)

    getattr(bot, method_name).assert_awaited_once_with(update, context)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("text", "period"),
    [("Kunlik", "daily"), ("Haftalik", "weekly"), ("Oylik", "monthly")],
)
async def test_admin_report_period_buttons_dispatch_to_location_prompt(text, period):
    bot = make_bot()
    update, _ = make_text_update(text)
    context = SimpleNamespace(user_data={})
    user = {"role": "admin", "first_name": "Admin"}
    bot._ask_report_location = AsyncMock()

    await bot.handle_admin_command(update, context, user)

    bot._ask_report_location.assert_awaited_once_with(update, context, period)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("text", "method_name"),
    [
        ("Smena ochish", "start_shift_opening"),
        ("Smena yopish", "start_shift_closing"),
        ("Sverka", "start_daily_reporting"),
        ("Rasm jo'natish", "start_payment_image_upload"),
        ("Hisobotlarni tahrirlash", "edit_reports"),
    ],
)
async def test_cashier_buttons_dispatch_to_expected_handlers(text, method_name):
    bot = make_bot()
    update, _ = make_text_update(text)
    context = SimpleNamespace(user_data={})
    user = {"role": "cashier", "first_name": "Cashier"}

    for name in (
        "start_shift_opening",
        "start_shift_closing",
        "start_daily_reporting",
        "start_payment_image_upload",
        "edit_reports",
    ):
        setattr(bot, name, AsyncMock())

    await bot.handle_cashier_command(update, context, user)

    getattr(bot, method_name).assert_awaited_once_with(update, context)


@pytest.mark.asyncio
async def test_cashier_password_success_resumes_saved_action():
    password = "0000"
    bot = make_bot(FakeDB(fetch_one_results=[{"role": "cashier", "first_name": "Ali", "password_hash": hash_password(password)}]))
    bot.start_shift_opening = AsyncMock(return_value=OPEN_SHIFT_AMOUNT)
    bot.show_cashier_menu = AsyncMock()
    update, _ = make_text_update(password)
    context = SimpleNamespace(user_data={"cashier_pending_password": True})
    bot._set_cashier_resume_action(context, "start_shift_opening")

    await bot.handle_message(update, context)

    assert context.user_data["cashier_authenticated"] is True
    assert context.user_data["cashier_pending_password"] is False
    bot.start_shift_opening.assert_awaited_once_with(update, context)
    bot.show_cashier_menu.assert_not_awaited()
    assert "cashier_resume_action" not in context.user_data


@pytest.mark.asyncio
async def test_cashier_password_success_with_expired_action_returns_to_menu():
    password = "0000"
    bot = make_bot(FakeDB(fetch_one_results=[{"role": "cashier", "first_name": "Ali", "password_hash": hash_password(password)}]))
    bot.start_shift_opening = AsyncMock()
    bot.show_cashier_menu = AsyncMock()
    update, message = make_text_update(password)
    expired_at = (bot._now_tashkent() - timedelta(minutes=16)).isoformat()
    context = SimpleNamespace(
        user_data={
            "cashier_pending_password": True,
            "cashier_resume_action": "start_shift_opening",
            "cashier_resume_action_at": expired_at,
        }
    )

    await bot.handle_message(update, context)

    bot.start_shift_opening.assert_not_awaited()
    bot.show_cashier_menu.assert_awaited_once_with(update, context)
    assert "sessiya muddati tugagan" in message.replies[-1]["text"].lower()


@pytest.mark.asyncio
async def test_admin_export_menu_uses_defined_buttons():
    bot = make_bot()
    update, message = make_text_update("Excel/PDF yuklab olish")
    context = SimpleNamespace(user_data={})

    await bot.export_data(update, context)

    markup = message.replies[-1]["reply_markup"]
    assert isinstance(markup, ReplyKeyboardMarkup)
    assert [[button.text for button in row] for row in markup.keyboard] == [list(row) for row in EXPORT_MENU_ROWS]


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("text", "generator_name", "report_type"),
    [
        ("Kunlik hisobot (Excel)", "generate_excel_report", "daily"),
        ("Kunlik hisobot (PDF)", "generate_pdf_report", "daily"),
        ("Kassirlar bo'yicha (Excel)", "generate_excel_report", "cashier_performance"),
        ("Kassirlar bo'yicha (PDF)", "generate_pdf_report", "cashier_performance"),
    ],
)
async def test_export_choice_uses_expected_report_generator(text, generator_name, report_type):
    bot = make_bot()
    bot.export_utils = SimpleNamespace(
        generate_excel_report=AsyncMock(return_value=BytesIO(b"excel")),
        generate_pdf_report=AsyncMock(return_value=BytesIO(b"pdf")),
    )
    update, message = make_text_update(text)
    context = SimpleNamespace(user_data={})

    await bot.handle_export_choice(update, context)

    getattr(bot.export_utils, generator_name).assert_awaited_once_with(report_type=report_type)
    assert message.documents


@pytest.mark.asyncio
async def test_menu_layouts_match_constants():
    bot = make_bot()
    context = SimpleNamespace(user_data={})

    admin_update, admin_message = make_text_update("x")
    cashier_update, cashier_message = make_text_update("x")

    await bot.show_admin_menu(admin_update, context)
    await bot.show_cashier_menu(cashier_update, context)

    admin_markup = admin_message.replies[-1]["reply_markup"]
    cashier_markup = cashier_message.replies[-1]["reply_markup"]

    assert [[button.text for button in row] for row in admin_markup.keyboard] == [list(row) for row in ADMIN_MENU_ROWS]
    assert [[button.text for button in row] for row in cashier_markup.keyboard] == [list(row) for row in CASHIER_MENU_ROWS]


@pytest.mark.asyncio
async def test_approve_cashier_prompts_password_setup_before_menu():
    fake_db = FakeDB(
        fetch_one_results=[
            {
                "telegram_id": 77,
                "first_name": "Ali",
                "last_name": "Valiyev",
                "phone_number": "+998901234567",
                "password_hash": "hashed",
                "status": "pending",
            },
            None,
        ]
    )
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        bot=SimpleNamespace(send_message=AsyncMock()),
        application=SimpleNamespace(_user_data=defaultdict(dict)),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=-1001))

    await bot.approve_cashier(update, context, 77)

    cashier_call = context.bot.send_message.await_args_list[0]
    assert cashier_call.kwargs["chat_id"] == 77
    assert "tasdiqlandi" in cashier_call.kwargs["text"]
    assert "yangi parol" in cashier_call.kwargs["text"]
    assert isinstance(cashier_call.kwargs["reply_markup"], ReplyKeyboardRemove)
    assert context.application._user_data[77]["cashier_set_password"] is True
    assert context.application._user_data[77]["cashier_set_password_confirm"] is False
    assert any(
        "INSERT INTO users" in query and params[-1] is None
        for query, params in fake_db.execute_calls
    )


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("payment_key", "expected_image_type", "expected_title", "success_reply", "missing_prompt"),
    [
        (
            "uzcard",
            "uzcard_payment",
            "Uzcard hisobot rasmi",
            "Uzcard hisobot rasmingiz qabul qilindi.",
            "Humo rasmini ham yuboring.",
        ),
        (
            "humo",
            "humo_payment",
            "Humo hisobot rasmi",
            "Humo hisobot rasmingiz qabul qilindi.",
            "Uzcard rasmini ham yuboring.",
        ),
    ],
)
async def test_upload_payment_image_prompts_for_missing_pair_after_first_upload(
    payment_key, expected_image_type, expected_title, success_reply, missing_prompt
):
    fake_db = FakeDB()
    bot = make_bot(fake_db)
    bot._get_image_file_id = lambda update: "file_1"
    bot._send_group_shift_photo = AsyncMock(return_value=True)
    bot._get_shift_meta = AsyncMock(return_value={"cashier": "Ali Valiyev", "location": "Sardoba"})
    bot._send_group_message = AsyncMock(return_value=True)
    if payment_key == "uzcard":
        bot._count_shift_images = AsyncMock(side_effect=[1, 0])
    else:
        bot._count_shift_images = AsyncMock(side_effect=[0, 1])
    bot.show_cashier_menu = AsyncMock()
    update = SimpleNamespace(
        message=FakeMessage(),
        effective_user=SimpleNamespace(id=42, first_name="Ali", last_name="Valiyev"),
    )
    context = SimpleNamespace(
        user_data={
            "pending_payment_image": payment_key,
            "current_shift_id": 5,
            "flow": "payment_image",
        }
    )

    state = await bot.upload_payment_image(update, context)

    assert state == UPLOAD_PAYMENT_IMAGE
    assert fake_db.execute_calls[0][1] == (5, "file_1", expected_image_type)
    bot._send_group_shift_photo.assert_awaited_once_with(
        context,
        5,
        "file_1",
        expected_title,
        event_time=None,
    )
    bot._send_group_message.assert_awaited_once()
    group_text = bot._send_group_message.await_args.args[1]
    assert expected_title in group_text
    assert "Ali Valiyev" in group_text
    assert context.user_data.get("flow") == "payment_image"
    assert context.user_data["pending_payment_image"] in {"uzcard", "humo"}
    assert context.user_data["pending_payment_image"] != payment_key
    assert update.message.replies[-2]["text"] == success_reply
    assert update.message.replies[-1]["text"] == missing_prompt
    bot.show_cashier_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_upload_payment_image_moves_to_close_amount_when_pair_complete_for_closing_flow():
    fake_db = FakeDB()
    bot = make_bot(fake_db)
    bot._get_image_file_id = lambda update: "file_1"
    bot._send_group_shift_photo = AsyncMock(return_value=True)
    bot._get_shift_meta = AsyncMock(return_value={"cashier": "Ali Valiyev", "location": "Sardoba"})
    bot._send_group_message = AsyncMock(return_value=True)
    bot._count_shift_images = AsyncMock(side_effect=[1, 1])
    bot.show_cashier_menu = AsyncMock()
    update = SimpleNamespace(
        message=FakeMessage(),
        effective_user=SimpleNamespace(id=42, first_name="Ali", last_name="Valiyev"),
    )
    context = SimpleNamespace(
        user_data={
            "pending_payment_image": "uzcard",
            "current_shift_id": 5,
            "flow": "payment_image",
            "awaiting_payment_images_for_close": True,
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    bot.show_sverka_menu = AsyncMock()

    state = await bot.upload_payment_image(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data.get("flow") == "sverka"
    assert context.user_data.get("sverka_entrypoint") == "closing"
    assert "pending_payment_image" not in context.user_data
    assert "awaiting_payment_images_for_close" not in context.user_data
    assert update.message.replies[-1]["text"] == "Uzcard va Humo rasmlari to'liq qabul qilindi."
    bot.show_sverka_menu.assert_awaited_once()
    bot.show_cashier_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_start_shift_closing_redirects_to_payment_upload_when_images_missing():
    fake_db = FakeDB(
        fetch_one_results=[
            {"id": 10},  # user
            {"id": 5},  # active shift
        ]
    )
    bot = make_bot(fake_db)
    bot._ensure_cashier_authenticated = AsyncMock(return_value=True)
    bot._ensure_opening_requirements_completed = AsyncMock(return_value=True)
    bot._count_shift_images = AsyncMock(side_effect=[0, 1])
    bot.start_payment_image_upload = AsyncMock(return_value=UPLOAD_PAYMENT_IMAGE)
    update, message = make_text_update("Smena yopish", user_id=42)
    context = SimpleNamespace(user_data={})

    state = await bot.start_shift_closing(update, context)

    assert state == UPLOAD_PAYMENT_IMAGE
    assert context.user_data["awaiting_payment_images_for_close"] is True
    assert context.user_data["current_shift_id"] == 5
    assert "majburiy" in message.replies[-1]["text"]
    bot.start_payment_image_upload.assert_awaited_once_with(update, context)


@pytest.mark.asyncio
async def test_start_shift_closing_starts_final_sverka_when_images_ready():
    fake_db = FakeDB(
        fetch_one_results=[
            {"id": 10},
            {"id": 5},
        ]
    )
    bot = make_bot(fake_db)
    bot._ensure_cashier_authenticated = AsyncMock(return_value=True)
    bot._ensure_opening_requirements_completed = AsyncMock(return_value=True)
    bot._count_shift_images = AsyncMock(side_effect=[1, 1])
    bot._start_sverka_flow = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, _ = make_text_update("Smena yopish", user_id=42)
    context = SimpleNamespace(user_data={})

    state = await bot.start_shift_closing(update, context)

    assert state == SUBMIT_DAILY_REPORT
    bot._start_sverka_flow.assert_awaited_once_with(
        update,
        context,
        5,
        entrypoint="closing",
        force_reset=True,
        note="Smenani yopishdan oldin yakuniy sverkani to'ldiring.",
    )


@pytest.mark.asyncio
async def test_close_shift_sends_only_text_to_group_without_excel_exports():
    fake_db = FakeDB(fetch_one_results=[{"id": 10}], fetch_all_result=[{"id": 5}])
    bot = make_bot(fake_db)
    bot._get_shift_summary = AsyncMock(
        return_value={
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "closing_amount": 350000,
            "closed_at": "2026-04-17 12:10:00",
        }
    )
    bot._send_group_message = AsyncMock(return_value=True)
    bot._send_group_document = AsyncMock()
    bot.show_cashier_menu = AsyncMock()
    amount_update, amount_message = make_text_update("350000", user_id=42)
    note_update, message = make_text_update("Hammasi joyida, kassa topshirildi.", user_id=42)
    context = SimpleNamespace(user_data={})

    amount_state = await bot.close_shift(amount_update, context)
    state = await bot.close_shift_note(note_update, context)

    assert amount_state == CLOSE_SHIFT_NOTE
    assert "Izoh kiriting" in amount_message.replies[-1]["text"]
    assert state == MAIN_MENU
    assert any("UPDATE shifts SET closing_amount" in query for query, _ in fake_db.execute_calls)
    bot._send_group_message.assert_awaited_once()
    sent_text = bot._send_group_message.await_args.args[1]
    assert "🔒 Smena yopildi" in sent_text
    assert "👤 Kassir: Ali Valiyev" in sent_text
    assert "💰 Yopish summasi: 350 000" in sent_text
    assert sent_text.strip().endswith("📝 Izoh: Hammasi joyida, kassa topshirildi.")
    bot._send_group_document.assert_not_awaited()
    assert message.replies[-1]["text"] == "Smena yopildi."
    bot.show_cashier_menu.assert_awaited_once_with(note_update, context)


@pytest.mark.asyncio
async def test_close_shift_warns_user_when_group_message_fails():
    fake_db = FakeDB(fetch_one_results=[{"id": 10}], fetch_all_result=[{"id": 5}])
    bot = make_bot(fake_db)
    bot._get_shift_summary = AsyncMock(
        return_value={
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "closing_amount": 200000,
            "closed_at": "2026-04-17 12:10:00",
        }
    )
    bot._send_group_message = AsyncMock(return_value=False)
    bot.show_cashier_menu = AsyncMock()
    amount_update, amount_message = make_text_update("200000", user_id=42)
    note_update, message = make_text_update("Bugun savdo sust bo'ldi.", user_id=42)
    context = SimpleNamespace(user_data={})

    amount_state = await bot.close_shift(amount_update, context)
    state = await bot.close_shift_note(note_update, context)

    assert amount_state == CLOSE_SHIFT_NOTE
    assert "Izoh kiriting" in amount_message.replies[-1]["text"]
    assert state == MAIN_MENU
    assert "guruhga yuborilmadi" in message.replies[-2]["text"]
    assert message.replies[-1]["text"] == "Smena yopildi."
    bot.show_cashier_menu.assert_awaited_once_with(note_update, context)


@pytest.mark.asyncio
async def test_finalize_shift_opening_flow_flushes_pending_photos_when_group_send_succeeds():
    bot = make_bot()
    bot._get_location_name = AsyncMock(return_value="Sardoba")
    bot._send_group_message = AsyncMock(return_value=True)
    bot._flush_opening_group_photos = AsyncMock()
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "location_id": 1,
            "opening_amount": 120000,
            "opening_amount_time": "2026-04-14 14:44:00",
            "pending_opening_group_photos": [{"file_id": "file_1", "image_title": "Ish joyi holati rasmi"}],
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    await bot._finalize_shift_opening_flow(
        context,
        chat_id=99,
        cashier_first_name="Ali",
        cashier_last_name="Valiyev",
    )

    bot._send_group_message.assert_awaited_once()
    group_text = bot._send_group_message.await_args.args[1]
    assert "Smena ochildi: Ali Valiyev" in group_text
    assert "Ochish summasi: 120 000" in group_text
    bot._flush_opening_group_photos.assert_awaited_once_with(context, 5)

    sent_texts = [call.kwargs["text"] for call in context.bot.send_message.await_args_list]
    assert any("Smena muvaffaqiyatli ochildi" in text for text in sent_texts)
    assert not any("guruhga yuborilmadi" in text for text in sent_texts)


@pytest.mark.asyncio
async def test_finalize_shift_opening_flow_warns_when_group_send_fails():
    bot = make_bot()
    bot._get_location_name = AsyncMock(return_value="Sardoba")
    bot._send_group_message = AsyncMock(return_value=False)
    bot._flush_opening_group_photos = AsyncMock()
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "location_id": 1,
            "opening_amount": 120000,
            "opening_amount_time": "2026-04-14 14:44:00",
            "pending_opening_group_photos": [{"file_id": "file_1", "image_title": "Ish joyi holati rasmi"}],
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    await bot._finalize_shift_opening_flow(
        context,
        chat_id=99,
        cashier_first_name="Ali",
        cashier_last_name="Valiyev",
    )

    bot._flush_opening_group_photos.assert_not_awaited()
    sent_texts = [call.kwargs["text"] for call in context.bot.send_message.await_args_list]
    assert any("guruhga yuborilmadi" in text for text in sent_texts)


@pytest.mark.asyncio
async def test_flush_opening_group_photos_falls_back_to_individual_uploads():
    bot = make_bot()
    bot._get_shift_meta = AsyncMock(return_value={"cashier": "Ali Valiyev", "location": "Sardoba"})
    bot._send_group_media_album = AsyncMock(return_value=False)
    bot._send_group_shift_photo = AsyncMock(return_value=True)
    context = SimpleNamespace(
        user_data={
            "pending_opening_group_photos": [
                {"file_id": "file_1", "image_title": "Ish joyi holati rasmi", "event_time": "2026-04-14 10:01:00"},
                {"file_id": "file_2", "image_title": "Ish joyi holati rasmi", "event_time": "2026-04-14 10:01:05"},
            ]
        }
    )

    await bot._flush_opening_group_photos(context, 5)

    bot._send_group_media_album.assert_awaited_once()
    assert bot._send_group_shift_photo.await_count == 2
    bot._send_group_shift_photo.assert_any_await(
        context,
        5,
        "file_1",
        "Ish joyi holati rasmi",
        event_time="2026-04-14 10:01:00",
    )
    bot._send_group_shift_photo.assert_any_await(
        context,
        5,
        "file_2",
        "Ish joyi holati rasmi",
        event_time="2026-04-14 10:01:05",
    )
    assert "pending_opening_group_photos" not in context.user_data
