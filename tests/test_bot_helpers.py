from io import BytesIO
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import ConversationHandler

from bot import REGISTER_PHONE, SardobaBot
from bot_constants import ADMIN_MENU_ROWS, CASHIER_MENU_ROWS, EXPORT_MENU_ROWS


class FakeDB:
    def __init__(self, fetch_one_results=None, fetch_all_result=None):
        self.fetch_one_results = list(fetch_one_results or [])
        self.fetch_all_result = list(fetch_all_result or [])
        self.fetch_one_calls = []
        self.fetch_all_calls = []
        self.execute_calls = []

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
            "uzcard_refund": 20,
            "humo_refund": 10,
            "other_payments": 40,
            "debt_payments": 30,
            "debt_refunds": 5,
        }
    )
    assert total == 1535


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
