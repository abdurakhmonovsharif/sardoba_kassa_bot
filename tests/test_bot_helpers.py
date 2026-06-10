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
    OPENING_GROUP_IMAGE_TITLES,
    REPORT_DEBT_PAYMENTS,
    REGISTER_PHONE,
    REPORT_DEBT_RECEIVED,
    REPORT_EXPENSES,
    REPORT_OTHER_PAYMENTS,
    REPORT_P2P,
    REPORT_SALES,
    REPORT_TAX_INFO,
    REPORT_UZCARD,
    SELECT_LOCATION,
    SUBMIT_DAILY_REPORT,
    UPLOAD_OPENING_NOTIFICATION,
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
        self.photo = []
        self.document = None
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
    assert total == 620


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
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot._after_sverka_step(update, context)

    assert state == SUBMIT_DAILY_REPORT
    bot.show_sverka_menu.assert_not_awaited()
    bot._finalize_sverka.assert_not_awaited()
    context.bot.send_message.assert_awaited_once()


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
    assert "🧮 Naqd kutiladigan summa: 9 725 000" in text


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
                    "items": [
                        {"text": "Mirshod Dastafka -- 10 000", "amount": 10_000},
                        {"text": "Ulug Paynet -- 190 000", "amount": 190_000},
                    ],
                    "cash_amount": 50_000,
                }
            },
        }
    )
    assert "📌 Chiqim tafsiloti" in text
    assert "• Mirshod Dastafka -- 10 000" in text
    assert "• Ulug Paynet -- 190 000" in text
    assert "💵 Naqd summa: 50 000" in text


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


def test_build_shift_summary_message_includes_debt_received_items():
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
            "sales_amount": 0,
            "debt_received": 80_000,
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
                    "items": [
                        {
                            "counterparty_name": "Jamshid",
                            "counterparty_phone": "+998901234567",
                            "amount": 50_000,
                            "payment_type": "P2P",
                        },
                        {
                            "counterparty_name": "Sharif",
                            "counterparty_phone": "+998931434413",
                            "amount": 30_000,
                            "payment_type": "Naqd",
                        },
                    ]
                }
            },
        }
    )
    assert "📌 Kelgan qarzlar tafsiloti" in text
    assert "1. 👤 Jamshid" in text
    assert "💳 P2P" in text
    assert "2. 👤 Sharif" in text
    assert "💳 Naqd" in text


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


def test_build_sverka_summary_message_can_use_closing_title_and_tax_info():
    bot = make_bot()
    text = bot._build_sverka_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "sales_amount": 10_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 99_000,
            "humo_amount": 88_000,
            "p2p_amount": 77_000,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "report_data": {
                "tax_info": {
                    "check_image": "file_tax",
                    "cash_amount": 50_000,
                    "card_amount": 25_000,
                }
            },
        },
        title="🔒 Kassa yopilishi ma'lumotlari",
        closing=True,
    )

    assert text.startswith("🔒 Kassa yopilishi ma'lumotlari")
    assert "🧾 Soliq ma'lumotlari" in text
    assert "💵 Soliq naqdga berilgan summa: 50 000" in text
    assert "💳 Soliq plastikka berilgan summa: 25 000" in text
    assert "📷 Soliq z-otchet rasmi biriktirilgan" in text
    assert "💳 Uzcard:" not in text
    assert "💳 Humo:" not in text
    assert "💳 P2P:" not in text
    assert "🧮 Naqd kutiladigan summa: 10 000" in text


def test_build_sverka_summary_message_includes_expected_cash_formula():
    bot = make_bot()
    text = bot._build_sverka_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "sales_amount": 1_000,
            "debt_received": 200,
            "debt_refunds": 5,
            "expenses": 50,
            "uzcard_amount": 100,
            "humo_amount": 300,
            "p2p_amount": 60,
            "uzcard_refund": 20,
            "humo_refund": 10,
            "other_payments": 40,
            "debt_payments": 30,
            "report_data": {},
        }
    )

    assert "🧮 Naqd kutiladigan summa: 620" in text


def test_build_sverka_summary_message_mentions_no_expenses_when_skipped():
    bot = make_bot()
    text = bot._build_sverka_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "sales_amount": 10_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "report_data": {"expense_detail": {"skipped": True}},
        },
        title="🔒 Kassa yopilishi ma'lumotlari",
        closing=True,
    )

    assert "📌 Chiqim tafsiloti" in text
    assert "Xarajat mavjud emas" in text


def test_build_sverka_summary_message_mentions_skipped_debts():
    bot = make_bot()
    text = bot._build_sverka_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-03-28 09:27:21",
            "sales_amount": 10_000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "report_data": {
                "debt_received_detail": {"skipped": True},
                "debt_payments_detail": {"skipped": True},
            },
        },
        title="🔒 Kassa yopilishi ma'lumotlari",
        closing=True,
    )

    assert "Kelgan qarz mavjud emas" in text
    assert "Qarzga berilgan summa mavjud emas" in text


def test_decorate_debt_payment_check_image_adds_header_and_border():
    bot = make_bot()
    raw = BytesIO()
    PILImage.new("RGB", (320, 180), (240, 240, 240)).save(raw, format="JPEG")

    decorated = bot._decorate_debt_payment_check_image(raw.getvalue(), 7, "Sharif")
    image = PILImage.open(decorated)

    assert image.width > 320
    assert image.height > 180
    assert image.getpixel((10, 10)) == (20, 20, 20)


def test_decorate_tax_info_check_image_adds_header_and_border():
    bot = make_bot()
    raw = BytesIO()
    PILImage.new("RGB", (320, 180), (240, 240, 240)).save(raw, format="JPEG")

    decorated = bot._decorate_tax_info_check_image(raw.getvalue(), 25_000)
    image = PILImage.open(decorated)

    assert image.width > 320
    assert image.height > 180
    assert image.getpixel((10, 10)) == (20, 20, 20)


def test_decorate_tax_info_check_image_uses_only_left_title():
    bot = make_bot()
    expected = BytesIO(b"decorated")
    bot._decorate_labeled_check_image = Mock(return_value=expected)

    result = bot._decorate_tax_info_check_image(b"raw", 25_000)

    assert result is expected
    bot._decorate_labeled_check_image.assert_called_once_with(b"raw", "Soliq z-otchet", "", large_title=True)


@pytest.mark.asyncio
async def test_add_image_label_uses_left_header_border_style():
    class FakeTelegramFile:
        async def download_to_memory(self, buf):
            buf.write(b"raw-image")

    bot = make_bot()
    expected = BytesIO(b"decorated")
    bot._decorate_labeled_check_image = Mock(return_value=expected)
    fake_bot = SimpleNamespace(get_file=AsyncMock(return_value=FakeTelegramFile()))

    result = await bot._add_image_label(fake_bot, "file_1", OPENING_GROUP_IMAGE_TITLES["workplace_status"])

    assert result is expected
    fake_bot.get_file.assert_awaited_once_with("file_1")
    bot._decorate_labeled_check_image.assert_called_once_with(
        b"raw-image",
        OPENING_GROUP_IMAGE_TITLES["workplace_status"],
        "",
        large_title=True,
    )


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
            "expense_items": [
                {"text": "Mirshod Dastafka -- 10 000", "amount": 10_000},
                {"text": "Ulug Paynet -- 110 000", "amount": 110_000},
            ],
            "expense_cash_amount": 25_000,
        }
    )

    await bot.save_daily_report(None, context)

    query, params = fake_db.execute_calls[0]
    assert "report_data" in query
    payload = json.loads(params["report_data"])
    assert payload["expense_detail"]["items"][0]["text"] == "Mirshod Dastafka -- 10 000"
    assert payload["expense_detail"]["items"][1]["amount"] == 110000
    assert payload["expense_detail"]["cash_amount"] == 25_000


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
async def test_save_daily_report_persists_debt_received_items_json():
    fake_db = FakeDB(fetch_one_results=[None])
    bot = make_bot(fake_db)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "sales_amount": 100,
            "debt_received": 80_000,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "debt_received_items": [
                {
                    "counterparty_name": "Sharif",
                    "counterparty_phone": "+998931434413",
                    "amount": 10_000,
                    "payment_type": "Naqd",
                },
                {
                    "counterparty_name": "Jamshid",
                    "counterparty_phone": "+998901234567",
                    "amount": 70_000,
                    "payment_type": "P2P",
                },
            ],
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    payload = json.loads(params["report_data"])
    items = payload["debt_received_detail"]["items"]
    assert items[0]["counterparty_name"] == "Sharif"
    assert items[0]["payment_type"] == "Naqd"
    assert items[1]["amount"] == 70_000


@pytest.mark.asyncio
async def test_save_daily_report_persists_tax_info_json():
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
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "tax_info_check_image": "file_tax",
            "tax_info_cash_amount": 25_000,
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    payload = json.loads(params["report_data"])
    assert payload["tax_info"]["check_image"] == "file_tax"
    assert payload["tax_info"]["cash_amount"] == 25_000


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
async def test_save_daily_report_persists_other_payments_comment_json():
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
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "other_payments_comment": "Terminal bo'yicha izoh qoldirildi",
        }
    )

    await bot.save_daily_report(None, context)

    _, params = fake_db.execute_calls[0]
    payload = json.loads(params["report_data"])
    assert payload["other_payments_comment"] == "Terminal bo'yicha izoh qoldirildi"


@pytest.mark.asyncio
async def test_report_debt_received_positive_amount_completes():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("50000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_received", "pending_sverka_state": REPORT_DEBT_RECEIVED}
    )

    state = await bot.report_debt_received(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_received"] == 50000
    assert "debt_received_detail_stage" not in context.user_data
    assert not message.replies
    bot._after_sverka_step.assert_awaited_once()


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
async def test_report_debt_received_saves_single_amount():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_received", "pending_sverka_state": REPORT_DEBT_RECEIVED}
    )

    state = await bot.report_debt_received(make_text_update("80000")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_received"] == 80000
    assert "debt_received_items" not in context.user_data
    assert "debt_received_detail_stage" not in context.user_data
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_other_payments_amount_completes_and_is_saved():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("45000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "other_payments", "pending_sverka_state": REPORT_OTHER_PAYMENTS}
    )

    state = await bot.report_other_payments(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["other_payments"] == 45000
    assert "other_payments_comment" not in context.user_data
    assert not message.replies
    bot._after_sverka_step.assert_awaited_once()


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
async def test_report_tax_info_collects_image_then_cash_amount():
    fake_db = FakeDB()
    bot = make_bot(fake_db)
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "pending_sverka_key": "tax_info",
            "pending_sverka_state": REPORT_TAX_INFO,
            "tax_info_stage": "check_image",
            "sverka_entrypoint": "closing",
            "sverka_status": {"tax_info": False},
        }
    )
    image_update, image_message = make_text_update("")
    image_update.message.photo = [
        SimpleNamespace(file_id="small_tax"),
        SimpleNamespace(file_id="large_tax"),
    ]

    state = await bot.report_tax_info(image_update, context)

    assert state == REPORT_TAX_INFO
    assert context.user_data["tax_info_check_image"] == "large_tax"
    assert context.user_data["tax_info_stage"] == "cash_amount"
    assert "Naqd summani kiriting" in image_message.replies[-1]["text"]
    assert fake_db.execute_calls[0][1] == (5, "large_tax", "tax_info_check")

    state = await bot.report_tax_info(make_text_update("25000")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["tax_info_cash_amount"] == 25000
    assert "tax_info_stage" not in context.user_data
    assert "pending_sverka_key" not in context.user_data
    assert context.user_data["sverka_status"]["tax_info"] is True
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_debt_payments_positive_amount_completes():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("120000")
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_payments", "pending_sverka_state": REPORT_DEBT_PAYMENTS}
    )

    state = await bot.report_debt_payments(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_payments"] == 120000
    assert "debt_payments_detail_stage" not in context.user_data
    assert not message.replies
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_debt_payments_saves_single_amount():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={"pending_sverka_key": "debt_payments", "pending_sverka_state": REPORT_DEBT_PAYMENTS}
    )

    state = await bot.report_debt_payments(make_text_update("120000")[0], context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_payments"] == 120000
    assert "debt_payments_items" not in context.user_data
    assert "debt_payments_detail_stage" not in context.user_data
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
async def test_report_expenses_saves_amount():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("10000")
    context = SimpleNamespace(
        user_data={
            "pending_sverka_key": "expenses",
            "pending_sverka_state": REPORT_EXPENSES,
        }
    )

    state = await bot.report_expenses(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["expenses"] == 10_000
    assert "expense_items" not in context.user_data
    assert "expense_detail_stage" not in context.user_data
    assert not message.replies
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_report_expenses_invalid_amount_reprompts():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    context = SimpleNamespace(
        user_data={
            "pending_sverka_key": "expenses",
            "pending_sverka_state": REPORT_EXPENSES,
        }
    )

    update, message = make_text_update("Yakunlash")
    state = await bot.report_expenses(update, context)

    assert state == REPORT_EXPENSES
    assert "faqat raqam kiriting" in message.replies[-1]["text"].lower()
    bot._after_sverka_step.assert_not_awaited()


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
async def test_finalize_sverka_clears_values_before_next_same_shift_sverka():
    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_shift_summary = AsyncMock(return_value={"location": "Sardoba", "report_data": {}})
    bot._send_group_message = AsyncMock()
    bot.show_cashier_menu = AsyncMock()
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "flow": "sverka",
            "sverka_shift_id": 5,
            "sverka_entrypoint": "standalone",
            "sales_amount": 1000,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
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
            },
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    await bot._finalize_sverka(update, context)

    for key, *_ in bot._sverka_config():
        assert key not in context.user_data

    bot.show_sverka_menu = AsyncMock()
    await bot._start_sverka_flow(update, context, 5, entrypoint="standalone")

    assert all(done is False for done in context.user_data["sverka_status"].values())


@pytest.mark.asyncio
async def test_finalize_sverka_sends_debt_payment_check_image_to_group():
    class FakeTelegramFile:
        async def download_as_bytearray(self):
            raw = BytesIO()
            PILImage.new("RGB", (320, 180), (240, 240, 240)).save(raw, format="JPEG")
            return bytearray(raw.getvalue())

    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_group_chat_id = AsyncMock(return_value=-100123)
    send_order = []
    bot._get_shift_summary = AsyncMock(
        return_value={
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-05-22 10:00:00",
            "sales_amount": 0,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 10_000,
            "debt_refunds": 0,
            "report_data": {
                "debt_payments_detail": {
                    "items": [
                        {
                            "counterparty_name": "Sharif",
                            "counterparty_phone": "+998931434413",
                            "amount": 10_000,
                            "check_image": "file_check",
                        }
                    ]
                }
            },
        }
    )
    async def fake_send_group_message(*args, **kwargs):
        send_order.append("summary")
        return True

    async def fake_send_photo(*args, **kwargs):
        send_order.append("photo")
        return True

    bot._send_group_message = AsyncMock(side_effect=fake_send_group_message)
    bot.show_cashier_menu = AsyncMock()
    context = SimpleNamespace(
        user_data={"current_shift_id": 5, "flow": "sverka", "sverka_entrypoint": "standalone"},
        bot=SimpleNamespace(
            send_message=AsyncMock(),
            get_file=AsyncMock(return_value=FakeTelegramFile()),
            send_photo=AsyncMock(side_effect=fake_send_photo),
        ),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    state = await bot._finalize_sverka(update, context)

    assert state == MAIN_MENU
    context.bot.get_file.assert_awaited_once_with("file_check")
    context.bot.send_photo.assert_awaited_once()
    kwargs = context.bot.send_photo.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert "caption" not in kwargs
    assert send_order == ["photo", "summary"]


@pytest.mark.asyncio
async def test_finalize_sverka_from_closing_sends_photo_album_before_group_summary():
    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_group_chat_id = AsyncMock(return_value=-100123)
    send_order = []
    shift_summary = {
        "first_name": "Ali",
        "last_name": "Valiyev",
        "location": "Sardoba",
        "opened_at": "2026-05-22 10:00:00",
        "sales_amount": 0,
        "debt_received": 0,
        "expenses": 0,
        "uzcard_amount": 0,
        "humo_amount": 0,
        "p2p_amount": 0,
        "uzcard_refund": 0,
        "humo_refund": 0,
        "other_payments": 0,
        "debt_payments": 0,
        "debt_refunds": 0,
        "report_data": {
            "tax_info": {
                "check_image": "file_tax",
                "cash_amount": 25_000,
            }
        },
    }
    bot._get_shift_summary = AsyncMock(return_value=shift_summary)

    async def fake_send_group_message(*args, **kwargs):
        send_order.append("summary")
        return True

    async def fake_send_closing_album(*args, **kwargs):
        send_order.append("photo_album")
        return True

    bot._send_group_message = AsyncMock(side_effect=fake_send_group_message)
    bot._send_closing_group_photo_album = AsyncMock(side_effect=fake_send_closing_album)
    bot._prompt_close_shift_amount = AsyncMock(return_value=CLOSE_SHIFT)
    context = SimpleNamespace(
        user_data={"current_shift_id": 5, "flow": "sverka", "sverka_entrypoint": "closing"},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    state = await bot._finalize_sverka(update, context)

    assert state == CLOSE_SHIFT
    assert send_order == ["photo_album", "summary"]
    bot._send_closing_group_photo_album.assert_awaited_once_with(context, 5, shift_summary)
    group_text = bot._send_group_message.await_args.args[1]
    assert group_text.startswith("🔒 Kassa yopilishi ma'lumotlari")


def test_build_sverka_summary_message_places_comment_under_other_payments():
    bot = make_bot()

    text = bot._build_sverka_summary_message(
        {
            "first_name": "Ali",
            "last_name": "Valiyev",
            "location": "Sardoba",
            "opened_at": "2026-04-21 10:00:00",
            "sales_amount": 0,
            "debt_received": 0,
            "expenses": 0,
            "uzcard_amount": 0,
            "humo_amount": 0,
            "p2p_amount": 0,
            "uzcard_refund": 0,
            "humo_refund": 0,
            "other_payments": 0,
            "debt_payments": 0,
            "debt_refunds": 0,
            "report_data": json.dumps(
                {
                    "other_payments_comment": "Terminal bo'yicha qo'shimcha izoh",
                }
            ),
        }
    )

    lines = text.splitlines()
    idx = lines.index("🧷 Boshqa to'lovlar: izoh kiritilgan")
    assert lines[idx + 1] == "   ↳ Izoh: Terminal bo'yicha qo'shimcha izoh"
    assert "💳 To'lov turlari" not in text


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
    group_text = bot._send_group_message.await_args.args[1]
    assert group_text.startswith("🔒 Kassa yopilishi ma'lumotlari")
    bot._prompt_close_shift_amount.assert_awaited_once()
    prompt_text = bot._prompt_close_shift_amount.await_args.kwargs["text"]
    assert "Kassa yopilishi ma'lumotlari" in prompt_text
    bot.show_cashier_menu.assert_not_awaited()
    assert "pending_sverka_key" not in context.user_data
    assert "expense_detail_stage" not in context.user_data


@pytest.mark.asyncio
async def test_finalize_sverka_from_closing_is_idempotent_after_group_send():
    bot = make_bot()
    bot.save_daily_report = AsyncMock()
    bot._get_shift_summary = AsyncMock(return_value={"location": "Sardoba", "report_data": {}})
    bot._send_group_message = AsyncMock()
    bot._prompt_close_shift_amount = AsyncMock(return_value=CLOSE_SHIFT)
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "flow": "sverka",
            "sverka_entrypoint": "closing",
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99), effective_message=FakeMessage())

    state1 = await bot._finalize_sverka(update, context)
    state2 = await bot._finalize_sverka(update, context)

    assert state1 == CLOSE_SHIFT
    assert state2 == CLOSE_SHIFT
    bot.save_daily_report.assert_awaited_once()
    bot._send_group_message.assert_awaited_once()
    assert context.user_data["sverka_finalized_entrypoint"] == "closing"
    assert "oldin yuborilgan" in bot._prompt_close_shift_amount.await_args_list[-1].kwargs["text"]


@pytest.mark.asyncio
async def test_finish_after_closing_finalize_does_not_show_incomplete_menu():
    bot = make_bot()
    query = SimpleNamespace(
        data="sv:finish",
        message=SimpleNamespace(chat_id=99),
        answer=AsyncMock(),
    )
    update = SimpleNamespace(
        callback_query=query,
        effective_chat=SimpleNamespace(id=99),
        effective_message=FakeMessage(),
    )
    context = SimpleNamespace(
        user_data={
            "sverka_finalized_entrypoint": "closing",
            "sverka_status": {"debt_payments": False},
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    bot.show_sverka_menu = AsyncMock()

    state = await bot.sverka_select_step(update, context)

    assert state == CLOSE_SHIFT
    sent_text = context.bot.send_message.await_args.kwargs["text"]
    assert "oldin yuborilgan" in sent_text
    assert "Hamma band" not in sent_text
    bot.show_sverka_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_finish_after_closing_finalize_preserves_pending_close_amount():
    bot = make_bot()
    query = SimpleNamespace(
        data="sv:finish",
        message=SimpleNamespace(chat_id=99),
        answer=AsyncMock(),
    )
    update = SimpleNamespace(callback_query=query, effective_chat=SimpleNamespace(id=99))
    context = SimpleNamespace(
        user_data={
            "sverka_finalized_entrypoint": "closing",
            "pending_close_amount": 500_000,
            "sverka_status": {"debt_payments": False},
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    bot.show_sverka_menu = AsyncMock()

    state = await bot.sverka_select_step(update, context)

    assert state == CLOSE_SHIFT_NOTE
    assert context.user_data["pending_close_amount"] == 500_000
    sent_text = context.bot.send_message.await_args.kwargs["text"]
    assert "izoh kiriting" in sent_text.lower()
    bot.show_sverka_menu.assert_not_awaited()


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
    assert kwargs["text"].endswith(chr(0x2800) * 36)


@pytest.mark.asyncio
async def test_start_sverka_flow_prompts_first_step_sequentially():
    bot = make_bot()
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99))

    state = await bot._start_sverka_flow(update, context, 5, entrypoint="standalone")

    assert state == REPORT_SALES
    assert context.user_data["sverka_sequential"] is True
    assert context.user_data["pending_sverka_key"] == "sales_amount"
    kwargs = context.bot.send_message.await_args.kwargs
    assert "Bugungi savdo" in kwargs["text"]


@pytest.mark.asyncio
async def test_after_sverka_step_prompts_next_step_in_sequence():
    bot = make_bot()
    context = SimpleNamespace(
        user_data={
            "sverka_sequential": True,
            "sverka_status": {
                "sales_amount": True,
                "debt_received": False,
                "expenses": False,
                "uzcard_amount": False,
                "humo_amount": False,
                "p2p_amount": False,
                "other_payments": False,
                "debt_payments": False,
            },
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99))

    state = await bot._after_sverka_step(update, context)

    assert state == REPORT_DEBT_RECEIVED
    assert context.user_data["pending_sverka_key"] == "debt_received"
    kwargs = context.bot.send_message.await_args.kwargs
    assert "Kelgan qarz summasini" in kwargs["text"]


@pytest.mark.asyncio
async def test_show_sverka_menu_uses_full_width_step_buttons():
    bot = make_bot()
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99))

    await bot.show_sverka_menu(update, context)

    markup = context.bot.send_message.await_args.kwargs["reply_markup"]
    rows = markup.inline_keyboard
    step_rows = rows[: len(bot._sverka_config())]

    assert all(len(row) == 1 for row in step_rows)
    labels = [row[0].text for row in step_rows]
    assert "☐ Qarzga berilgan summa" in labels


@pytest.mark.asyncio
async def test_show_sverka_menu_uses_closing_image_flow_order():
    bot = make_bot()
    context = SimpleNamespace(
        user_data={"sverka_entrypoint": "closing"},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )
    update = SimpleNamespace(effective_chat=SimpleNamespace(id=99))

    await bot.show_sverka_menu(update, context)

    markup = context.bot.send_message.await_args.kwargs["reply_markup"]
    labels = [button.text for row in markup.inline_keyboard for button in row]

    expected = [
        "☐ Uzcard summasi",
        "☐ Uzcard rasmi",
        "☐ Humo summasi",
        "☐ Humo rasmi",
        "☐ P2P summasi",
        "☐ P2P rasmi",
        "☐ Soliq naqdga berilgan summa miqdori",
        "☐ Soliq plastikka berilgan summa miqdori",
        "☐ Soliq z-otchet rasmi",
        "☐ Xarajatlar",
        "☐ Kelgan qarzlar bo'yicha ma'lumot",
        "☐ Qarzga berilgan summalar bo'yicha ma'lumot",
    ]
    assert labels[: len(expected)] == expected
    assert "☐ Uzcard vozvrat" not in labels
    assert "☐ Humo vozvrat" not in labels
    assert "☐ Vozvrat qarzlar" not in labels
    assert "debt_refunds" not in context.user_data["sverka_status"]
    assert "☐ Savdo summasi" not in labels
    assert "tax_z_report_image" in context.user_data["sverka_status"]
    assert "🔴 Tugatish" in labels
    assert "🟢 Yakunlash" not in labels


@pytest.mark.asyncio
async def test_sverka_expenses_step_shows_amount_prompt():
    bot = make_bot()
    query = SimpleNamespace(data="sv:expenses", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(user_data={"sverka_status": {}}, bot=SimpleNamespace(send_message=AsyncMock()))

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_EXPENSES
    assert context.user_data["pending_sverka_key"] == "expenses"
    kwargs = context.bot.send_message.await_args.kwargs
    assert isinstance(kwargs["reply_markup"], ReplyKeyboardRemove)
    assert "Chiqim summasini kiriting" in kwargs["text"]


@pytest.mark.asyncio
async def test_closing_expenses_step_shows_add_or_skip_choice():
    bot = make_bot()
    query = SimpleNamespace(data="sv:expenses", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={"sverka_entrypoint": "closing", "sverka_status": {}},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_EXPENSES
    assert context.user_data["expense_detail_stage"] == "choice"
    kwargs = context.bot.send_message.await_args.kwargs
    assert kwargs["text"] == "Xarajat qo'shish"
    assert isinstance(kwargs["reply_markup"], ReplyKeyboardMarkup)


@pytest.mark.asyncio
async def test_closing_expenses_can_be_skipped_and_marked_absent():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=REPORT_DEBT_RECEIVED)
    update, message = make_text_update("⏭ O'tkazib yuborish")
    context = SimpleNamespace(
        user_data={
            "sverka_entrypoint": "closing",
            "expense_detail_stage": "choice",
            "pending_sverka_key": "expenses",
            "pending_sverka_state": REPORT_EXPENSES,
            "sverka_status": {"expenses": False},
        }
    )

    state = await bot.report_expenses(update, context)

    assert state == REPORT_DEBT_RECEIVED
    assert context.user_data["expenses"] == 0
    assert context.user_data["expenses_skipped"] is True
    assert context.user_data["sverka_status"]["expenses"] is True
    assert "Xarajat mavjud emas" in message.replies[-1]["text"]
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_sverka_debt_received_step_resets_previous_items_on_menu_entry():
    bot = make_bot()
    query = SimpleNamespace(data="sv:debt_received", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={
            "sverka_status": {},
            "debt_received_items": [{"counterparty_name": "Old", "amount": 1000}],
            "debt_received_detail_stage": "loop",
            "debt_received_current_amount": 1000,
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_DEBT_RECEIVED
    assert "debt_received_items" not in context.user_data
    assert "debt_received_detail_stage" not in context.user_data
    kwargs = context.bot.send_message.await_args.kwargs
    assert "Kelgan qarz summasini kiriting" in kwargs["text"]
    assert isinstance(kwargs["reply_markup"], ReplyKeyboardRemove)


@pytest.mark.asyncio
async def test_closing_debt_received_step_shows_add_or_skip_choice():
    bot = make_bot()
    query = SimpleNamespace(data="sv:debt_received", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={"sverka_entrypoint": "closing", "sverka_status": {}},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_DEBT_RECEIVED
    assert context.user_data["debt_received_detail_stage"] == "choice"
    kwargs = context.bot.send_message.await_args.kwargs
    assert kwargs["text"] == "Kelgan qarz qo'shish"
    assert isinstance(kwargs["reply_markup"], ReplyKeyboardMarkup)


@pytest.mark.asyncio
async def test_closing_debt_received_can_be_skipped_and_marked_absent():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=REPORT_DEBT_PAYMENTS)
    update, message = make_text_update("⏭ O'tkazib yuborish")
    context = SimpleNamespace(
        user_data={
            "sverka_entrypoint": "closing",
            "debt_received_detail_stage": "choice",
            "pending_sverka_key": "debt_received",
            "pending_sverka_state": REPORT_DEBT_RECEIVED,
            "sverka_status": {"debt_received": False},
        }
    )

    state = await bot.report_debt_received(update, context)

    assert state == REPORT_DEBT_PAYMENTS
    assert context.user_data["debt_received"] == 0
    assert context.user_data["debt_received_skipped"] is True
    assert context.user_data["sverka_status"]["debt_received"] is True
    assert "Kelgan qarz mavjud emas" in message.replies[-1]["text"]
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_closing_debt_payments_step_shows_add_or_skip_choice():
    bot = make_bot()
    query = SimpleNamespace(data="sv:debt_payments", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={"sverka_entrypoint": "closing", "sverka_status": {}},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_DEBT_PAYMENTS
    assert context.user_data["debt_payments_detail_stage"] == "choice"
    kwargs = context.bot.send_message.await_args.kwargs
    assert kwargs["text"] == "Qarzga berish qo'shish"
    assert isinstance(kwargs["reply_markup"], ReplyKeyboardMarkup)


@pytest.mark.asyncio
async def test_closing_debt_payments_can_be_skipped_and_marked_absent():
    bot = make_bot()
    bot._after_sverka_step = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    update, message = make_text_update("⏭ O'tkazib yuborish")
    context = SimpleNamespace(
        user_data={
            "sverka_entrypoint": "closing",
            "debt_payments_detail_stage": "choice",
            "pending_sverka_key": "debt_payments",
            "pending_sverka_state": REPORT_DEBT_PAYMENTS,
            "sverka_status": {"debt_payments": False},
        }
    )

    state = await bot.report_debt_payments(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["debt_payments"] == 0
    assert context.user_data["debt_payments_skipped"] is True
    assert context.user_data["sverka_status"]["debt_payments"] is True
    assert "Qarzga berilgan summa mavjud emas" in message.replies[-1]["text"]
    bot._after_sverka_step.assert_awaited_once()


@pytest.mark.asyncio
async def test_closing_last_skipped_step_shows_red_finish_menu():
    bot = make_bot()
    update, message = make_text_update("⏭ O'tkazib yuborish")
    context = SimpleNamespace(
        user_data={
            "sverka_entrypoint": "closing",
            "sverka_sequential": True,
            "debt_payments_detail_stage": "choice",
            "pending_sverka_key": "debt_payments",
            "pending_sverka_state": REPORT_DEBT_PAYMENTS,
            "sverka_status": {
                "uzcard_amount": True,
                "uzcard_payment_image": True,
                "humo_amount": True,
                "humo_payment_image": True,
                "p2p_amount": True,
                "p2p_payment_image": True,
                "tax_cash_amount": True,
                "tax_card_amount": True,
                "tax_z_report_image": True,
                "expenses": True,
                "debt_received": True,
                "debt_payments": False,
            },
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.report_debt_payments(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert "Qarzga berilgan summa mavjud emas" in message.replies[-1]["text"]
    kwargs = context.bot.send_message.await_args.kwargs
    assert "Barcha bandlar to'ldirildi" in kwargs["text"]
    labels = [button.text for row in kwargs["reply_markup"].inline_keyboard for button in row]
    assert "🔴 Tugatish" in labels
    assert "🟢 Yakunlash" not in labels


@pytest.mark.asyncio
async def test_sverka_tax_cash_step_starts_with_cash_amount_stage():
    bot = make_bot()
    query = SimpleNamespace(data="sv:tax_cash_amount", message=SimpleNamespace(chat_id=99), answer=AsyncMock())
    update = SimpleNamespace(callback_query=query)
    context = SimpleNamespace(
        user_data={
            "sverka_entrypoint": "closing",
            "sverka_status": {},
            "tax_info_stage": "tax_z_report_image",
            "tax_info_check_image": "old_file",
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    state = await bot.sverka_select_step(update, context)

    assert state == REPORT_TAX_INFO
    assert context.user_data["pending_sverka_key"] == "tax_cash_amount"
    assert context.user_data["tax_info_stage"] == "tax_cash_amount"
    kwargs = context.bot.send_message.await_args.kwargs
    assert "Soliq naqdga berilgan summa" in kwargs["text"]


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

    sent = await bot._send_group_shift_photo(
        context,
        5,
        "file_photo",
        OPENING_GROUP_IMAGE_TITLES["workplace_status"],
    )

    assert sent is True
    context.bot.send_photo.assert_awaited_once()
    kwargs = context.bot.send_photo.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert kwargs["photo"] == "file_photo"
    assert OPENING_GROUP_IMAGE_TITLES["workplace_status"] in kwargs["caption"]
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

    sent = await bot._send_group_shift_photo(
        context,
        5,
        "file_doc",
        OPENING_GROUP_IMAGE_TITLES["receipt_roll"],
    )

    assert sent is True
    context.bot.send_photo.assert_awaited_once()
    context.bot.send_document.assert_awaited_once()
    kwargs = context.bot.send_document.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert kwargs["document"] == "file_doc"
    assert OPENING_GROUP_IMAGE_TITLES["receipt_roll"] in kwargs["caption"]


@pytest.mark.asyncio
async def test_flush_opening_group_photos_sends_media_album_with_all_items():
    class FakeTelegramFile:
        async def download_to_memory(self, buf):
            PILImage.new("RGB", (320, 180), (240, 240, 240)).save(buf, format="JPEG")

    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"group_chat_id": -100123},
            ]
        )
    )
    context = SimpleNamespace(
        bot=SimpleNamespace(
            get_file=AsyncMock(return_value=FakeTelegramFile()),
            send_media_group=AsyncMock(),
        ),
        user_data={
            "pending_opening_group_photos": [
                {
                    "file_id": "file_1",
                    "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"],
                    "event_time": "2026-04-14 10:01:00",
                    "media_kind": "photo",
                },
                {
                    "file_id": "file_2",
                    "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"],
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
    assert context.bot.get_file.await_count == 2
    assert all(item.media not in {"file_1", "file_2"} for item in kwargs["media"])
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
async def test_open_shift_amount_allows_third_shift_for_location_today():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                {"shifts_count": 2, "has_open_shift": False, "location": "Sardoba"},
                {"id": 5},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    update, message = make_text_update("120000")
    context = SimpleNamespace(user_data={"location_id": 1})

    state = await bot.open_shift_amount(update, context)

    assert state == UPLOAD_WORKPLACE_STATUS
    assert context.user_data["current_shift_id"] == 5
    assert bot.db.execute_calls
    assert message.replies[0]["text"] == "Summa tasdiqlandi."
    bot.show_cashier_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_open_shift_amount_blocks_fourth_shift_for_location_today():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                {"shifts_count": 3, "has_open_shift": False, "location": "Sardoba"},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    update, message = make_text_update("120000")
    context = SimpleNamespace(user_data={"location_id": 1})

    state = await bot.open_shift_amount(update, context)

    assert state == MAIN_MENU
    assert context.user_data["flow"] is None
    assert not bot.db.execute_calls
    assert "bugun 3 ta smena ochilgan: Sardoba" in message.replies[0]["text"]
    bot.show_cashier_menu.assert_awaited_once_with(update, context)


@pytest.mark.asyncio
async def test_start_shift_opening_shows_locations_before_open_shift_conflict_check():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7, "role": "cashier", "password_hash": hash_password("0000")},
            ]
        )
    )
    bot.show_location_selection = AsyncMock()
    update, _ = make_text_update("Smena ochish")
    context = SimpleNamespace(user_data={"cashier_authenticated": True, "location_id": 9})

    state = await bot.start_shift_opening(update, context)

    assert state == SELECT_LOCATION
    assert "location_id" not in context.user_data
    bot.show_location_selection.assert_awaited_once_with(update, context)
    assert len(bot.db.fetch_one_calls) == 1


@pytest.mark.asyncio
async def test_select_location_blocks_existing_user_shift_with_location_name():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                {"id": 5, "location_id": 1, "location": "Sardoba"},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    query = SimpleNamespace(data="loc_1", answer=AsyncMock(), edit_message_text=AsyncMock())
    update = SimpleNamespace(
        callback_query=query,
        effective_user=SimpleNamespace(id=42),
        effective_chat=SimpleNamespace(id=99),
    )
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))

    state = await bot.select_location(update, context)

    assert state == MAIN_MENU
    assert context.user_data["flow"] is None
    assert "location_id" not in context.user_data
    text = context.bot.send_message.await_args.kwargs["text"]
    assert "Bu filialda sizda ochiq smena bor: Sardoba" in text
    bot.show_cashier_menu.assert_awaited_once_with(update, context)
    query.edit_message_text.assert_not_awaited()


@pytest.mark.asyncio
async def test_select_location_blocks_location_open_shift_with_location_name():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                None,
                {"shifts_count": 1, "has_open_shift": True, "location": "Sardoba"},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    query = SimpleNamespace(data="loc_1", answer=AsyncMock(), edit_message_text=AsyncMock())
    update = SimpleNamespace(
        callback_query=query,
        effective_user=SimpleNamespace(id=42),
        effective_chat=SimpleNamespace(id=99),
    )
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))

    state = await bot.select_location(update, context)

    assert state == MAIN_MENU
    assert context.user_data["flow"] is None
    assert "location_id" not in context.user_data
    text = context.bot.send_message.await_args.kwargs["text"]
    assert "Bu filialda hozirda ochiq smena mavjud: Sardoba" in text
    bot.show_cashier_menu.assert_awaited_once_with(update, context)
    query.edit_message_text.assert_not_awaited()


@pytest.mark.asyncio
async def test_select_location_allows_third_shift_for_location_today():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                None,
                {"shifts_count": 2, "has_open_shift": False, "location": "Sardoba"},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    query = SimpleNamespace(data="loc_1", answer=AsyncMock(), edit_message_text=AsyncMock())
    update = SimpleNamespace(
        callback_query=query,
        effective_user=SimpleNamespace(id=42),
        effective_chat=SimpleNamespace(id=99),
    )
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))

    state = await bot.select_location(update, context)

    assert state == OPEN_SHIFT_AMOUNT
    assert context.user_data["location_id"] == 1
    assert context.user_data["flow"] == "opening"
    query.edit_message_text.assert_awaited_once()
    bot.show_cashier_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_select_location_blocks_fourth_shift_for_location_today():
    bot = make_bot(
        FakeDB(
            fetch_one_results=[
                {"id": 7},
                None,
                {"shifts_count": 3, "has_open_shift": False, "location": "Sardoba"},
            ]
        )
    )
    bot.show_cashier_menu = AsyncMock()
    query = SimpleNamespace(data="loc_1", answer=AsyncMock(), edit_message_text=AsyncMock())
    update = SimpleNamespace(
        callback_query=query,
        effective_user=SimpleNamespace(id=42),
        effective_chat=SimpleNamespace(id=99),
    )
    context = SimpleNamespace(user_data={}, bot=SimpleNamespace(send_message=AsyncMock()))

    state = await bot.select_location(update, context)

    assert state == MAIN_MENU
    assert context.user_data["flow"] is None
    text = context.bot.send_message.await_args.kwargs["text"]
    assert "bugun 3 ta smena ochilgan: Sardoba" in text
    query.edit_message_text.assert_not_awaited()


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
async def test_opening_notification_requires_two_images_before_receipt_roll():
    bot = make_bot()
    bot._save_shift_image = AsyncMock()
    bot._count_shift_images = AsyncMock(side_effect=[1, 2])
    bot._get_image_file_id = lambda update: getattr(update.message, "_file_id", None)

    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "opening_stage": "opening_notification",
        }
    )

    def make_image_update(file_id):
        message = FakeMessage()
        message._file_id = file_id
        return SimpleNamespace(
            message=message,
            effective_user=SimpleNamespace(id=42),
            effective_chat=SimpleNamespace(id=99),
        ), message

    update1, message1 = make_image_update("file_iiko")
    state1 = await bot.upload_opening_notification(update1, context)

    assert state1 == UPLOAD_OPENING_NOTIFICATION
    assert context.user_data["opening_stage"] == "opening_notification"
    assert "Yana bitta rasm yuboring" in message1.replies[-1]["text"]

    update2, message2 = make_image_update("file_epos")
    state2 = await bot.upload_opening_notification(update2, context)

    assert state2 == UPLOAD_RECEIPT_ROLL
    assert context.user_data["opening_stage"] == "receipt_roll"
    assert "Zaxira cheklar" in message2.replies[-1]["text"]
    assert bot._save_shift_image.await_count == 2


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
                {"file_id": "file_prev", "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"], "event_time": None}
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
        ("🔄 Restart", "restart_session"),
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
        "restart_session",
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
        ("✅ Smena ochish", "start_shift_opening"),
        ("🔒 Smena yopish", "start_shift_closing"),
        ("📋 Sverka", "start_daily_reporting"),
        ("Rasm jo'natish", "start_payment_image_upload"),
        ("Hisobotlarni tahrirlash", "edit_reports"),
        ("🔄 Restart", "restart_session"),
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
        "restart_session",
    ):
        setattr(bot, name, AsyncMock())

    await bot.handle_cashier_command(update, context, user)

    getattr(bot, method_name).assert_awaited_once_with(update, context)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("text", "method_name"),
    [
        ("Smena ochish", "start_shift_opening"),
        ("Smena yopish", "start_shift_closing"),
        ("Sverka", "start_daily_reporting"),
    ],
)
async def test_legacy_cashier_buttons_still_dispatch(text, method_name):
    bot = make_bot()
    update, _ = make_text_update(text)
    context = SimpleNamespace(user_data={})
    user = {"role": "cashier", "first_name": "Cashier"}

    for name in ("start_shift_opening", "start_shift_closing", "start_daily_reporting"):
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
async def test_restart_button_clears_active_cashier_flow_and_requires_password():
    bot = make_bot(
        FakeDB(fetch_one_results=[{"role": "cashier", "first_name": "Ali", "password_hash": hash_password("0000")}])
    )
    update, message = make_text_update("🔄 Restart")
    context = SimpleNamespace(
        user_data={
            "flow": "opening",
            "cashier_authenticated": True,
            "pending_sverka_key": "sales_amount",
        }
    )

    await bot.handle_message(update, context)

    assert context.user_data["flow"] is None
    assert context.user_data["cashier_pending_password"] is True
    assert "cashier_authenticated" not in context.user_data
    assert message.replies[0]["text"] == "Bot qayta ishga tushirildi."
    assert message.replies[-1]["text"] == "Parolni kiriting:"


@pytest.mark.asyncio
async def test_legacy_restart_text_still_resets_session():
    bot = make_bot(
        FakeDB(fetch_one_results=[{"role": "cashier", "first_name": "Ali", "password_hash": hash_password("0000")}])
    )
    update, message = make_text_update("Restart")
    context = SimpleNamespace(user_data={"flow": "sverka", "cashier_authenticated": True})

    await bot.handle_message(update, context)

    assert context.user_data["cashier_pending_password"] is True
    assert message.replies[0]["text"] == "Bot qayta ishga tushirildi."


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
    cashier_labels = [button.text for row in cashier_markup.keyboard for button in row]
    assert "Rasm jo'natish" not in cashier_labels
    assert "Hisobotlarni tahrirlash" not in cashier_labels
    assert "✅ Smena ochish" in cashier_labels
    assert "🔒 Smena yopish" in cashier_labels
    assert "📋 Sverka" in cashier_labels
    assert "🔄 Restart" in cashier_labels


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
    state = await bot.upload_payment_image(update, context)

    assert state == REPORT_UZCARD
    assert context.user_data.get("flow") == "sverka"
    assert context.user_data.get("sverka_entrypoint") == "closing"
    assert context.user_data.get("sverka_sequential") is True
    assert context.user_data.get("pending_sverka_key") == "uzcard_amount"
    assert "pending_payment_image" not in context.user_data
    assert "awaiting_payment_images_for_close" not in context.user_data
    assert update.message.replies[-1]["text"] == "Uzcard va Humo rasmlari to'liq qabul qilindi."
    assert "Uzcard summasini kiriting" in context.bot.send_message.await_args.kwargs["text"]
    bot.show_cashier_menu.assert_not_awaited()


@pytest.mark.asyncio
async def test_upload_payment_image_completes_closing_sverka_image_step():
    bot = make_bot()
    bot._get_image_file_id = lambda update: "file_p2p"
    bot._save_shift_image = AsyncMock()
    bot._send_group_shift_photo = AsyncMock(return_value=True)
    bot._after_sverka_step = AsyncMock(return_value=REPORT_TAX_INFO)
    update = SimpleNamespace(message=FakeMessage(), effective_user=SimpleNamespace(id=42))
    context = SimpleNamespace(
        user_data={
            "flow": "sverka",
            "sverka_entrypoint": "closing",
            "pending_sverka_key": "p2p_payment_image",
            "pending_payment_image": "p2p_payment_image",
            "current_shift_id": 5,
            "sverka_status": {"p2p_payment_image": False},
        }
    )

    state = await bot.upload_payment_image(update, context)

    assert state == REPORT_TAX_INFO
    assert context.user_data["p2p_payment_image"] == "file_p2p"
    assert context.user_data["sverka_status"]["p2p_payment_image"] is True
    assert "pending_sverka_key" not in context.user_data
    assert "pending_payment_image" not in context.user_data
    bot._save_shift_image.assert_awaited_once_with(5, "p2p_payment", "file_p2p")
    bot._send_group_shift_photo.assert_not_awaited()


@pytest.mark.asyncio
async def test_start_shift_closing_starts_final_sverka_without_payment_images():
    fake_db = FakeDB(
        fetch_one_results=[
            {"id": 10},  # user
            {"id": 5},  # active shift
        ]
    )
    bot = make_bot(fake_db)
    bot._ensure_cashier_authenticated = AsyncMock(return_value=True)
    bot._ensure_opening_requirements_completed = AsyncMock(return_value=True)
    bot._count_shift_images = AsyncMock()
    bot._start_sverka_flow = AsyncMock(return_value=SUBMIT_DAILY_REPORT)
    bot.start_payment_image_upload = AsyncMock(return_value=UPLOAD_PAYMENT_IMAGE)
    update, message = make_text_update("Smena yopish", user_id=42)
    context = SimpleNamespace(user_data={})

    state = await bot.start_shift_closing(update, context)

    assert state == SUBMIT_DAILY_REPORT
    assert context.user_data["current_shift_id"] == 5
    bot._count_shift_images.assert_not_awaited()
    bot.start_payment_image_upload.assert_not_awaited()
    bot._start_sverka_flow.assert_awaited_once_with(
        update,
        context,
        5,
        entrypoint="closing",
        force_reset=True,
        note="Smenani yopishdan oldin yakuniy sverkani to'ldiring.",
    )


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
            "expenses": 110_000,
            "report_data": {
                "expense_detail": {
                    "items": [
                        {"text": "Mirshod Dastafka -- 10 000", "amount": 10_000},
                        {"text": "Ulug Paynet -- 100 000", "amount": 100_000},
                    ]
                }
            },
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
    assert "📝 Izoh: Hammasi joyida, kassa topshirildi." in sent_text
    assert "📉 Chiqim: 110 000" in sent_text
    assert "• Mirshod Dastafka -- 10 000" in sent_text
    assert "• Ulug Paynet -- 100 000" in sent_text
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
    bot._schedule_opening_group_notifications = Mock()
    context = SimpleNamespace(
        user_data={
            "current_shift_id": 5,
            "location_id": 1,
            "opening_amount": 120000,
            "opening_amount_time": "2026-04-14 14:44:00",
            "pending_opening_group_photos": [
                {"file_id": "file_1", "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"]}
            ],
        },
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    await bot._finalize_shift_opening_flow(
        context,
        chat_id=99,
        cashier_first_name="Ali",
        cashier_last_name="Valiyev",
    )

    bot._schedule_opening_group_notifications.assert_called_once()
    scheduled_args = bot._schedule_opening_group_notifications.call_args.args
    assert scheduled_args[0] is context
    assert scheduled_args[1] == [
        {"file_id": "file_1", "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"]}
    ]
    group_text = scheduled_args[2]
    assert "Smena ochildi: Ali Valiyev" in group_text
    assert "Ochish summasi: 120 000" in group_text
    assert scheduled_args[3] == 99

    sent_texts = [call.kwargs["text"] for call in context.bot.send_message.await_args_list]
    assert any("Smena muvaffaqiyatli ochildi" in text for text in sent_texts)
    assert not any("guruhga yuborilmadi" in text for text in sent_texts)


@pytest.mark.asyncio
async def test_finalize_shift_opening_flow_warns_when_group_send_fails():
    bot = make_bot()
    bot._send_opening_group_photo_album = AsyncMock(return_value=True)
    bot._send_group_message = AsyncMock(return_value=False)
    context = SimpleNamespace(
        user_data={},
        bot=SimpleNamespace(send_message=AsyncMock()),
    )

    await bot._send_opening_group_notifications(
        context,
        [{"file_id": "file_1", "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"]}],
        "Smena ochildi: Ali Valiyev",
        99,
    )

    bot._send_opening_group_photo_album.assert_awaited_once()
    bot._send_group_message.assert_awaited_once_with(context, "Smena ochildi: Ali Valiyev")
    sent_texts = [call.kwargs["text"] for call in context.bot.send_message.await_args_list]
    assert any("guruhga yuborilmadi" in text for text in sent_texts)


@pytest.mark.asyncio
async def test_flush_opening_group_photos_clears_queue_without_bot_client():
    bot = make_bot()
    context = SimpleNamespace(
        user_data={
            "pending_opening_group_photos": [
                {
                    "file_id": "file_1",
                    "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"],
                    "event_time": "2026-04-14 10:01:00",
                },
                {
                    "file_id": "file_2",
                    "image_title": OPENING_GROUP_IMAGE_TITLES["workplace_status"],
                    "event_time": "2026-04-14 10:01:05",
                },
            ]
        }
    )

    await bot._flush_opening_group_photos(context, 5)

    assert "pending_opening_group_photos" not in context.user_data


@pytest.mark.asyncio
async def test_send_closing_group_photo_album_sends_labeled_media_group():
    class FakeTelegramFile:
        async def download_to_memory(self, buf):
            raw = BytesIO()
            PILImage.new("RGB", (320, 180), (240, 240, 240)).save(raw, format="JPEG")
            buf.write(raw.getvalue())

    bot = make_bot(FakeDB(fetch_one_results=[{"group_chat_id": -100123}]))
    file_ids = {"file_uz", "file_hu", "file_p2p", "file_debt", "file_tax"}
    context = SimpleNamespace(
        user_data={
            "uzcard_payment_image": "file_uz",
            "humo_payment_image": "file_hu",
            "p2p_payment_image": "file_p2p",
            "uzcard_payment_image_media_kind": "photo",
            "humo_payment_image_media_kind": "photo",
            "p2p_payment_image_media_kind": "photo",
        },
        bot=SimpleNamespace(
            get_file=AsyncMock(return_value=FakeTelegramFile()),
            send_media_group=AsyncMock(),
        ),
    )
    row = {
        "report_data": {
            "debt_payments_detail": {
                "items": [
                    {
                        "counterparty_name": "Sharif",
                        "amount": 10_000,
                        "check_image": "file_debt",
                    }
                ]
            },
            "tax_info": {
                "check_image": "file_tax",
                "cash_amount": 25_000,
            },
        }
    }

    sent = await bot._send_closing_group_photo_album(context, 5, row)

    assert sent is True
    context.bot.send_media_group.assert_awaited_once()
    kwargs = context.bot.send_media_group.await_args.kwargs
    assert kwargs["chat_id"] == -100123
    assert len(kwargs["media"]) == 5
    assert context.bot.get_file.await_count == 5
    assert all(item.media not in file_ids for item in kwargs["media"])
    assert "pending_opening_group_photos" not in context.user_data
