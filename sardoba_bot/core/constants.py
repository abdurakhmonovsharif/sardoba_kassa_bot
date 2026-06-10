def _flatten(rows):
    return {label for row in rows for label in row}


RESTART_MENU_TEXT = "🔄 Restart"
LEGACY_RESTART_MENU_TEXT = "Restart"
OPEN_SHIFT_MENU_TEXT = "✅ Smena ochish"
CLOSE_SHIFT_MENU_TEXT = "🔒 Smena yopish"
SVERKA_MENU_TEXT = "📋 Sverka"
LEGACY_OPEN_SHIFT_MENU_TEXT = "Smena ochish"
LEGACY_CLOSE_SHIFT_MENU_TEXT = "Smena yopish"
LEGACY_SVERKA_MENU_TEXT = "Sverka"

(
    SELECT_ROLE, REGISTER_FIRSTNAME, REGISTER_LASTNAME,
    REGISTER_PHONE, REGISTER_PASSWORD, VERIFY_PASSWORD, ADMIN_LOGIN,
    CASHIER_LOGIN, MAIN_MENU, OPEN_SHIFT_AMOUNT, SELECT_LOCATION,
    UPLOAD_WORKPLACE_STATUS, UPLOAD_TERMINAL_POWER, UPLOAD_ZERO_REPORT,
    UPLOAD_OPENING_NOTIFICATION, UPLOAD_RECEIPT_ROLL, SUBMIT_SHIFT_OPENING,
    SELECT_PAYMENT_IMAGE, UPLOAD_PAYMENT_IMAGE,
    REPORT_SALES, REPORT_DEBT_RECEIVED, REPORT_EXPENSES, REPORT_UZCARD,
    REPORT_HUMO, REPORT_P2P, REPORT_UZCARD_REFUND, REPORT_HUMO_REFUND,
    REPORT_OTHER_PAYMENTS, REPORT_DEBT_PAYMENTS, REPORT_DEBT_REFUNDS, SUBMIT_DAILY_REPORT, CLOSE_SHIFT,
    CLOSE_SHIFT_NOTE,
    EDIT_REPORT_SELECT, EDIT_REPORT_VALUE,
    ADMIN_REGISTER_PHONE, ADMIN_REGISTER_PASSWORD, ADMIN_VERIFY_PASSWORD,
    REPORT_TAX_INFO,
) = range(39)

ADMIN_MENU_ROWS = (
    ("Hisobotlar", "Barcha kassirlar"),
    ("Kassir so'rovlari", "Ma'lumotlarni o'zgartirish"),
    ("Excel/PDF yuklab olish",),
    (RESTART_MENU_TEXT,),
)

ADMIN_REPORTS_MENU_ROWS = (
    ("Kunlik", "Haftalik"),
    ("Oylik", "Vaqt oralig'i"),
    ("Orqaga",),
)

CASHIER_MENU_ROWS = (
    (OPEN_SHIFT_MENU_TEXT, CLOSE_SHIFT_MENU_TEXT),
    (SVERKA_MENU_TEXT, RESTART_MENU_TEXT),
)

EXPORT_MENU_ROWS = (
    ("Kunlik hisobot (Excel)", "Kunlik hisobot (PDF)"),
    ("Kassirlar bo'yicha (Excel)", "Kassirlar bo'yicha (PDF)"),
    ("Orqaga",),
)

ADMIN_MENU_TEXTS = _flatten(ADMIN_MENU_ROWS)
ADMIN_REPORT_TEXTS = _flatten(ADMIN_REPORTS_MENU_ROWS)
CASHIER_MENU_TEXTS = _flatten(CASHIER_MENU_ROWS)
EXPORT_MENU_TEXTS = _flatten(EXPORT_MENU_ROWS)
KNOWN_MENU_TEXTS = ADMIN_MENU_TEXTS | ADMIN_REPORT_TEXTS | CASHIER_MENU_TEXTS | EXPORT_MENU_TEXTS

ADMIN_REPORT_PERIODS = {
    "Kunlik": "daily",
    "Haftalik": "weekly",
    "Oylik": "monthly",
}

ADMIN_DIRECT_ACTIONS = {
    "Hisobotlar": "show_admin_reports_menu",
    "Orqaga": "show_admin_menu",
    "Barcha kassirlar": "send_all_cashiers",
    "Kassir so'rovlari": "handle_approval_requests",
    "Ma'lumotlarni o'zgartirish": "modify_user_data",
    "Excel/PDF yuklab olish": "export_data",
    RESTART_MENU_TEXT: "restart_session",
    LEGACY_RESTART_MENU_TEXT: "restart_session",
}

CASHIER_DIRECT_ACTIONS = {
    OPEN_SHIFT_MENU_TEXT: "start_shift_opening",
    CLOSE_SHIFT_MENU_TEXT: "start_shift_closing",
    SVERKA_MENU_TEXT: "start_daily_reporting",
    LEGACY_OPEN_SHIFT_MENU_TEXT: "start_shift_opening",
    LEGACY_CLOSE_SHIFT_MENU_TEXT: "start_shift_closing",
    LEGACY_SVERKA_MENU_TEXT: "start_daily_reporting",
    "Rasm jo'natish": "start_payment_image_upload",
    "Hisobotlarni tahrirlash": "edit_reports",
    RESTART_MENU_TEXT: "restart_session",
    LEGACY_RESTART_MENU_TEXT: "restart_session",
}
