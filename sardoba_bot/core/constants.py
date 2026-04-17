def _flatten(rows):
    return {label for row in rows for label in row}


(
    SELECT_ROLE, REGISTER_FIRSTNAME, REGISTER_LASTNAME,
    REGISTER_PHONE, REGISTER_PASSWORD, VERIFY_PASSWORD, ADMIN_LOGIN,
    CASHIER_LOGIN, MAIN_MENU, OPEN_SHIFT_AMOUNT, SELECT_LOCATION,
    UPLOAD_WORKPLACE_STATUS, UPLOAD_TERMINAL_POWER, UPLOAD_ZERO_REPORT,
    UPLOAD_OPENING_NOTIFICATION, UPLOAD_RECEIPT_ROLL, SUBMIT_SHIFT_OPENING,
    SELECT_PAYMENT_IMAGE, UPLOAD_PAYMENT_IMAGE,
    REPORT_SALES, REPORT_DEBT_RECEIVED, REPORT_EXPENSES, REPORT_UZCARD,
    REPORT_HUMO, REPORT_UZCARD_REFUND, REPORT_HUMO_REFUND, REPORT_OTHER_PAYMENTS,
    REPORT_DEBT_PAYMENTS, REPORT_DEBT_REFUNDS, SUBMIT_DAILY_REPORT, CLOSE_SHIFT,
    CLOSE_SHIFT_NOTE,
    EDIT_REPORT_SELECT, EDIT_REPORT_VALUE,
    ADMIN_REGISTER_PHONE, ADMIN_REGISTER_PASSWORD, ADMIN_VERIFY_PASSWORD,
) = range(37)

ADMIN_MENU_ROWS = (
    ("Hisobotlar", "Barcha kassirlar"),
    ("Kassir so'rovlari", "Ma'lumotlarni o'zgartirish"),
    ("Excel/PDF yuklab olish",),
)

ADMIN_REPORTS_MENU_ROWS = (
    ("Kunlik", "Haftalik"),
    ("Oylik", "Vaqt oralig'i"),
    ("Orqaga",),
)

CASHIER_MENU_ROWS = (
    ("Smena ochish", "Smena yopish"),
    ("Sverka", "Rasm jo'natish"),
    ("Hisobotlarni tahrirlash",),
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
}

CASHIER_DIRECT_ACTIONS = {
    "Smena ochish": "start_shift_opening",
    "Smena yopish": "start_shift_closing",
    "Sverka": "start_daily_reporting",
    "Rasm jo'natish": "start_payment_image_upload",
    "Hisobotlarni tahrirlash": "edit_reports",
}
