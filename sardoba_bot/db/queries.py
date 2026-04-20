class CommonQueries:
    ACTIVE_USER_BY_TELEGRAM_ID = (
        "SELECT * FROM users WHERE telegram_id = %s AND is_active = TRUE"
    )
    USER_BY_TELEGRAM_ID = "SELECT * FROM users WHERE telegram_id = %s"
    USER_ID_BY_TELEGRAM_ID = "SELECT id FROM users WHERE telegram_id = %s"
    ACTIVE_LOCATION_LIST = (
        "SELECT id, name, address, is_active FROM locations WHERE is_active = TRUE ORDER BY id"
    )
    LOCATION_NAME_BY_ID = "SELECT name FROM locations WHERE id = %s"
    ACTIVE_APPROVED_REQUEST_BY_TELEGRAM_ID = (
        "SELECT * FROM approval_requests WHERE telegram_id = %s AND status = 'approved' "
        "ORDER BY approved_at DESC LIMIT 1"
    )
    PENDING_REQUEST_BY_TELEGRAM_ID = (
        "SELECT * FROM approval_requests WHERE telegram_id = %s AND status = 'pending'"
    )
    INSERT_APPROVED_CASHIER_USER = """
        INSERT INTO users (telegram_id, first_name, last_name, phone_number, role, password_hash, is_active)
        VALUES (%s, %s, %s, %s, 'cashier', %s, TRUE)
    """
    UPDATE_PENDING_APPROVAL_REQUEST = """
        UPDATE approval_requests
        SET first_name=%s, last_name=%s, phone_number=%s
        WHERE telegram_id=%s AND status='pending'
    """
    INSERT_PENDING_APPROVAL_REQUEST = """
        INSERT INTO approval_requests (telegram_id, first_name, last_name, phone_number, role)
        VALUES (%(telegram_id)s, %(first_name)s, %(last_name)s, %(phone_number)s, %(role)s)
    """
    UPDATE_PENDING_APPROVAL_REQUEST_WITH_PASSWORD = """
        UPDATE approval_requests
        SET first_name=%s, last_name=%s, phone_number=%s, password_hash=%s
        WHERE telegram_id=%s AND status='pending'
    """
    INSERT_PENDING_APPROVAL_REQUEST_WITH_PASSWORD = """
        INSERT INTO approval_requests (telegram_id, first_name, last_name, phone_number, role, password_hash)
        VALUES (%(telegram_id)s, %(first_name)s, %(last_name)s, %(phone_number)s, %(role)s, %(password_hash)s)
    """
    BOT_GROUP_CHAT_ID = "SELECT group_chat_id FROM bot_settings WHERE id = 1"
    UPSERT_BOT_GROUP = """
        INSERT INTO bot_settings (id, group_chat_id, updated_by, updated_at)
        VALUES (1, %s, %s, NOW())
        ON CONFLICT (id) DO UPDATE
        SET group_chat_id = EXCLUDED.group_chat_id,
            updated_by = EXCLUDED.updated_by,
            updated_at = NOW()
    """


class AdminQueries:
    ACTIVE_ADMIN_ID_BY_TELEGRAM_ID = (
        "SELECT id FROM users WHERE telegram_id = %s AND role = 'admin' AND is_active = TRUE"
    )
    ACTIVE_ADMIN_BY_TELEGRAM_ID = (
        "SELECT * FROM users WHERE telegram_id = %s AND role = 'admin'"
    )
    ACTIVE_ADMIN_COUNT = (
        "SELECT COUNT(*) as cnt FROM users WHERE role = 'admin' AND is_active = TRUE"
    )
    ACTIVE_ADMIN_TELEGRAM_IDS = (
        "SELECT telegram_id FROM users WHERE role = 'admin' AND is_active = TRUE"
    )
    CASHIER_BY_TELEGRAM_ID = (
        "SELECT * FROM users WHERE telegram_id = %s AND role = 'cashier'"
    )
    RESET_CASHIER_PASSWORD = "UPDATE users SET password_hash = NULL WHERE telegram_id = %s"


class CashierQueries:
    OPEN_SHIFT_BY_USER_ID = (
        "SELECT * FROM shifts WHERE user_id=%s AND is_open=TRUE"
    )
    LATEST_OPEN_SHIFT_BY_USER_ID = (
        "SELECT id FROM shifts WHERE user_id=%s AND is_open=TRUE ORDER BY id DESC LIMIT 1"
    )
    ALL_OPEN_SHIFTS_BY_USER_ID = (
        "SELECT id FROM shifts WHERE user_id=%s AND is_open=TRUE ORDER BY id DESC"
    )
    SHIFT_IMAGE_COUNT = (
        "SELECT COUNT(*) AS count FROM images WHERE shift_id=%s AND image_type=%s"
    )


class ReportQueries:
    DAILY_EXCEL_BASE = """
        SELECT
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
            COALESCE(r.debt_refunds, 0) AS debt_refunds,
            (
                COALESCE(r.sales_amount, 0)
                + COALESCE(r.debt_received, 0)
                + COALESCE(r.uzcard_amount, 0)
                + COALESCE(r.humo_amount, 0)
                + COALESCE(r.p2p_amount, 0)
                + COALESCE(r.other_payments, 0)
                + COALESCE(r.debt_refunds, 0)
                - COALESCE(r.expenses, 0)
                - COALESCE(r.debt_payments, 0)
                - COALESCE(r.uzcard_refund, 0)
                - COALESCE(r.humo_refund, 0)
            ) AS total_balance
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
        WHERE 1=1
    """

    CASHIER_PERFORMANCE_EXCEL = """
        SELECT
            u.first_name,
            u.last_name,
            u.phone_number,
            COUNT(DISTINCT s.id) AS shifts_count,
            COALESCE(SUM(r.sales_amount), 0) AS total_sales,
            COALESCE(AVG(r.sales_amount), 0) AS avg_sales,
            COALESCE(SUM(r.expenses), 0) AS total_expenses
        FROM users u
        LEFT JOIN shifts s ON u.id = s.user_id
        LEFT JOIN LATERAL (
            SELECT sales_amount, expenses
            FROM reports
            WHERE shift_id = s.id AND report_type = 'daily_report'
            ORDER BY id DESC
            LIMIT 1
        ) r ON TRUE
        WHERE u.role = 'cashier'
        GROUP BY u.id
    """

    DAILY_PDF_BASE = """
        SELECT
            CONCAT(u.first_name, ' ', u.last_name) AS cashier_name,
            l.name AS location,
            CAST(s.opened_at AS DATE) AS date,
            CAST(s.opened_at AS TIME) AS open_time,
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
            COALESCE(r.debt_refunds, 0) AS debt_refunds,
            (
                COALESCE(r.sales_amount, 0)
                + COALESCE(r.debt_received, 0)
                + COALESCE(r.uzcard_amount, 0)
                + COALESCE(r.humo_amount, 0)
                + COALESCE(r.p2p_amount, 0)
                + COALESCE(r.other_payments, 0)
                + COALESCE(r.debt_refunds, 0)
                - COALESCE(r.expenses, 0)
                - COALESCE(r.debt_payments, 0)
                - COALESCE(r.uzcard_refund, 0)
                - COALESCE(r.humo_refund, 0)
            ) AS total_balance
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
        WHERE 1=1
    """

    CASHIER_PERFORMANCE_PDF = """
        SELECT
            CONCAT(u.first_name, ' ', u.last_name) AS cashier_name,
            u.phone_number,
            COUNT(DISTINCT s.id) AS shifts_count,
            COALESCE(SUM(r.sales_amount), 0) AS total_sales,
            COALESCE(AVG(r.sales_amount), 0) AS avg_sales,
            COALESCE(SUM(r.expenses), 0) AS total_expenses
        FROM users u
        LEFT JOIN shifts s ON u.id = s.user_id
        LEFT JOIN LATERAL (
            SELECT sales_amount, expenses
            FROM reports
            WHERE shift_id = s.id AND report_type = 'daily_report'
            ORDER BY id DESC
            LIMIT 1
        ) r ON TRUE
        WHERE u.role = 'cashier'
        GROUP BY u.id
    """
