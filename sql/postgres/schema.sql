-- PostgreSQL schema for Sardoba Restaurant Telegram Bot

-- Users
CREATE TABLE IF NOT EXISTS users (
    id BIGSERIAL PRIMARY KEY,
    telegram_id BIGINT UNIQUE NOT NULL,
    first_name VARCHAR(255) NOT NULL,
    last_name VARCHAR(255),
    phone_number VARCHAR(20),
    role VARCHAR(20) NOT NULL CHECK (role IN ('admin', 'cashier')),
    password_hash VARCHAR(255),
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Locations
CREATE TABLE IF NOT EXISTS locations (
    id BIGSERIAL PRIMARY KEY,
    name VARCHAR(255) NOT NULL UNIQUE,
    address VARCHAR(500) NOT NULL,
    is_active BOOLEAN DEFAULT TRUE
);

INSERT INTO locations (name, address) VALUES
('Sardoba (Geofizika)', 'Geofizika district'),
('Sardoba (G''ijduvon)', 'G''ijduvon district'),
('Sardoba (Severniy)', 'Severniy district'),
('Sardoba (MK-5)', 'MK-5 district')
ON CONFLICT (name) DO NOTHING;

-- Shifts
CREATE TABLE IF NOT EXISTS shifts (
    id BIGSERIAL PRIMARY KEY,
    user_id BIGINT NOT NULL REFERENCES users(id),
    location_id BIGINT NOT NULL REFERENCES locations(id),
    opened_at TIMESTAMPTZ DEFAULT NOW(),
    closed_at TIMESTAMPTZ NULL,
    opening_amount NUMERIC(12, 2),
    closing_amount NUMERIC(12, 2) DEFAULT NULL,
    is_open BOOLEAN DEFAULT TRUE
);

-- Reports
CREATE TABLE IF NOT EXISTS reports (
    id BIGSERIAL PRIMARY KEY,
    shift_id BIGINT NOT NULL REFERENCES shifts(id),
    report_type VARCHAR(30) NOT NULL CHECK (report_type IN ('shift_opening', 'shift_closing', 'daily_report')),
    sales_amount NUMERIC(12, 2) DEFAULT 0,
    debt_received NUMERIC(12, 2) DEFAULT 0,
    expenses NUMERIC(12, 2) DEFAULT 0,
    uzcard_amount NUMERIC(12, 2) DEFAULT 0,
    humo_amount NUMERIC(12, 2) DEFAULT 0,
    uzcard_refund NUMERIC(12, 2) DEFAULT 0,
    humo_refund NUMERIC(12, 2) DEFAULT 0,
    other_payments NUMERIC(12, 2) DEFAULT 0,
    debt_payments NUMERIC(12, 2) DEFAULT 0,
    debt_refunds NUMERIC(12, 2) DEFAULT 0,
    report_data JSONB,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Images
CREATE TABLE IF NOT EXISTS images (
    id BIGSERIAL PRIMARY KEY,
    report_id BIGINT REFERENCES reports(id),
    shift_id BIGINT REFERENCES shifts(id),
    image_url VARCHAR(500),
    image_type VARCHAR(40) NOT NULL CHECK (
      image_type IN (
        'workplace_status',
        'terminal_power',
        'zero_report',
        'opening_notification',
        'receipt_roll',
        'uzcard_payment',
        'humo_payment'
      )
    ),
    uploaded_at TIMESTAMPTZ DEFAULT NOW()
);

-- Approval requests
CREATE TABLE IF NOT EXISTS approval_requests (
    id BIGSERIAL PRIMARY KEY,
    telegram_id BIGINT NOT NULL,
    first_name VARCHAR(255) NOT NULL,
    last_name VARCHAR(255),
    phone_number VARCHAR(20),
    role VARCHAR(20) NOT NULL CHECK (role = 'cashier'),
    password_hash VARCHAR(255),
    status VARCHAR(20) DEFAULT 'pending' CHECK (status IN ('pending', 'approved', 'rejected')),
    requested_at TIMESTAMPTZ DEFAULT NOW(),
    approved_at TIMESTAMPTZ NULL
);

-- Bot settings
CREATE TABLE IF NOT EXISTS bot_settings (
    id SMALLINT PRIMARY KEY,
    group_chat_id BIGINT,
    updated_by BIGINT,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

INSERT INTO bot_settings (id, group_chat_id, updated_by)
VALUES (1, NULL, NULL)
ON CONFLICT (id) DO NOTHING;

-- Useful indexes
CREATE INDEX IF NOT EXISTS idx_shifts_user_open ON shifts(user_id, is_open);
CREATE INDEX IF NOT EXISTS idx_shifts_opened_at ON shifts(opened_at DESC);
CREATE INDEX IF NOT EXISTS idx_shifts_user_opened_at ON shifts(user_id, opened_at DESC);
CREATE INDEX IF NOT EXISTS idx_reports_shift_type ON reports(shift_id, report_type);
CREATE INDEX IF NOT EXISTS idx_reports_shift_type_latest ON reports(shift_id, report_type, id DESC);
CREATE INDEX IF NOT EXISTS idx_reports_type_created_at ON reports(report_type, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_images_shift_type ON images(shift_id, image_type);
CREATE INDEX IF NOT EXISTS idx_approval_status ON approval_requests(status);
CREATE INDEX IF NOT EXISTS idx_approval_requests_tg_status ON approval_requests(telegram_id, status, requested_at DESC);
CREATE INDEX IF NOT EXISTS idx_users_role_active ON users(role, is_active);
