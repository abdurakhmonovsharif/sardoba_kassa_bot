from datetime import date, datetime
import time

import pytest

from db_config import DatabaseConnection


class FakeTransaction:
    def __init__(self):
        self.entered = False

    async def __aenter__(self):
        self.entered = True
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False


class FakeConnection:
    def __init__(self, row=None, rows=None):
        self.row = row
        self.rows = list(rows or [])
        self.fetchrow_calls = []
        self.fetch_calls = []
        self.execute_calls = []
        self.transaction_manager = FakeTransaction()

    async def fetchrow(self, query, *args):
        self.fetchrow_calls.append((query, args))
        return self.row

    async def fetch(self, query, *args):
        self.fetch_calls.append((query, args))
        return self.rows

    async def execute(self, query, *args):
        self.execute_calls.append((query, args))
        return "EXECUTED"

    def transaction(self):
        return self.transaction_manager


class FakeAcquire:
    def __init__(self, conn):
        self.conn = conn

    async def __aenter__(self):
        return self.conn

    async def __aexit__(self, exc_type, exc, tb):
        return False


class FakePool:
    def __init__(self, conn):
        self.conn = conn
        self.closed = False

    def acquire(self):
        return FakeAcquire(self.conn)

    async def close(self):
        self.closed = True


def test_normalize_database_url():
    assert DatabaseConnection._normalize_database_url("") == ""
    assert (
        DatabaseConnection._normalize_database_url("postgresql+asyncpg://u:p@localhost/db")
        == "postgresql://u:p@localhost/db"
    )


def test_prepare_query_supports_positional_and_named_params():
    db = DatabaseConnection()
    query, args = db._prepare_query("SELECT * FROM t WHERE a=%s AND b=%s", (1, 2))
    assert query == "SELECT * FROM t WHERE a=$1 AND b=$2"
    assert args == [1, 2]

    query, args = db._prepare_query(
        "UPDATE t SET a=%(alpha)s, b=%(beta)s WHERE c=%(alpha)s",
        {"alpha": 7, "beta": 9},
    )
    assert query == "UPDATE t SET a=$1, b=$2 WHERE c=$1"
    assert args == [7, 9]

    query, args = db._prepare_query(
        "SELECT * FROM t WHERE opened_at >= %s AND day = %s",
        ("2026-03-15 00:00:00", "2026-03-15"),
    )
    assert isinstance(args[0], datetime)
    assert isinstance(args[1], date)


@pytest.mark.asyncio
async def test_connect_creates_pool(monkeypatch):
    captured = {}

    async def fake_create_pool(**kwargs):
        captured.update(kwargs)
        return object()

    monkeypatch.setattr("db_config.asyncpg.create_pool", fake_create_pool)
    db = DatabaseConnection()
    db.database_url = "postgresql://user:pass@localhost/db"
    assert await db.connect() is True
    assert captured["dsn"] == "postgresql://user:pass@localhost/db"
    assert captured["min_size"] == db.minconn
    assert captured["max_size"] == db.maxconn
    assert captured["server_settings"]["timezone"] == "Asia/Tashkent"


@pytest.mark.asyncio
async def test_fetch_one_uses_pool_connection():
    conn = FakeConnection(row={"id": 1})
    db = DatabaseConnection()
    db.pool = FakePool(conn)

    result = await db.fetch_one("SELECT * FROM test WHERE id=%s", (1,))

    assert result == {"id": 1}
    assert conn.fetchrow_calls == [("SELECT * FROM test WHERE id=$1", (1,))]


@pytest.mark.asyncio
async def test_execute_query_uses_transaction():
    conn = FakeConnection()
    db = DatabaseConnection()
    db.pool = FakePool(conn)

    ok = await db.execute_query("UPDATE test SET a=%s", (1,))

    assert ok is True
    assert conn.transaction_manager.entered is True
    assert conn.execute_calls == [("UPDATE test SET a=$1", (1,))]


def test_log_query_timing_logs_warning(caplog):
    db = DatabaseConnection()
    with caplog.at_level("WARNING"):
        db._log_query_timing("SELECT * FROM test", time.perf_counter() - 1)
    assert "Slow query" in caplog.text
