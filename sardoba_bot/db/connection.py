import asyncio
import logging
import os
import re
import time
from datetime import date, datetime

import asyncpg
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

logger = logging.getLogger(__name__)
TASHKENT_TIMEZONE = "Asia/Tashkent"


class DatabaseConnection:
    _named_param_pattern = re.compile(r"%\(([^)]+)\)s")
    _date_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}$")
    _datetime_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}$")

    def __init__(self):
        self.database_url = self._normalize_database_url(os.getenv("DATABASE_URL", ""))
        self.host = os.getenv("DB_HOST", "localhost")
        self.database = os.getenv("DB_NAME", "sardoba_bot")
        self.user = os.getenv("DB_USER", "root")
        self.password = os.getenv("DB_PASSWORD", "")
        self.port = int(os.getenv("DB_PORT", 5432))
        self.pool = None
        self.minconn = max(1, int(os.getenv("DB_POOL_MIN", 1)))
        self.maxconn = max(self.minconn, int(os.getenv("DB_POOL_MAX", 8)))
        self.connect_timeout = int(os.getenv("DB_CONNECT_TIMEOUT", 5))
        self.command_timeout = float(os.getenv("DB_COMMAND_TIMEOUT", 60))
        self.slow_query_ms = float(os.getenv("DB_SLOW_QUERY_MS", 300))

    @staticmethod
    def _normalize_database_url(database_url):
        if not database_url:
            return ""
        if database_url.startswith("postgresql+asyncpg://"):
            return "postgresql://" + database_url.split("://", 1)[1]
        return database_url

    def _pool_kwargs(self):
        if self.database_url:
            return {
                "dsn": self.database_url,
                "timeout": self.connect_timeout,
                "command_timeout": self.command_timeout,
                "server_settings": {"timezone": TASHKENT_TIMEZONE},
            }
        return {
            "host": self.host,
            "database": self.database,
            "user": self.user,
            "password": self.password,
            "port": self.port,
            "timeout": self.connect_timeout,
            "command_timeout": self.command_timeout,
            "server_settings": {"timezone": TASHKENT_TIMEZONE},
        }

    @staticmethod
    def _restore_percent_literals(query):
        return query.replace("__PERCENT_LITERAL__", "%")

    def _coerce_param(self, value):
        if not isinstance(value, str):
            return value
        if self._datetime_pattern.match(value):
            try:
                return datetime.fromisoformat(value.replace(" ", "T"))
            except ValueError:
                return value
        if self._date_pattern.match(value):
            try:
                return date.fromisoformat(value)
            except ValueError:
                return value
        return value

    def _prepare_query(self, query, params=None):
        query = (query or "").replace("%%", "__PERCENT_LITERAL__")

        if params is None:
            return self._restore_percent_literals(query), []

        if isinstance(params, dict):
            order = []
            indexes = {}

            def repl(match):
                key = match.group(1)
                if key not in indexes:
                    indexes[key] = len(order) + 1
                    order.append(self._coerce_param(params[key]))
                return f"${indexes[key]}"

            prepared_query = self._named_param_pattern.sub(repl, query)
            return self._restore_percent_literals(prepared_query), order

        if not isinstance(params, (list, tuple)):
            params = [params]

        idx = 0

        def repl(match):
            nonlocal idx
            idx += 1
            return f"${idx}"

        prepared_query = re.sub(r"%s", repl, query)
        return self._restore_percent_literals(prepared_query), [self._coerce_param(value) for value in params]

    async def connect(self):
        """Establish database connection pool."""
        try:
            if self.pool:
                return True
            self.pool = await asyncpg.create_pool(
                min_size=self.minconn,
                max_size=self.maxconn,
                **self._pool_kwargs(),
            )
            print(
                f"Successfully connected to PostgreSQL database "
                f"(pool {self.minconn}-{self.maxconn})"
            )
            return True
        except Exception as e:
            print(f"Error connecting to PostgreSQL: {e}")
            self.pool = None
            return False

    async def _ensure_connection(self):
        """Ensure there is an active DB connection (best-effort)."""
        if self.pool:
            return True
        return bool(await self.connect())

    async def disconnect(self):
        """Close database connection pool."""
        if self.pool:
            await self.pool.close()
            self.pool = None
            print("PostgreSQL connection closed")

    def _log_query_timing(self, query, started_at):
        elapsed_ms = (time.perf_counter() - started_at) * 1000
        if elapsed_ms < self.slow_query_ms:
            return
        compact_query = " ".join((query or "").split())
        if len(compact_query) > 240:
            compact_query = compact_query[:237] + "..."
        logger.warning("Slow query %.1f ms: %s", elapsed_ms, compact_query)

    async def _execute(self, query, params=None, fetch_mode=None):
        prepared_query, args = self._prepare_query(query, params)
        started_at = time.perf_counter()
        try:
            if not await self._ensure_connection():
                return [] if fetch_mode == "all" else None
            async with self.pool.acquire() as conn:
                if fetch_mode == "one":
                    row = await conn.fetchrow(prepared_query, *args)
                    return dict(row) if row else None
                if fetch_mode == "all":
                    rows = await conn.fetch(prepared_query, *args)
                    return [dict(row) for row in rows]
                async with conn.transaction():
                    await conn.execute(prepared_query, *args)
                return True
        finally:
            self._log_query_timing(prepared_query, started_at)

    async def execute_query(self, query, params=None):
        """Execute a query that doesn't return results."""
        for attempt in range(2):
            try:
                return bool(await self._execute(query, params=params, fetch_mode=None))
            except Exception as e:
                print(f"Error executing query: {e}")
                if attempt == 0:
                    try:
                        await self.disconnect()
                    except Exception:
                        pass
                    continue
                return False

    async def fetch_one(self, query, params=None):
        """Fetch one record."""
        for attempt in range(2):
            try:
                return await self._execute(query, params=params, fetch_mode="one")
            except Exception as e:
                print(f"Error fetching record: {e}")
                if attempt == 0:
                    try:
                        await self.disconnect()
                    except Exception:
                        pass
                    continue
                return None

    async def fetch_all(self, query, params=None):
        """Fetch all records."""
        for attempt in range(2):
            try:
                rows = await self._execute(query, params=params, fetch_mode="all")
                return rows or []
            except Exception as e:
                print(f"Error fetching records: {e}")
                if attempt == 0:
                    try:
                        await self.disconnect()
                    except Exception:
                        pass
                    continue
                return []

    def _run_sync(self, coro):
        try:
            asyncio.get_running_loop()
        except RuntimeError:
            return asyncio.run(coro)
        raise RuntimeError("Synchronous DB access is not allowed inside a running event loop")

    def connect_sync(self):
        return self._run_sync(self.connect())

    def disconnect_sync(self):
        return self._run_sync(self.disconnect())

    def execute_query_sync(self, query, params=None):
        return self._run_sync(self.execute_query(query, params))

    def fetch_one_sync(self, query, params=None):
        return self._run_sync(self.fetch_one(query, params))

    def fetch_all_sync(self, query, params=None):
        return self._run_sync(self.fetch_all(query, params))
