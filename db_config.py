import os
from dotenv import load_dotenv
import psycopg2
from psycopg2.extras import RealDictCursor

# Load environment variables
load_dotenv()

class DatabaseConnection:
    def __init__(self):
        self.database_url = self._normalize_database_url(os.getenv('DATABASE_URL', ''))
        self.host = os.getenv('DB_HOST', 'localhost')
        self.database = os.getenv('DB_NAME', 'sardoba_bot')
        self.user = os.getenv('DB_USER', 'root')
        self.password = os.getenv('DB_PASSWORD', '')
        self.port = int(os.getenv('DB_PORT', 5432))
        self.connection = None

    @staticmethod
    def _normalize_database_url(database_url):
        if not database_url:
            return ''
        if database_url.startswith('postgresql+asyncpg://'):
            return 'postgresql://' + database_url.split('://', 1)[1]
        return database_url

    def connect(self):
        """Establish database connection"""
        try:
            if self.database_url:
                self.connection = psycopg2.connect(self.database_url)
            else:
                self.connection = psycopg2.connect(
                    host=self.host,
                    dbname=self.database,
                    user=self.user,
                    password=self.password,
                    port=self.port
                )
            self.connection.autocommit = False
            print("Successfully connected to PostgreSQL database")
            return True
        except psycopg2.Error as e:
            print(f"Error connecting to PostgreSQL: {e}")
            return False
        except Exception as e:
            print(f"Unexpected error connecting to PostgreSQL: {e}")
            return False

    def _ensure_connection(self):
        """Ensure there is an active DB connection (best-effort)."""
        try:
            if self.connection and self.connection.closed == 0:
                return True
        except Exception:
            # connection object may be partially initialized
            pass
        return bool(self.connect())

    def disconnect(self):
        """Close database connection"""
        if self.connection and self.connection.closed == 0:
            self.connection.close()
            print("PostgreSQL connection closed")

    def execute_query(self, query, params=None):
        """Execute a query that doesn't return results"""
        # Retry once after reconnect in case DB was restarted.
        for attempt in range(2):
            try:
                if not self._ensure_connection():
                    return False
                cursor = self.connection.cursor()
                cursor.execute(query, params)
                self.connection.commit()
                cursor.close()
                return True
            except psycopg2.Error as e:
                print(f"Error executing query: {e}")
                try:
                    self.connection.rollback()
                except Exception:
                    pass
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return False
            except Exception as e:
                print(f"Unexpected error executing query: {e}")
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return False

    def fetch_one(self, query, params=None):
        """Fetch one record"""
        for attempt in range(2):
            try:
                if not self._ensure_connection():
                    return None
                cursor = self.connection.cursor(cursor_factory=RealDictCursor)
                cursor.execute(query, params)
                result = cursor.fetchone()
                cursor.close()
                return result
            except psycopg2.Error as e:
                print(f"Error fetching record: {e}")
                try:
                    self.connection.rollback()
                except Exception:
                    pass
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return None
            except Exception as e:
                print(f"Unexpected error fetching record: {e}")
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return None

    def fetch_all(self, query, params=None):
        """Fetch all records"""
        for attempt in range(2):
            try:
                if not self._ensure_connection():
                    return []
                cursor = self.connection.cursor(cursor_factory=RealDictCursor)
                cursor.execute(query, params)
                result = cursor.fetchall()
                cursor.close()
                return result
            except psycopg2.Error as e:
                print(f"Error fetching records: {e}")
                try:
                    self.connection.rollback()
                except Exception:
                    pass
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return []
            except Exception as e:
                print(f"Unexpected error fetching records: {e}")
                if attempt == 0:
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    continue
                return []
