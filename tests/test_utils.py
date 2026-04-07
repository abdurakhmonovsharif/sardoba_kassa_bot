from datetime import datetime

from utils import (
    format_currency,
    get_current_timestamp,
    hash_password,
    validate_phone_number,
    verify_password,
)


def test_hash_and_verify_password_roundtrip():
    hashed = hash_password("secret123")
    assert len(hashed) == 96
    assert verify_password(hashed, "secret123") is True
    assert verify_password(hashed, "wrong") is False


def test_verify_password_supports_legacy_plain_text():
    assert verify_password("plain-pass", "plain-pass") is True
    assert verify_password("plain-pass", "other") is False
    assert verify_password("", "other") is False


def test_get_current_timestamp_format():
    value = get_current_timestamp()
    parsed = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    assert isinstance(parsed, datetime)


def test_format_currency():
    assert format_currency(1234.5) == "1,234.50"


def test_validate_phone_number():
    assert validate_phone_number("+998901234567") is True
    assert validate_phone_number("901234567") is True
    assert validate_phone_number("123") is False
