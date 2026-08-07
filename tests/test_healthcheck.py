"""Юнит-тесты healthcheck.py: валидация ответа list_databases.

healthcheck.py импортирует mcp лениво (внутри run), поэтому чистые функции
можно тестировать без запущенного сервера.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from healthcheck import (
    _probe_com,
    check_databases_payload,
    redact_secrets,
)


class _FakeResult:
    is_error = False

    def __init__(self, text):
        self.content = [type("_C", (), {"text": text})()]


class _FakeClient:
    """call_tool возвращает текст ответа либо бросает исключение."""

    def __init__(self, text=None, exc=None):
        self._text = text
        self._exc = exc

    async def call_tool(self, *args, **kwargs):
        if self._exc is not None:
            raise self._exc
        return _FakeResult(self._text)


def test_com_error_detail_redacts_secrets():
    """HEALTH_COM_DETAIL для error-ответа: токен и Pwd вырезаются."""
    import asyncio
    import contextlib
    import io
    import json

    err_text = 'Ошибка: Pwd="secret"; URL=/mcp/FAKETOKEN123'
    client = _FakeClient(text=json.dumps({"error": err_text}, ensure_ascii=False))
    stderr = io.StringIO()
    with contextlib.redirect_stderr(stderr):
        failures = asyncio.run(_probe_com(client, ["buh"], "FAKETOKEN123"))
    out = stderr.getvalue()
    assert failures == ["buh"]
    assert "FAKETOKEN123" not in out
    assert "secret" not in out
    assert "/mcp/<token>" in out
    # ascii_json экранирует кавычки внутри JSON-строки
    assert 'Pwd=\\"***\\"' in out


def test_com_debug_traceback_redacts_secrets(monkeypatch):
    """Debug-traceback COM-пробы (ONEC_HEALTHCHECK_DEBUG=1) редактируется."""
    import asyncio
    import contextlib
    import io

    monkeypatch.setenv("ONEC_HEALTHCHECK_DEBUG", "1")

    class Boom(Exception):
        pass

    client = _FakeClient(
        exc=Boom('connect url http://127.0.0.1:8000/mcp/FAKETOKEN123')
    )
    stderr = io.StringIO()
    with contextlib.redirect_stderr(stderr):
        asyncio.run(_probe_com(client, ["buh"], "FAKETOKEN123"))
    out = stderr.getvalue()
    assert "FAKETOKEN123" not in out
    assert "/mcp/<token>" in out
    assert "HEALTH_COM_DETAIL" in out


def test_redact_secrets():
    """Редокция: прямой токен, любой /mcp/<token>, Pwd="..."."""
    secret = "SECRETTOKEN"
    text = (
        f"http://127.0.0.1:8000/mcp/{secret} "
        'Pwd="секрет1" Pwd = "x"'
    )
    redacted = redact_secrets(text, secret)
    assert secret not in redacted
    assert "/mcp/<token>" in redacted
    assert 'Pwd="***"' in redacted
    # любой токен в /mcp/<...> редактируется даже без прямого вхождения
    other = redact_secrets("url /mcp/OtherToken123 x", "NOT_IT")
    assert other == "url /mcp/<token> x"


def test_payload_not_object():
    """Ответ list_databases не является объектом (null/[]/строка/число)."""
    for bad in (None, [], "text", 42):
        code, marker, databases = check_databases_payload(bad)
        assert code == 2
        assert marker == "HEALTH_DATABASES_INVALID"
        assert databases == []


def test_payload_ok():
    code, marker, databases = check_databases_payload(
        {"databases": {"UT10": {}, "Buh": {}}, "default_database": "UT10"}
    )
    assert code == 0 and marker == ""
    assert databases == ["Buh", "UT10"]


def test_payload_databases_not_object():
    code, marker, databases = check_databases_payload({"databases": "не объект"})
    assert code == 2
    assert marker == "HEALTH_DATABASES_INVALID"
    assert databases == []


def test_payload_databases_null():
    code, marker, _ = check_databases_payload({})
    assert code == 2
    assert marker == "HEALTH_DATABASES_INVALID"


def test_payload_empty_databases():
    code, marker, _ = check_databases_payload(
        {"databases": {}, "default_database": ""}
    )
    assert code == 2
    assert marker == "HEALTH_NO_DATABASES"


def test_payload_default_missing():
    code, marker, _ = check_databases_payload(
        {"databases": {"UT10": {}}, "default_database": "NOPE"}
    )
    assert code == 2
    assert marker == "HEALTH_DEFAULT_DATABASE_INVALID"


def test_payload_default_empty():
    """default_database отсутствует/пуст при непустых базах — тоже ошибка."""
    code, marker, _ = check_databases_payload(
        {"databases": {"UT10": {}}, "default_database": ""}
    )
    assert code == 2
    assert marker == "HEALTH_DEFAULT_DATABASE_INVALID"
