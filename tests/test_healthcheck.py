"""Юнит-тесты healthcheck.py: валидация ответа list_databases.

healthcheck.py импортирует mcp лениво (внутри run), поэтому чистые функции
можно тестировать без запущенного сервера.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from healthcheck import check_databases_payload


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
