"""Юнит- и fake-COM тесты чтения журнала регистрации (get_registration_log).

Не требуют 1С: чистая валидация и нормализация из src/event_log.py, а также
COM-логика из mcp_server_1c построена через duck typing и гоняется на fake.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

SRC = str(Path(__file__).resolve().parent.parent / "src")
sys.path.insert(0, SRC)

from event_log import (
    DEFAULT_DOCUMENT_EVENT_CODES,
    EventLogValidationError,
    build_event_log_filter,
    classify_event_log_com_error,
    filter_fingerprint,
    is_successful_status,
    make_common_filter,
    normalize_row,
    parse_local_datetime,
    transaction_status_code,
    validate_event_log_request,
    validate_metadata_path,
)

DOC_TYPE = "Документ.ПоступлениеДопРасходов"
UUID_A = "00000000-0000-0000-0000-000000000001"


def _ok(**overrides):
    base = {
        "start_datetime": "2026-05-29T00:00:00",
        "end_datetime": "2026-08-17T23:59:59",
        "events": ["_$Data$_.Post", "_$Data$_.Unpost", "_$Data$_.Update"],
        "metadata_path": DOC_TYPE,
        "data_refs": [{"_ref": UUID_A, "_type": DOC_TYPE}],
        "limit": 200,
    }
    base.update(overrides)
    return validate_event_log_request(**base)


# ---------------------------------------------------------------------------
# Даты
# ---------------------------------------------------------------------------

class TestDatetime:
    def test_valid_iso(self):
        d = parse_local_datetime("2026-05-29T11:02:21", "f")
        assert d.year == 2026 and d.month == 5 and d.day == 29
        assert d.hour == 11 and d.second == 21

    def test_reject_z(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("2026-05-29T11:02:21Z", "f")
        assert e.value.code == "TIMEZONE_NOT_SUPPORTED"

    def test_reject_offset(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("2026-05-29T11:02:21+05:00", "f")
        assert e.value.code == "TIMEZONE_NOT_SUPPORTED"

    def test_reject_offset_without_colon(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("2026-05-29T11:02:21+0500", "f")
        assert e.value.code == "TIMEZONE_NOT_SUPPORTED"

    def test_reject_offset_short(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("2026-05-29T11:02:21+05", "f")
        assert e.value.code == "TIMEZONE_NOT_SUPPORTED"

    def test_reject_lowercase_z(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("2026-05-29T11:02:21z", "f")
        assert e.value.code == "TIMEZONE_NOT_SUPPORTED"

    def test_reject_garbage(self):
        with pytest.raises(EventLogValidationError) as e:
            parse_local_datetime("not-a-date", "f")
        assert e.value.code == "INVALID_DATETIME"

    def test_end_before_start(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(end_datetime="2026-05-28T00:00:00")
        assert e.value.code == "INVALID_PERIOD"


# ---------------------------------------------------------------------------
# Период и лимит
# ---------------------------------------------------------------------------

class TestPeriodAndLimit:
    def test_object_request_allows_93_days(self):
        req = _ok()
        assert req.limit == 200

    def test_broad_range_limited(self):
        # только metadata+events (без refs) -> 7 дней максимум
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-01-01T00:00:00",
                end_datetime="2026-02-01T00:00:00",
                events=["_$Data$_.Update"],
                metadata_path=DOC_TYPE,
            )
        assert e.value.code == "DATE_RANGE_TOO_LARGE"

    def test_object_range_too_large(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(end_datetime="2026-12-31T00:00:00")
        assert e.value.code == "DATE_RANGE_TOO_LARGE"

    def test_limit_zero_rejected(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(limit=0)
        assert e.value.code == "INVALID_LIMIT"

    def test_limit_string_rejected(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(limit="abc")
        assert e.value.code == "INVALID_LIMIT"

    def test_limit_above_hard_cap_capped(self):
        req = _ok(limit=99999)
        assert req.limit == 1000


# ---------------------------------------------------------------------------
# Селективность
# ---------------------------------------------------------------------------

class TestSelectivity:
    def test_unfiltered_request_rejected(self):
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                events=None,
                metadata_path=None,
                data_refs=None,
                document_numbers=None,
                limit=200,
            )
        assert e.value.code == "UNSAFE_UNFILTERED_REQUEST"

    def test_metadata_plus_events_allowed(self):
        req = validate_event_log_request(
            start_datetime="2026-05-29T00:00:00",
            end_datetime="2026-05-30T00:00:00",
            events=["_$Data$_.Unpost"],
            metadata_path=DOC_TYPE,
        )
        assert req.metadata_path == DOC_TYPE

    def test_default_events_applied_for_object_filter(self):
        # events=None + data_refs -> применяются Post/Unpost/Update
        req = validate_event_log_request(
            start_datetime="2026-05-29T00:00:00",
            end_datetime="2026-08-17T23:59:59",
            events=None,
            metadata_path=DOC_TYPE,
            data_refs=[{"_ref": UUID_A, "_type": DOC_TYPE}],
        )
        assert req.events == DEFAULT_DOCUMENT_EVENT_CODES

    def test_default_events_applied_for_document_numbers(self):
        req = validate_event_log_request(
            start_datetime="2026-05-29T00:00:00",
            end_datetime="2026-08-17T23:59:59",
            events=None,
            metadata_path=DOC_TYPE,
            document_numbers=["ААСГ0000070"],
            document_date_from="2026-05-29T00:00:00",
            document_date_to="2026-05-30T00:00:00",
        )
        assert req.events == DEFAULT_DOCUMENT_EVENT_CODES

    def test_metadata_without_events_still_rejected(self):
        # metadata_path без events и без объекта -> всё ещё неселективно
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                events=None,
                metadata_path=DOC_TYPE,
                data_refs=None,
                document_numbers=None,
            )
        assert e.value.code == "UNSAFE_UNFILTERED_REQUEST"


# ---------------------------------------------------------------------------
# metadata_path и номера документов
# ---------------------------------------------------------------------------

class TestMetadataAndNumbers:
    def test_valid_metadata_path(self):
        validate_metadata_path(DOC_TYPE)

    @pytest.mark.parametrize(
        "bad",
        ["Документ.", "Документ.Имя СПробелом", "Документ.Имя;DROP",
         "Справочник", "Что-То.Имя", "Документ.Имя.Лишнее"],
    )
    def test_invalid_metadata_path(self, bad):
        with pytest.raises(EventLogValidationError) as e:
            validate_metadata_path(bad)
        assert e.value.code == "INVALID_METADATA_PATH"

    def test_too_many_documents(self):
        nums = [f"ААСГ{n:07d}" for n in range(21)]
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                metadata_path=DOC_TYPE,
                document_numbers=nums,
                document_date_from="2026-05-29T00:00:00",
                document_date_to="2026-05-30T00:00:00",
            )
        assert e.value.code == "TOO_MANY_DOCUMENTS"

    def test_duplicate_numbers_rejected(self):
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                metadata_path=DOC_TYPE,
                document_numbers=["ААСГ0000070", "ААСГ0000070"],
                document_date_from="2026-05-29T00:00:00",
                document_date_to="2026-05-30T00:00:00",
            )
        assert e.value.code == "DUPLICATE_DOCUMENT_NUMBER"

    def test_document_numbers_require_date_range(self):
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                metadata_path=DOC_TYPE,
                document_numbers=["ААСГ0000070"],
            )
        assert e.value.code == "DOCUMENT_DATE_RANGE_REQUIRED"

    def test_document_numbers_require_document_type(self):
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                metadata_path="Справочник.Контрагенты",
                document_numbers=["ААСГ0000070"],
                document_date_from="2026-05-29T00:00:00",
                document_date_to="2026-05-30T00:00:00",
            )
        assert e.value.code == "INVALID_METADATA_PATH"


# ---------------------------------------------------------------------------
# data_refs
# ---------------------------------------------------------------------------

class TestDataRefs:
    def test_invalid_ref_type_mismatch(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(data_refs=[{"_ref": UUID_A, "_type": "Документ.Другой"}])
        assert e.value.code == "INVALID_DATA_REF"

    def test_invalid_uuid(self):
        with pytest.raises(EventLogValidationError) as e:
            _ok(data_refs=[{"_ref": "не-uuid", "_type": DOC_TYPE}])
        assert e.value.code == "INVALID_DATA_REF"

    def test_data_refs_and_document_numbers_conflict(self):
        with pytest.raises(EventLogValidationError) as e:
            validate_event_log_request(
                start_datetime="2026-05-29T00:00:00",
                end_datetime="2026-05-30T00:00:00",
                metadata_path=DOC_TYPE,
                data_refs=[{"_ref": UUID_A, "_type": DOC_TYPE}],
                document_numbers=["ААСГ0000070"],
                document_date_from="2026-05-29T00:00:00",
                document_date_to="2026-05-30T00:00:00",
            )
        assert e.value.code == "INVALID_FILTER_COMBINATION"


# ---------------------------------------------------------------------------
# fingerprint
# ---------------------------------------------------------------------------

class TestFingerprint:
    def test_stable(self):
        a = filter_fingerprint({"x": 1, "y": [1, 2]})
        b = filter_fingerprint({"y": [1, 2], "x": 1})
        assert a == b

    def test_does_not_reveal_values(self):
        secret = "ААСГ0000070"
        fp = filter_fingerprint({"number": secret})
        assert secret not in fp
        assert len(fp) == 16


# ---------------------------------------------------------------------------
# Нормализация статусов
# ---------------------------------------------------------------------------

class TestNormalization:
    def test_transaction_status(self):
        assert transaction_status_code("Зафиксирована") == "committed"
        assert transaction_status_code("Отменена") == "rolled_back"
        assert transaction_status_code("НетТранзакции") == "not_applicable"
        assert transaction_status_code("НеЗавершена") == "unfinished"
        assert transaction_status_code("ЧтоТо") == "unknown"

    def test_successful(self):
        assert is_successful_status("Зафиксирована") is True
        assert is_successful_status("НетТранзакции") is True
        assert is_successful_status("Отменена") is False

    def test_normalize_empty(self):
        assert normalize_row({}) == {}

    def test_normalize_status_adds_triple(self):
        out = normalize_row({"transaction_status": "Отменена"})
        assert out["transaction_status"] == "Отменена"
        assert out["transaction_status_code"] == "rolled_back"
        assert out["successful"] is False

    def test_unknown_field_ignored(self):
        out = normalize_row({"неизвестное": 1})
        assert out == {}


# ---------------------------------------------------------------------------
# Fake-COM тесты функций mcp_server_1c
# ---------------------------------------------------------------------------

class FakeOnec:
    """Минимальный duck-typed fake соединения/объекта 1С."""

    def __init__(self, name):
        self._name = name
        self._dict = {}

    def Вставить(self, key, value):
        self._dict[key] = value

    def Добавить(self, value):
        self._list.append(value)

    def __getitem__(self, key):
        return self._dict[key]

    @property
    def Имя(self):
        return self._name


class FakeConn:
    def __init__(self):
        self.created = []
        self.last_export_args = None
        self.СтатусТранзакцииЗаписиЖурналаРегистрации = FakeStatuses()

    def NewObject(self, kind, *args):
        self.created.append(kind)
        if kind == "Массив":
            obj = FakeOnec("Массив")
            obj._list = []
            return obj
        return FakeOnec(kind)


class FakeStatuses:
    Зафиксирована = "Зафиксирована"
    Отменена = "Отменена"
    НеЗавершена = "НеЗавершена"
    НетТранзакции = "НетТранзакции"


FAKE_REQUEST = validate_event_log_request(
    start_datetime="2026-05-29T00:00:00",
    end_datetime="2026-08-17T23:59:59",
    events=["_$Data$_.Unpost"],
    metadata_path=DOC_TYPE,
    data_refs=[{"_ref": UUID_A, "_type": DOC_TYPE},
               {"_ref": "00000000-0000-0000-0000-000000000002", "_type": DOC_TYPE}],
    limit=200,
)


class TestFakeCom:
    def test_two_refs_two_or_structures(self):
        conn = FakeConn()
        refs = [object(), object()]
        result = build_event_log_filter(conn, FAKE_REQUEST, refs)
        # верхний уровень — массив из двух структур
        assert result._name == "Массив"
        assert len(result._list) == 2
        for item in result._list:
            assert item._dict["Данные"] is not None
            # общие условия одинаковые
            assert item._dict["ДатаНачала"] == FAKE_REQUEST.start_datetime

    def test_successful_only_adds_statuses(self):
        conn = FakeConn()
        item = make_common_filter(conn, FAKE_REQUEST)
        assert "СтатусТранзакции" in item._dict
        statuses = item._dict["СтатусТранзакции"]
        assert statuses._list == ["Зафиксирована", "НетТранзакции"]

    def test_no_successful_filter(self):
        conn = FakeConn()
        req = validate_event_log_request(
            start_datetime="2026-05-29T00:00:00",
            end_datetime="2026-05-30T00:00:00",
            events=["_$Data$_.Unpost"],
            metadata_path=DOC_TYPE,
            successful_only=False,
        )
        item = make_common_filter(conn, req)
        assert "СтатусТранзакции" not in item._dict

    def test_no_write_methods_called(self):
        conn = FakeConn()
        build_event_log_filter(conn, FAKE_REQUEST, [object()])
        # не должно быть вызовов записи журнала
        assert not hasattr(conn, "ЗаписьЖурналаРегистрации")


# ---------------------------------------------------------------------------
# Fake-COM тест сериализации таблицы журнала (полный путь, без 1С)
# ---------------------------------------------------------------------------

class TestSerializeTable:
    def test_limit_keeps_newest_rows(self):
        """При переполнении сохраняются НАИБОЛЕЕ НОВЫЕ limit строк.

        mcp_server_1c на Windows при импорте переопределяет sys.stdout, что
        ломает pytest capture, поэтому проверяем в подпроцессе (как это делает
        tests/test_http_smoke.py для всего сервера).
        """
        import subprocess

        script = '''
import sys
sys.path.insert(0, "SRC")
import mcp_server_1c as server

class Cell:
    def __init__(self, values): self._v = values
    def Get(self, i): return self._v[i]

class Col:
    def __init__(self, name): self.Имя = name

class Table:
    def __init__(self, cols, rows):
        self.Колонки = [Col(c) for c in cols]
        self._rows = rows
    def __iter__(self):
        for r in self._rows:
            yield Cell(r)

class Statuses:
    Зафиксирована = "Зафиксирована"; НетТранзакции = "НетТранзакции"
    Отменена = "Отменена"; НеЗавершена = "НеЗавершена"

class Conn:
    СтатусТранзакцииЗаписиЖурналаРегистрации = Statuses()
    def XMLСтрока(self, v): return "fake-uuid"

conn = Conn()
cols = ["Дата", "Событие", "Данные", "СтатусТранзакции"]
rows = [
    ["2026-05-30T10:00:00", "_$Data$_.Post", "doc-uuid", "Зафиксирована"],
    ["2026-06-01T10:00:00", "_$Data$_.Update", "doc-uuid", "Зафиксирована"],
    ["2026-07-09T13:11:44", "_$Data$_.Unpost", "doc-uuid", "Зафиксирована"],
]
result, truncated = server.serialize_event_log_table(Table(cols, rows), 2, conn)
print("TRUNCATED", truncated)
print("COUNT", len(result))
for r in result:
    print("EVENT", r["event"])
'''
        # экранируем обратные слэши, иначе \U в Windows-пути сломает unicode-escape
        safe_src = str(SRC).replace("\\", "\\\\")
        script = script.replace("SRC", safe_src)
        proc = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=30, check=False,
        )
        assert proc.returncode == 0, proc.stderr
        out = proc.stdout
        assert "TRUNCATED True" in out
        assert "COUNT 2" in out
        assert "_$Data$_.Unpost" in out
        assert "_$Data$_.Update" in out
        assert "_$Data$_.Post" not in out


class TestClassifier:
    def test_access_denied_explicit(self):
        assert classify_event_log_com_error("Недостаточно прав на журнал регистрации") == "EVENT_LOG_ACCESS_DENIED"
        assert classify_event_log_com_error("Недостаточно прав") == "EVENT_LOG_ACCESS_DENIED"
        assert classify_event_log_com_error("Violation insufficient rights") == "EVENT_LOG_ACCESS_DENIED"

    def test_other_error_not_access_denied(self):
        assert classify_event_log_com_error("Нет доступа к файлу журнала (файл занят)") == "ONEC_ERROR"

    def test_unavailable(self):
        assert classify_event_log_com_error("Метод не поддерживается этой версией платформы") == "EVENT_LOG_UNAVAILABLE"

    def test_language_unsupported(self):
        assert classify_event_log_com_error("Колонка не найдена") == "EVENT_LOG_LANGUAGE_UNSUPPORTED"
