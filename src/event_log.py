"""
Чтение журнала регистрации 1С — чистые (без COM) константы, валидация и нормализация.

Модуль специально НЕ импортирует `pythoncom`/`win32com`: он тестируется на любой ОС
(в том числе на Linux CI) и покрыт юнит-тестами в tests/test_event_log.py.

COM-зависимая логика (создание объектов 1С, разрешение ссылок, вызов
`ВыгрузитьЖурналРегистрации()`) живёт в `mcp_server_1c.py` и спроектирована через
duck typing, чтобы её можно было гонять на fake-объектах.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import uuid
from dataclasses import dataclass
from typing import Any

__all__ = [
    "DEFAULT_DOCUMENT_EVENT_CODES",
    "EVENT_LOG_JSON_FIELDS",
    "EventLogAccessError",
    "EventLogValidationError",
    "ValidatedEventLogRequest",
    "build_event_log_filter",
    "filter_fingerprint",
    "is_successful_status",
    "make_common_filter",
    "normalize_row",
    "onec_array",
    "onec_transaction_statuses",
    "parse_local_datetime",
    "transaction_status_code",
    "validate_data_ref",
    "validate_document_number",
    "validate_event_code",
    "validate_event_log_request",
    "validate_metadata_path",
]


# События по умолчанию для объектного расследования (проведение / отмена / изменение).
DEFAULT_DOCUMENT_EVENT_CODES = (
    "_$Data$_.Post",
    "_$Data$_.Unpost",
    "_$Data$_.Update",
)

# Полное имя метаданных: тип.Имя, без пробелов, кавычек и операторов запроса.
_METADATA_PATH_RE = re.compile(
    r"^(Документ|Справочник|РегистрНакопления|РегистрСведений|"
    r"РегистрБухгалтерии|ПланВидовХарактеристик|ПланСчетов|"
    r"ПланВидовРасчета)\.[A-Za-zА-Яа-яЁё0-9_]+$"
)

# Управляющие символы, запрещённые в номерах документов.
_CONTROL_CHARS_RE = re.compile(r"[\x00-\x1f\x7f]")

# 1С -> JSON код статуса транзакции.
_TRANSACTION_STATUS_CODES = {
    "Зафиксирована": "committed",
    "НетТранзакции": "not_applicable",
    "Отменена": "rolled_back",
    "НеЗавершена": "unfinished",
}

# Статусы, которые считаются фактически выполненным действием.
SUCCESSFUL_TRANSACTION_STATUSES = ("Зафиксирована", "НетТранзакции")

# Стабильные англ. ключи ответа (фактически сериализуемые колонки).
EVENT_LOG_JSON_FIELDS = (
    "timestamp",
    "user_id",
    "user_name",
    "computer",
    "application",
    "application_presentation",
    "connection_number",
    "session",
    "event",
    "event_presentation",
    "metadata",
    "metadata_presentation",
    "data",
    "data_presentation",
    "comment",
    "transaction_status",
    "transaction_status_code",
    "successful",
    "transaction_id",
    "server_name",
    "port",
    "sync_port",
)


class EventLogValidationError(ValueError):
    """Ошибка валидации запроса журнала регистрации с машинным кодом."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


class EventLogAccessError(Exception):
    """Пользователь соединения не имеет права ЖурналРегистрации."""


def parse_local_datetime(value: str, field: str) -> dt.datetime:
    """Разбирает ISO datetime в локальном (без timezone) времени ИБ.

    Смещения `Z` и `±HH:MM` в v1 отклоняются: журнал регистрации показывает
    локальное время ИБ без явного смещения UTC.
    """
    if not isinstance(value, str) or not value.strip():
        raise EventLogValidationError("INVALID_DATETIME", f"{field}: ожидается ISO-дата")
    raw = value.strip()
    if raw.endswith("Z") or re.search(r"[+-]\d\d:\d\d$", raw):
        raise EventLogValidationError(
            "TIMEZONE_NOT_SUPPORTED",
            f"{field}: передайте локальное время информационной базы без timezone",
        )
    try:
        return dt.datetime.fromisoformat(raw)
    except ValueError as exc:
        raise EventLogValidationError(
            "INVALID_DATETIME", f"{field}: некорректная ISO-дата"
        ) from exc


def validate_metadata_path(path: str | None) -> None:
    """Проверяет формат полного имени метаданных."""
    if path is not None and not _METADATA_PATH_RE.fullmatch(path):
        raise EventLogValidationError(
            "INVALID_METADATA_PATH", "Некорректное полное имя метаданных"
        )


def validate_event_code(code: str) -> None:
    if not isinstance(code, str) or not code.strip():
        raise EventLogValidationError("INVALID_EVENT", "Пустой технический код события")
    if len(code) > 128:
        raise EventLogValidationError(
            "INVALID_EVENT", "Технический код события длиннее 128 символов"
        )


def validate_document_number(number: str) -> None:
    if not isinstance(number, str) or not number.strip():
        raise EventLogValidationError("INVALID_DOCUMENT_NUMBER", "Пустой номер документа")
    if _CONTROL_CHARS_RE.search(number):
        raise EventLogValidationError(
            "INVALID_DOCUMENT_NUMBER", "Номер документа содержит управляющие символы"
        )


def validate_data_ref(ref: Any, metadata_path: str | None) -> dict:
    """Проверяет один объект data_ref {"_ref": uuid, "_type": "..."}."""
    if not isinstance(ref, dict):
        raise EventLogValidationError("INVALID_DATA_REF", "Элемент data_refs должен быть объектом")
    uuid_str = ref.get("_ref")
    type_path = ref.get("_type")
    if not isinstance(uuid_str, str) or not isinstance(type_path, str):
        raise EventLogValidationError(
            "INVALID_DATA_REF",
            "Элемент data_refs должен содержать строковые '_ref' (uuid) и '_type'",
        )
    try:
        uuid.UUID(uuid_str)
    except (ValueError, AttributeError):
        raise EventLogValidationError("INVALID_DATA_REF", f"Некорректный UUID: {uuid_str!r}")
    validate_metadata_path(type_path)
    if metadata_path and type_path != metadata_path:
        raise EventLogValidationError(
            "INVALID_DATA_REF",
            f"Тип ссылки {type_path!r} не совпадает с metadata_path {metadata_path!r}",
        )
    return {"_ref": uuid_str, "_type": type_path}


def filter_fingerprint(payload: dict) -> str:
    """Стабильный SHA-256 fingerprint фильтра без раскрытия исходных значений."""
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8", "replace")).hexdigest()[:16]


@dataclass
class ValidatedEventLogRequest:
    """Результат чистой валидации параметров запроса журнала регистрации."""

    start_datetime: dt.datetime
    end_datetime: dt.datetime
    events: tuple[str, ...]
    metadata_path: str | None
    data_refs: tuple[dict, ...]
    document_numbers: tuple[str, ...]
    document_date_from: dt.datetime | None
    document_date_to: dt.datetime | None
    successful_only: bool
    order: str
    limit: int

    @property
    def has_object_filter(self) -> bool:
        return bool(self.data_refs) or bool(self.document_numbers)

    def period_payload(self) -> dict:
        return {
            "from": self.start_datetime.isoformat(),
            "to": self.end_datetime.isoformat(),
            "timezone_semantics": "infobase-local",
        }

    def audit_payload(self) -> dict:
        """Данные только для fingerprint: без номеров документов, UUID и текста."""
        return {
            "from": self.start_datetime.isoformat(),
            "to": self.end_datetime.isoformat(),
            "events": list(self.events),
            "metadata_path": self.metadata_path,
            "data_refs_count": len(self.data_refs),
            "document_numbers_count": len(self.document_numbers),
            "successful_only": self.successful_only,
            "order": self.order,
            "limit": self.limit,
        }


def validate_event_log_request(
    *,
    start_datetime: str,
    end_datetime: str,
    events: list[str] | None = None,
    metadata_path: str | None = None,
    data_refs: list[dict] | None = None,
    document_numbers: list[str] | None = None,
    document_date_from: str | None = None,
    document_date_to: str | None = None,
    successful_only: bool = True,
    order: str = "asc",
    limit: int = 200,
    max_range_days: int = 93,
    broad_range_days: int = 7,
    max_documents: int = 20,
    max_events: int = 20,
    max_refs: int = 50,
    hard_limit: int = 1000,
) -> ValidatedEventLogRequest:
    """Полная чистая валидация параметров. Кидает EventLogValidationError."""
    # 1. Даты
    start = parse_local_datetime(start_datetime, "start_datetime")
    end = parse_local_datetime(end_datetime, "end_datetime")

    # 2. Порядок периода
    if end <= start:
        raise EventLogValidationError("INVALID_PERIOD", "Конец периода должен быть позже начала")

    # 3. Limit (нельзя обойти строкой/булевым/отрицательным)
    if isinstance(limit, bool):
        raise EventLogValidationError("INVALID_LIMIT", "Параметр limit должен быть целым числом")
    try:
        requested_limit = int(limit)
    except (TypeError, ValueError):
        raise EventLogValidationError("INVALID_LIMIT", "Параметр limit должен быть целым числом")
    if requested_limit < 1:
        raise EventLogValidationError("INVALID_LIMIT", "Параметр limit должен быть не меньше 1")
    limit = min(requested_limit, hard_limit)

    # 4. Размеры массивов + форматы
    if isinstance(events, (str, bytes)):
        raise EventLogValidationError("INVALID_EVENT", "events должен быть списком")
    evts = tuple(events or ())
    if len(evts) > max_events:
        raise EventLogValidationError(
            "TOO_MANY_EVENTS", f"Слишком много событий ({len(evts)} > {max_events})"
        )
    for ev in evts:
        validate_event_code(ev)

    if isinstance(data_refs, (str, bytes, dict)):
        raise EventLogValidationError("INVALID_DATA_REFS", "data_refs должен быть списком объектов")
    refs = tuple(data_refs or ())
    if len(refs) > max_refs:
        raise EventLogValidationError(
            "TOO_MANY_REFS", f"Слишком много ссылок ({len(refs)} > {max_refs})"
        )
    parsed_refs = tuple(validate_data_ref(r, metadata_path) for r in refs)

    if isinstance(document_numbers, (str, bytes)):
        raise EventLogValidationError(
            "INVALID_DOCUMENT_NUMBER", "document_numbers должен быть списком"
        )
    numbers = tuple(document_numbers or ())
    if len(numbers) > max_documents:
        raise EventLogValidationError(
            "TOO_MANY_DOCUMENTS",
            f"Слишком много номеров документов ({len(numbers)} > {max_documents})",
        )
    seen: set[str] = set()
    clean_numbers: list[str] = []
    for num in numbers:
        validate_document_number(num)
        key = num.strip()
        if key in seen:
            raise EventLogValidationError(
                "DUPLICATE_DOCUMENT_NUMBER", f"Дубликат номера: {key}"
            )
        seen.add(key)
        clean_numbers.append(key)

    # 5. metadata_path
    validate_metadata_path(metadata_path)

    # 6. document_numbers требуют metadata_path вида Документ.* и диапазон даты
    if clean_numbers:
        if not metadata_path or not metadata_path.startswith("Документ."):
            raise EventLogValidationError(
                "INVALID_METADATA_PATH",
                "document_numbers разрешены только при metadata_path вида Документ.*",
            )
        if not document_date_from or not document_date_to:
            raise EventLogValidationError(
                "DOCUMENT_DATE_RANGE_REQUIRED",
                "document_numbers требуют document_date_from и document_date_to",
            )
    elif document_date_from is not None or document_date_to is not None:
        raise EventLogValidationError(
            "DOCUMENT_DATE_RANGE_REQUIRED",
            "Диапазон даты документа задаётся только вместе с document_numbers",
        )

    doc_from = None
    doc_to = None
    if clean_numbers:
        doc_from = parse_local_datetime(document_date_from, "document_date_from")
        doc_to = parse_local_datetime(document_date_to, "document_date_to")
        if doc_to <= doc_from:
            raise EventLogValidationError(
                "INVALID_DOCUMENT_DATE_RANGE",
                "document_date_to должен быть позже document_date_from",
            )

    # 7. order
    if order not in ("asc", "desc"):
        raise EventLogValidationError("INVALID_ORDER", "order должен быть 'asc' или 'desc'")

    # 8. Обязательная селективность
    has_object = bool(parsed_refs)
    has_numbers = bool(clean_numbers)
    has_events = bool(evts) and bool(metadata_path)
    if not (has_object or has_numbers or has_events):
        raise EventLogValidationError(
            "UNSAFE_UNFILTERED_REQUEST",
            "Запрос только по периоду запрещён. Укажите data_refs, document_numbers "
            "или пару metadata_path+events.",
        )

    # 9. Ограничение периода
    range_days = (end - start).total_seconds() / 86400.0
    allowed = max_range_days if (has_object or has_numbers) else broad_range_days
    if range_days > allowed:
        raise EventLogValidationError(
            "DATE_RANGE_TOO_LARGE",
            f"Период слишком велик (максимум {allowed} дней для этого фильтра)",
        )

    return ValidatedEventLogRequest(
        start_datetime=start,
        end_datetime=end,
        events=evts,
        metadata_path=metadata_path,
        data_refs=parsed_refs,
        document_numbers=tuple(clean_numbers),
        document_date_from=doc_from,
        document_date_to=doc_to,
        successful_only=bool(successful_only),
        order=order,
        limit=limit,
    )


# ---------------------------------------------------------------------------
# Нормализация строк таблицы журнала в JSON
# ---------------------------------------------------------------------------

def transaction_status_code(status_name: str) -> str:
    """Переводит русское представление статуса транзакции в машинный код."""
    return _TRANSACTION_STATUS_CODES.get(status_name, "unknown")


def is_successful_status(status_name: str) -> bool:
    return status_name in SUCCESSFUL_TRANSACTION_STATUSES


_MISSING = object()


def normalize_row(row: dict) -> dict:
    """Приводит сырые значения строки журнала к стабильным JSON-ключам.

    Отсутствующие поля не выдумываются: функция возвращает только те ключи,
    значения которых присутствуют в `row` (по ключу или по `null`).
    """
    out: dict = {}

    def put(key: str, value: Any) -> None:
        if value is _MISSING:
            return
        out[key] = value

    ts = row.get("timestamp", _MISSING)
    put("timestamp", _iso(ts))

    for key in (
        "user_id",
        "user_name",
        "computer",
        "application",
        "application_presentation",
        "connection_number",
        "session",
        "event",
        "event_presentation",
        "metadata",
        "metadata_presentation",
        "data",
        "data_presentation",
        "comment",
        "transaction_id",
        "server_name",
        "port",
        "sync_port",
    ):
        put(key, row.get(key, _MISSING))

    status = row.get("transaction_status", _MISSING)
    if status is not _MISSING:
        status_name = "" if status is None else str(status)
        code = transaction_status_code(status_name)
        out["transaction_status"] = status_name
        out["transaction_status_code"] = code
        out["successful"] = is_successful_status(status_name)

    return out


def _iso(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return value


# ---------------------------------------------------------------------------
# Формирование отбора 1С (duck typing, без импорта COM)
# ---------------------------------------------------------------------------
#
# Эти функции не импортируют pythoncom/pywintypes — все взаимодействия идут
# через переданный объект соединения. Это позволяет гонять их на fake-объектах
# в tests/test_event_log.py на любой ОС.


def onec_array(conn: Any, values: list[Any]) -> Any:
    """Создаёт COM-массив 1С. Нельзя передавать Python-список напрямую."""
    result = conn.NewObject("Массив")
    for value in values:
        result.Добавить(value)
    return result


def onec_transaction_statuses(conn: Any, successful_only: bool) -> Any | None:
    """Массив статусов транзакции для фильтра, либо None (без фильтра)."""
    if not successful_only:
        return None
    statuses = conn.СтатусТранзакцииЗаписиЖурналаРегистрации
    return onec_array(conn, [statuses.Зафиксирована, statuses.НетТранзакции])


def make_common_filter(conn: Any, request: Any) -> Any:
    """Общая структура условий (без ссылки Данные)."""
    item = conn.NewObject("Структура")
    item.Вставить("ДатаНачала", request.start_datetime)
    item.Вставить("ДатаОкончания", request.end_datetime)
    if request.events:
        item.Вставить("Событие", onec_array(conn, list(request.events)))
    if request.metadata_path:
        item.Вставить("Метаданные", onec_array(conn, [request.metadata_path]))
    statuses = onec_transaction_statuses(conn, request.successful_only)
    if statuses is not None:
        item.Вставить("СтатусТранзакции", statuses)
    return item


def build_event_log_filter(conn: Any, request: Any, resolved_refs: list[Any]) -> Any:
    """Строит отбор 1С: общие условия AND по одной точной ссылке на структуру.

    При нескольких ссылках создаётся по одной структуре на ссылку (верхний
    уровень — Массив структур, соединённых по OR). Нельзя помещать список
    ссылок в поле `Данные` одной структуры.
    """
    if resolved_refs:
        filter_1c = conn.NewObject("Массив")
        for ref in resolved_refs:
            item = make_common_filter(conn, request)
            item.Вставить("Данные", ref)
            filter_1c.Добавить(item)
        return filter_1c
    return make_common_filter(conn, request)
