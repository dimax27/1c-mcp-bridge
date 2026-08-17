"""
1C MCP Bridge — MCP-сервер для подключения AI-ассистентов к 1С:Предприятию через COM.

v0.3.0: поддержка нескольких MCP-клиентов (Claude Desktop, Qwen Desktop, Kimi Desktop).
Список баз хранится в databases.json (рядом с этим скриптом или в пути, заданном
переменной ONEC_DATABASES_FILE). Каждый из инструментов принимает опциональный
параметр `database` — ключ из списка. При отсутствии параметра используется
default_database из конфига.

Формат databases.json:
{
    "version": 1,
    "default_database": "ut",
    "databases": {
        "ut": {
            "description": "Управление торговлей 10.3",
            "progid": "V83.COMConnector",
            "connection_string": "Srvr=\"127.0.0.1\";Ref=\"ut10\""
        },
        ...
    }
}
"""

from __future__ import annotations

import datetime
import io
import json
import logging
import os
import secrets
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client
from mcp.server import MCPServer
from mcp_types import ToolAnnotations

from version import VERSION

# Все инструменты моста только читают данные из 1С — клиент может
# использовать подсказку для классификации действия и решений о подтверждениях.
READ_ONLY_1C_TOOL = ToolAnnotations(read_only_hint=True, open_world_hint=False)

# На Windows stdout/stderr по умолчанию в cp1251 — MCP-клиенты пишут лог в
# UTF-8. Принудительно перекодируем.
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", line_buffering=True
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", line_buffering=True
        )
    except (AttributeError, ValueError):
        pass

logging.basicConfig(
    level=os.environ.get("ONEC_LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("mcp-1c")

# ---------------------------------------------------------------------------
# Загрузка списка баз
# ---------------------------------------------------------------------------

from config import (
    DEFAULT_LIMIT,
    EVENT_LOG_DEFAULT_LIMIT,
    EVENT_LOG_HARD_LIMIT,
    EVENT_LOG_MAX_BROAD_RANGE_DAYS,
    EVENT_LOG_MAX_DOCUMENTS,
    EVENT_LOG_MAX_EVENTS,
    EVENT_LOG_MAX_RANGE_DAYS,
    EVENT_LOG_MAX_REFS,
    EVENT_LOG_MAX_RESPONSE_BYTES,
    HARD_LIMIT,
    MAX_COLUMNS,
    MAX_PARAMETERS,
    MAX_QUERY_LENGTH,
)
from credentials import build_conn_str, migrate_to_encrypted
from event_log import (
    EVENT_LOG_JSON_FIELDS,
    EventLogAccessError,
    EventLogValidationError,
    build_event_log_filter,
    classify_event_log_com_error,
    filter_fingerprint,
    onec_array,
    safe_event_log_error,
    transaction_status_code,
    validate_event_log_request,
)
from query_timeout import check_slow


def find_databases_file() -> Path:
    """Ищем databases.json в порядке приоритета:
    ONEC_DATABASES_FILE → ProgramData → Program Files → рядом со скриптом.
    """
    env_path = os.environ.get("ONEC_DATABASES_FILE", "").strip()
    if env_path:
        return Path(env_path)

    program_data = os.environ.get("PROGRAMDATA", "C:/ProgramData")
    standard = Path(program_data) / "1cMcpBridge" / "databases.json"
    if standard.exists():
        return standard

    here = Path(__file__).resolve().parent
    for legacy in [Path("C:/Program Files/1cMcpBridge/databases.json"),
                   here / "databases.json"]:
        if legacy.exists():
            return legacy

    return standard


def load_databases() -> dict:
    """Читает и валидирует databases.json. Возвращает словарь с ключами:
    'default_database': str
    'databases': {key: {description, progid, connection_string, dll_path?}}
    """
    path = find_databases_file()
    if not path.exists():
        raise FileNotFoundError(
            f"Не найден файл со списком баз: {path}\n"
            f"Создайте его или укажите путь в переменной ONEC_DATABASES_FILE."
        )

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Ошибка JSON в {path}: {e}")

    if not isinstance(data, dict):
        raise RuntimeError(f"{path}: ожидался объект на верхнем уровне")

    databases = data.get("databases")
    if not isinstance(databases, dict) or not databases:
        raise RuntimeError(f"{path}: пустой или отсутствует ключ 'databases'")

    # Валидация каждой базы
    migrated = False
    for key, cfg in databases.items():
        if not isinstance(cfg, dict):
            raise RuntimeError(f"databases.{key}: должна быть объектом")
        for required in ("progid", "connection_string"):
            if not cfg.get(required):
                raise RuntimeError(
                    f"databases.{key}: отсутствует обязательное поле '{required}'"
                )
        # Описание короткое, notes — длинное (что в этой базе можно найти)
        cfg.setdefault("description", key)
        cfg.setdefault("notes", "")
        cfg.setdefault("enabled", True)

        # Auto-migrate plaintext passwords to DPAPI (non-destructive)
        if migrate_to_encrypted(cfg):
            log.info("Мигрирован пароль для '%s' в DPAPI", key)
            migrated = True

    # Фильтруем отключённые — AI-клиенты их вообще не должны видеть
    enabled_databases = {k: v for k, v in databases.items() if v.get("enabled", True)}
    if not enabled_databases:
        raise RuntimeError(
            f"{path}: все базы отключены (enabled=false). "
            f"Включи хотя бы одну в 1C Bridge Manager."
        )

    default_db = data.get("default_database") or next(iter(enabled_databases))
    if default_db not in enabled_databases:
        default_db = next(iter(enabled_databases))

    # Save back if passwords were migrated to DPAPI
    if migrated:
        try:
            data["databases"] = databases
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            log.info("databases.json обновлён с DPAPI-миграцией")
        except Exception as exc:
            log.warning("Не удалось сохранить миграцию: %s", exc)

    return {
        "default_database": default_db,
        "databases": enabled_databases,
    }


# Загружаем единожды при старте; перечитывать на лету не будем — переустановка
# всё равно требует перезапуска MCP-клиента.
try:
    DB_CONFIG = load_databases()
except Exception as e:
    log.error("Не удалось загрузить databases.json: %s", e)
    log.error("Сервер запускается, но все вызовы инструментов будут падать.")
    DB_CONFIG = {"default_database": "", "databases": {}}


def list_database_keys() -> list[str]:
    return list(DB_CONFIG["databases"].keys())


def get_db_descriptions() -> str:
    """Полное описание баз для подсказки AI в tools/list.
    Включает короткое description и длинные notes (если есть) — AI-клиент
    увидит заметки до того как пользователь задаст вопрос, и сможет
    осознанно выбирать в какую базу обращаться.
    """
    if not DB_CONFIG["databases"]:
        return "(нет настроенных баз)"
    lines = ["Настроенные базы данных:"]
    default = DB_CONFIG["default_database"]
    for key, cfg in DB_CONFIG["databases"].items():
        marker = " [по умолчанию]" if key == default else ""
        lines.append(f"\n  • '{key}'{marker}: {cfg['description']}")
        notes = cfg.get("notes", "").strip()
        if notes:
            # Каждая строка заметок с отступом
            for note_line in notes.splitlines():
                lines.append(f"      {note_line}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Управление COM-соединениями
# ---------------------------------------------------------------------------
#
# Для каждой базы держим живое соединение в thread-local словаре. COM-объекты
# привязаны к потоку (STA), поэтому индексируем по (thread_id, db_key).

_tls = threading.local()


def _ensure_com() -> None:
    if not getattr(_tls, "com_init", False):
        pythoncom.CoInitialize()
        _tls.com_init = True


def get_connection(db_key: str) -> Any:
    """Возвращает живое COM-соединение к базе db_key. Лениво создаёт."""
    _ensure_com()

    if not hasattr(_tls, "connections"):
        _tls.connections = {}

    conn = _tls.connections.get(db_key)
    if conn is not None:
        # Проверка живости
        try:
            _ = conn.Метаданные.Имя
            return conn
        except pywintypes.com_error:
            log.warning("Соединение к %s умерло, переподключаюсь", db_key)
            _tls.connections.pop(db_key, None)

    if db_key not in DB_CONFIG["databases"]:
        raise ValueError(
            f"База '{db_key}' не настроена. Доступные: {list_database_keys()}"
        )

    cfg = DB_CONFIG["databases"][db_key]
    progid = cfg["progid"]
    conn_str = build_conn_str(cfg)

    log.info("Подключаюсь к '%s' через %s", db_key, progid)
    # EnsureDispatch кеширует type library (быстрее), но может не работать
    # у обычного пользователя (PermissionError при записи gen_py в venv).
    try:
        connector = win32com.client.gencache.EnsureDispatch(progid)
    except Exception as exc:
        log.warning("EnsureDispatch(%s) недоступен, использую Dispatch: %s",
                    progid, exc)
        connector = win32com.client.Dispatch(progid)
    conn = connector.Connect(conn_str)
    _tls.connections[db_key] = conn

    try:
        log.info("Подключение к '%s' установлено: ИБ %s",
                 db_key, conn.Метаданные.Имя)
    except pywintypes.com_error:
        log.info("Подключение к '%s' установлено", db_key)

    return conn


def resolve_database(db_param: str | None) -> str:
    """Возвращает ключ базы. None → default_database. Невалидный → ошибка."""
    if not db_param:
        if not DB_CONFIG["default_database"]:
            raise ValueError("Нет настроенных баз — заполните databases.json")
        return DB_CONFIG["default_database"]
    if db_param not in DB_CONFIG["databases"]:
        raise ValueError(
            f"База '{db_param}' не настроена. Доступные: {list_database_keys()}"
        )
    return db_param


# ---------------------------------------------------------------------------
# Сериализация значений (без изменений из v0.1.0)
# ---------------------------------------------------------------------------

EMPTY_DATE_YEAR = 1900


def serialize_value(v: Any, depth: int = 0) -> Any:
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float, str)):
        return v
    if isinstance(v, datetime.datetime):
        if v.year < EMPTY_DATE_YEAR:
            return None
        return v.isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, pywintypes.TimeType):
        try:
            d = datetime.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
            return None if d.year < EMPTY_DATE_YEAR else d.isoformat()
        except Exception:
            return str(v)
    if depth > 1:
        try:
            return str(v)
        except Exception:
            return f"<COM:{type(v).__name__}>"
    try:
        uuid = str(v.УникальныйИдентификатор())
        try:
            type_name = str(v.Метаданные().ПолноеИмя())
        except Exception:
            type_name = None
        try:
            presentation = str(v)
        except Exception:
            presentation = None
        return {
            "_ref": uuid,
            "_type": type_name,
            "_presentation": presentation,
        }
    except (AttributeError, pywintypes.com_error):
        pass
    try:
        type_name = str(v.Метаданные().ПолноеИмя())
        return {"_enum": type_name, "_value": str(v)}
    except (AttributeError, pywintypes.com_error):
        pass
    try:
        return str(v)
    except Exception:
        return f"<нерасшифровано: {type(v).__name__}>"


def parse_parameter(value: Any, conn: Any) -> Any:
    if isinstance(value, dict) and "_ref" in value and "_type" in value:
        return get_ref_by_uuid(conn, value["_type"], value["_ref"])
    if isinstance(value, str):
        v = value.strip()
        try:
            if len(v) == 10 and v[4] == "-" and v[7] == "-":
                return datetime.datetime.fromisoformat(v)
            if "T" in v:
                return datetime.datetime.fromisoformat(v.replace("Z", ""))
        except ValueError:
            pass
        return value
    if isinstance(value, list):
        arr = conn.NewObject("Массив")
        for item in value:
            arr.Добавить(parse_parameter(item, conn))
        return arr
    return value


def get_ref_by_uuid(conn: Any, type_path: str, uuid_str: str) -> Any:
    parts = type_path.split(".")
    if len(parts) != 2:
        raise ValueError(f"Неверный формат типа: {type_path!r}")
    kind, name = parts
    coll_map = {
        "Справочник": "Справочники",
        "Документ": "Документы",
        "Перечисление": "Перечисления",
        "ПланВидовХарактеристик": "ПланыВидовХарактеристик",
        "ПланСчетов": "ПланыСчетов",
        "ПланВидовРасчета": "ПланыВидовРасчета",
    }
    if kind not in coll_map:
        raise ValueError(f"Тип {kind!r} не поддерживает разрешение по UUID")
    manager = getattr(getattr(conn, coll_map[kind]), name)
    uuid_obj = conn.NewObject("УникальныйИдентификатор", uuid_str)
    return manager.ПолучитьСсылку(uuid_obj)


# ---------------------------------------------------------------------------
# Метаданные (без изменений из v0.1.0)
# ---------------------------------------------------------------------------

def hasattr_safe(obj: Any, name: str) -> bool:
    try:
        getattr(obj, name)
        return True
    except (AttributeError, pywintypes.com_error):
        return False


def list_collection(coll: Any) -> list[dict]:
    out: list[dict] = []
    try:
        for item in coll:
            entry: dict = {"name": str(item.Имя)}
            for prop, key in (("Тип", "type"), ("Синоним", "synonym")):
                try:
                    val = getattr(item, prop)
                    s = str(val) if val is not None else ""
                    if s:
                        entry[key] = s
                except (AttributeError, pywintypes.com_error):
                    pass
            out.append(entry)
    except pywintypes.com_error:
        pass
    return out


METADATA_COLLECTION_MAP = {
    "Справочник": "Справочники",
    "Документ": "Документы",
    "РегистрНакопления": "РегистрыНакопления",
    "РегистрСведений": "РегистрыСведений",
    "РегистрБухгалтерии": "РегистрыБухгалтерии",
    "Перечисление": "Перечисления",
    "ПланВидовХарактеристик": "ПланыВидовХарактеристик",
    "ПланСчетов": "ПланыСчетов",
    "Константа": "Константы",
    "Отчет": "Отчеты",
    "Обработка": "Обработки",
}


def resolve_metadata(conn: Any, path: str) -> Any:
    parts = path.split(".")
    if len(parts) != 2:
        return None
    kind, name = parts
    coll_name = METADATA_COLLECTION_MAP.get(kind)
    if not coll_name:
        return None
    coll = getattr(conn.Метаданные, coll_name)
    try:
        return coll.Найти(name)
    except (AttributeError, pywintypes.com_error):
        for obj in coll:
            if str(obj.Имя) == name:
                return obj
        return None


def virtual_tables_for(kind: str, obj: Any) -> list[str]:
    if kind == "РегистрНакопления":
        try:
            view = str(obj.ВидРегистра)
        except Exception:
            view = ""
        if view in ("Остатки", "ОстаткиИОбороты") or "статк" in view:
            return ["Остатки", "Обороты", "ОстаткиИОбороты"]
        return ["Обороты"]
    if kind == "РегистрСведений":
        return ["СрезПервых", "СрезПоследних"]
    if kind == "РегистрБухгалтерии":
        return ["Остатки", "Обороты", "ОстаткиИОбороты", "ОборотыДтКт",
                "Движения", "ДвиженияССубконто"]
    return []


def parse_com_error(e: pywintypes.com_error) -> str:
    try:
        if len(e.args) >= 3 and e.args[2]:
            exc_info = e.args[2]
            if len(exc_info) >= 3 and exc_info[2]:
                return str(exc_info[2])
        return str(e)
    except Exception:
        return repr(e)


# ---------------------------------------------------------------------------
# Журнал регистрации (get_registration_log)
# ---------------------------------------------------------------------------
#
# COM-логику держим через duck typing: функции принимают объект соединения,
# у которого есть NewObject(...), ВыгрузитьЖурналРегистрации(...), статусы
# транзакции и т.п. Это позволяет тестировать на fake-объектах без 1С.

# Исходные колонки для современной 8.3. На старых платформах часть полей может
# отсутствовать — spike фиксирует фактический набор; обязательное ядро:
# Дата, Пользователь, ИмяПользователя, Событие, Метаданные, Данные,
# ПредставлениеДанных, СтатусТранзакции.
EVENT_LOG_COLUMNS = (
    "Дата,Пользователь,ИмяПользователя,Компьютер,"
    "ИмяПриложения,ПредставлениеПриложения,"
    "Событие,ПредставлениеСобытия,Комментарий,"
    "Метаданные,ПредставлениеМетаданных,Данные,"
    "ПредставлениеДанных,СтатусТранзакции,Транзакция,"
    "Сеанс,Соединение,РабочийСервер,"
    "ОсновнойIPПорт,ВспомогательныйIPПорт"
)

# Маппинг колонок 1С -> ключи нормализованной строки. "Соединение" и
# "НомерСоединения" оба могут встречаться на разных платформах.
_EVENT_LOG_COLUMN_MAP = {
    "Дата": "timestamp",
    "Пользователь": "user_id",
    "ИмяПользователя": "user_name",
    "Компьютер": "computer",
    "ИмяПриложения": "application",
    "ПредставлениеПриложения": "application_presentation",
    "Соединение": "connection_number",
    "НомерСоединения": "connection_number",
    "Событие": "event",
    "ПредставлениеСобытия": "event_presentation",
    "Метаданные": "metadata",
    "ПредставлениеМетаданных": "metadata_presentation",
    "Данные": "data",
    "ПредставлениеДанных": "data_presentation",
    "Комментарий": "comment",
    "СтатусТранзакции": "transaction_status",
    "Транзакция": "transaction_id",
    "Сеанс": "session",
    "РабочийСервер": "server_name",
    "ОсновнойIPПорт": "port",
    "ВспомогательныйIPПорт": "sync_port",
}

# Однопоточный замок: не более одного чтения журнала на одну базу в рамках
# процесса. Разные базы читаются независимо.
_event_log_locks: dict[str, threading.Lock] = {}
_event_log_locks_guard = threading.Lock()


def _event_log_lock(db_key: str) -> threading.Lock:
    with _event_log_locks_guard:
        lock = _event_log_locks.get(db_key)
        if lock is None:
            lock = threading.Lock()
            _event_log_locks[db_key] = lock
        return lock


def _resolve_document_numbers(conn: Any, request: Any) -> tuple[list[Any], list[dict]]:
    """Разрешает номера документов в однозначные COM-ссылки.

    Возвращает (ссылки, публичное представление разрешённых документов).
    Кидает EventLogValidationError при 0 или >1 документе на номер.
    """
    table_name = request.metadata_path
    # Безопасность: имя таблицы уже проверено regex'ом Документ.<Имя> и
    # resolve_metadata в вызывающем коде; здесь подставляем только валидное имя.
    query = conn.NewObject("Запрос")
    query.Text = (
        "ВЫБРАТЬ Док.Ссылка КАК Ссылка, Док.Номер КАК Номер, "
        "Док.Дата КАК Дата, ПРЕДСТАВЛЕНИЕ(Док.Ссылка) КАК Представление "
        f"ИЗ {table_name} КАК Док "
        "ГДЕ Док.Номер В (&Номера) И Док.Дата >= &ДатаНачалаДокумента "
        "И Док.Дата < &ДатаОкончанияДокумента"
    )
    numbers_arr = onec_array(conn, list(request.document_numbers))
    query.УстановитьПараметр("Номера", numbers_arr)
    query.УстановитьПараметр("ДатаНачалаДокумента", request.document_date_from)
    query.УстановитьПараметр("ДатаОкончанияДокумента", request.document_date_to)

    try:
        result = query.Выполнить()
    except pywintypes.com_error as exc:
        raise EventLogValidationError(
            "ONEC_ERROR", f"Не удалось разрешить номера документов: {parse_com_error(exc)}"
        ) from exc

    by_number: dict[str, list[dict]] = {}
    selection = result.Выбрать()
    while selection.Следующий():
        num = str(selection.Номер)
        entry = {
            "number": num,
            "date": _serialize_event_log_datetime(selection.Дата),
            "presentation": str(selection.Представление),
            "ref": selection.Ссылка,
        }
        by_number.setdefault(num, []).append(entry)

    resolved_refs: list[Any] = []
    public_payload: list[dict] = []
    for num in request.document_numbers:
        candidates = by_number.get(num, [])
        if len(candidates) == 0:
            raise EventLogValidationError(
                "DOCUMENT_NOT_FOUND",
                f"Документ с номером {num!r} не найден в указанном диапазоне дат",
            )
        if len(candidates) > 1:
            dates = ", ".join(str(c["date"]) for c in candidates)
            raise EventLogValidationError(
                "DOCUMENT_AMBIGUOUS",
                f"Номер {num!r} неоднозначен: найдено {len(candidates)} документов "
                f"(даты: {dates}). Сузьте диапазон даты документа или передайте _ref.",
            )
        c = candidates[0]
        resolved_refs.append(c["ref"])
        public_payload.append(
            {
                "type_path": request.metadata_path,
                "number": c["number"],
                "date": c["date"],
                "presentation": c["presentation"],
            }
        )
    return resolved_refs, public_payload


def _resolve_data_refs(conn: Any, request: Any) -> list[Any]:
    """Разрешает data_refs (uuid) в COM-ссылки через существующий get_ref_by_uuid."""
    refs = []
    for rr in request.data_refs:
        refs.append(get_ref_by_uuid(conn, rr["_type"], rr["_ref"]))
    return refs


def serialize_event_log_table(
    table: Any, limit: int, conn: Any
) -> tuple[list[dict], bool]:
    """Преобразует ТаблицуЗначений журнала в список нормализованных строк.

    Возвращает (rows, truncated). Строки читаются по индексу колонки (row.Get(i)),
    как того требует ТаблицаЗначений 1С. Неизвестные COM-типы не роняют вызов.

    Ссылки `Данные` сериализуются в {"_ref","_type"} через conn.XMLСтрока() и
    Метаданные().ПолноеИмя() (str(ref) в pywin32 312 не даёт UUID).
    Статус транзакции сравнивается с перечислением СтатусТранзакцииЗаписи...
    """
    column_names = list(table.Колонки)
    col_keys = [_EVENT_LOG_COLUMN_MAP.get(str(c.Имя)) for c in column_names]
    statuses = conn.СтатусТранзакцииЗаписиЖурналаРегистрации

    rows: list[dict] = []
    truncated = False
    for row in table:
        # читаем значения по индексу колонки
        values: list[Any] = []
        for ci in range(len(column_names)):
            try:
                values.append(row.Get(ci))
            except (AttributeError, pywintypes.com_error, KeyError):
                values.append(None)

        ser: dict = {}
        for ci, key in enumerate(col_keys):
            if key is None:
                continue
            v = values[ci]
            if key == "data":
                ser[key] = _serialize_event_log_ref(v, conn)
            elif key == "user_id":
                # Пользователь — не ссылка справочника; оставляем как есть,
                # а имя пользователя уже в user_name.
                ser[key] = _serialize_event_log_user(v, conn)
            elif key == "transaction_status":
                ser["transaction_status"] = _transaction_status_name(v, statuses)
                code = transaction_status_code(_transaction_status_name(v, statuses))
                ser["transaction_status_code"] = code
                ser["successful"] = code in ("committed", "not_applicable")
            elif key == "timestamp":
                ser[key] = _serialize_event_log_datetime(v)
            elif isinstance(v, (str, int, float, bool)) or v is None:
                ser[key] = v
            else:
                try:
                    ser[key] = str(v)
                except Exception:
                    ser[key] = serialize_value(v)

        rows.append(ser)

    # Платформа возвращает последние limit+1 событий; при переполнении нужно
    # сохранить НАИБОЛЕЕ НОВЫЕ limit строк, а не первые. Сортируем стабильно по
    # timestamp asc и забираем последние limit (самая новая запись не теряется).
    if len(rows) > limit:
        truncated = True
        rows.sort(key=lambda r: (r.get("timestamp") or "", ), reverse=False)
        rows = rows[-limit:]

    return rows, truncated


def _serialize_event_log_ref(v: Any, conn: Any) -> Any:
    """Поле Данные -> {"_ref","_type"} для ссылки, либо обычное значение.

    Для большинства событий `Данные` — ссылка на объект 1С. Но для некоторых
    событий это строка, число, структура или таблица значений. Определяем по
    наличию `Метаданные()`: если метод есть и возвращает ПолноеИмя — это ссылка
    (сериализуем через XMLСтрока + Метаданные), иначе возвращаем значение как
    есть (строка) или через serialize_value.
    """
    if v is None:
        return None
    type_name = None
    try:
        type_name = str(v.Метаданные().ПолноеИмя())
    except (AttributeError, pywintypes.com_error):
        type_name = None

    if type_name is None:
        # Не ссылка — возвращаем обычное значение (строка/число/структура/таблица)
        if isinstance(v, (str, int, float, bool)):
            return v
        try:
            return serialize_value(v)
        except Exception:
            return str(v)

    ref = None
    try:
        ref = conn.XMLСтрока(v)
    except Exception:
        ref = None
    return {"_ref": ref, "_type": type_name}


def _serialize_event_log_user(v: Any, conn: Any) -> Any:
    """Пользователь (поле Пользователь) — UUID через XMLСтрока, иначе None."""
    if v is None:
        return None
    try:
        return conn.XMLСтрока(v)
    except Exception:
        return None


def _transaction_status_name(v: Any, statuses: Any) -> str:
    """Сравнивает COM-значение статуса с перечислением и возвращает имя."""
    for name in ("Зафиксирована", "НетТранзакции", "Отменена", "НеЗавершена"):
        try:
            if v == getattr(statuses, name):
                return name
        except Exception:  # noqa: S112 — сравнение с COM-перечислением, best-effort
            continue
    return "" if v is None else str(v)


def _serialize_event_log_datetime(v: Any) -> Any:
    """Дата журнала — ISO без timezone-offset (локальное время ИБ)."""
    if v is None:
        return None
    if isinstance(v, pywintypes.TimeType):
        try:
            d = datetime.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
            return "" if d.year < EMPTY_DATE_YEAR else d.isoformat()
        except Exception:
            return str(v)
    if isinstance(v, datetime.datetime):
        if v.tzinfo is not None:
            try:
                v = v.replace(tzinfo=None)
            except Exception:  # noqa: S110 — timezone-aware datetime может не иметь replace
                pass
        return v.isoformat()
    return str(v)


def log_event_log_read(
    *,
    request_id: str,
    db_key: str,
    elapsed_ms: float,
    row_count: int,
    truncated: bool,
    fingerprint: str,
) -> None:
    """Логирует только безопасные признаки чтения журнала (без номеров/UUID)."""
    log.info(
        "event_log read req=%s db=%s rows=%d truncated=%s %.1fms filter=%s",
        request_id,
        db_key,
        row_count,
        truncated,
        elapsed_ms,
        fingerprint,
    )


# ---------------------------------------------------------------------------
# MCP server + tools
# ---------------------------------------------------------------------------

mcp = MCPServer("1c-bridge")

# Описание баз — попадает в docstring каждого инструмента и AI-клиент видит его
# в tools/list ДО первого вопроса от пользователя.
_DB_BLOCK = "\n\n" + get_db_descriptions() + "\n"


def _with_db_info(doc: str) -> str:
    """Дописывает информацию о базах в конец docstring."""
    return doc.rstrip() + _DB_BLOCK


@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def execute_query(
    text: str,
    parameters: dict | None = None,
    limit: int = DEFAULT_LIMIT,
    database: str | None = None,
) -> dict:
    """Выполняет запрос на языке запросов 1С 8.2 и возвращает табличный результат.

    Используй язык запросов 1С (русские ключевые слова: ВЫБРАТЬ, ИЗ, ГДЕ,
    СГРУППИРОВАТЬ ПО, УПОРЯДОЧИТЬ ПО, ИМЕЮЩИЕ, ОБЪЕДИНИТЬ ВСЕ, ЛЕВОЕ СОЕДИНЕНИЕ).
    Виртуальные таблицы регистров: РегистрНакопления.Имя.Обороты(&НачДата, &КонДата).
    Параметры в тексте — через &ИмяПараметра, передавай их в `parameters`.

    Параметры могут быть:
        - строки/числа/булевы — передаются как есть;
        - даты — ISO-строкой "2026-04-01" или "2026-04-01T00:00:00";
        - ссылки — объектом {"_ref": "uuid", "_type": "Справочник.Контрагенты"};
        - массивы — обычным списком JSON.

    Args:
        text: Текст запроса 1С.
        parameters: Словарь параметров (опционально).
        limit: Максимум строк в ответе (по умолчанию 1000, hard cap 10000).
        database: Ключ базы данных (см. описание в начале списка инструментов).

    Returns:
        columns: схема результата
        rows: список строк
        row_count, truncated, execution_time_ms, database
    """
    if not text or not text.strip():
        return {"error": "Пустой текст запроса"}
    if len(text) > MAX_QUERY_LENGTH:
        return {"error": f"Слишком длинный запрос ({len(text)} > {MAX_QUERY_LENGTH})"}
    if parameters and len(parameters) > MAX_PARAMETERS:
        return {"error": f"Слишком много параметров ({len(parameters)} > {MAX_PARAMETERS})"}
    try:
        requested_limit = int(limit)
    except (TypeError, ValueError):
        return {"error": "Параметр limit должен быть целым числом"}
    limit = min(max(1, requested_limit), HARD_LIMIT)

    try:
        db_key = resolve_database(database)
        conn = get_connection(db_key)
        query = conn.NewObject("Запрос")
        query.Текст = text

        if parameters:
            for name, raw in parameters.items():
                try:
                    parsed = parse_parameter(raw, conn)
                    query.УстановитьПараметр(name, parsed)
                except Exception as e:
                    return {"error": f"Параметр '{name}': {e}", "database": db_key}

        t0 = time.perf_counter()
        result = query.Выполнить()
        elapsed_ms = lambda: round((time.perf_counter() - t0) * 1000, 1)
        check_slow(t0, db_key, text)

        try:
            empty = bool(result.Пустой())
        except (AttributeError, pywintypes.com_error):
            empty = False

        if empty:
            return {
                "database": db_key,
                "columns": [],
                "rows": [],
                "row_count": 0,
                "truncated": False,
                "columns_truncated": False,
                "total_column_count": 0,
                "execution_time_ms": elapsed_ms(),
            }

        # Collect columns (may be truncated)
        all_cols = list(result.Колонки)
        total_column_count = len(all_cols)
        columns_truncated = total_column_count > MAX_COLUMNS
        cols_to_process = all_cols[:MAX_COLUMNS]
        columns_meta = []
        col_names = []
        for col in cols_to_process:
            n = str(col.Имя)
            col_names.append(n)
            try:
                t_str = str(col.ТипЗначения)
            except Exception:
                t_str = ""
            columns_meta.append({"name": n, "type": t_str})

        selection = result.Выбрать()
        rows = []
        truncated = False
        while selection.Следующий():
            if len(rows) >= limit:
                truncated = True
                break
            row = {}
            for cn in col_names:
                try:
                    val = getattr(selection, cn)
                except (AttributeError, pywintypes.com_error):
                    val = None
                row[cn] = serialize_value(val)
            rows.append(row)

        return {
            "database": db_key,
            "columns": columns_meta,
            "rows": rows,
            "row_count": len(rows),
            "truncated": truncated,
            "columns_truncated": columns_truncated,
            "total_column_count": total_column_count,
            "execution_time_ms": elapsed_ms(),
        }

    except ValueError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}",
                "query_preview": text[:500]}
    except Exception as e:
        log.exception("execute_query failed")
        return {"error": f"Внутренняя ошибка: {e}"}


from copy import deepcopy
from threading import RLock

_metadata_cache: dict[tuple[str, str], dict] = {}
_metadata_cache_lock = RLock()
MAX_METADATA_CACHE = 512


def _cached_metadata_get(db_key: str, path: str) -> dict | None:
    with _metadata_cache_lock:
        value = _metadata_cache.get((db_key, path))
        return deepcopy(value) if value is not None else None


def _cached_metadata_put(db_key: str, path: str, result: dict) -> None:
    with _metadata_cache_lock:
        if len(_metadata_cache) >= MAX_METADATA_CACHE:
            _metadata_cache.clear()
        _metadata_cache[(db_key, path)] = deepcopy(result)

@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def describe_object(path: str, database: str | None = None) -> dict:
    """Возвращает структуру объекта метаданных конфигурации.

    Поддерживаются справочники, документы, регистры (накопления / сведений /
    бухгалтерии), перечисления, планы видов характеристик. Используй для разведки
    перед написанием запроса.

    Args:
        path: Полное имя, например "РегистрНакопления.Продажи",
              "Справочник.Контрагенты", "Документ.РеализацияТоваровУслуг".
        database: Ключ базы данных.
    """
    try:
        db_key = resolve_database(database)
        conn = get_connection(db_key)

        # Return cached metadata (invalidated on server restart or cache clear)
        cached = _cached_metadata_get(db_key, path)
        if cached is not None:
            return cached

        obj = resolve_metadata(conn, path)
        if obj is None:
            return {"error": f"Объект не найден: {path}", "database": db_key}

        kind = path.split(".")[0]
        result: dict = {"database": db_key, "path": path, "kind": kind,
                        "name": str(obj.Имя)}

        for prop, key in (("Синоним", "synonym"), ("Комментарий", "comment")):
            try:
                v = getattr(obj, prop)
                if v:
                    result[key] = str(v)
            except (AttributeError, pywintypes.com_error):
                pass

        if hasattr_safe(obj, "Реквизиты"):
            result["attributes"] = list_collection(obj.Реквизиты)
        if hasattr_safe(obj, "СтандартныеРеквизиты"):
            try:
                result["standard_attributes"] = [
                    str(a.Имя) for a in obj.СтандартныеРеквизиты
                ]
            except (AttributeError, pywintypes.com_error):
                pass
        if hasattr_safe(obj, "Измерения"):
            result["dimensions"] = list_collection(obj.Измерения)
        if hasattr_safe(obj, "Ресурсы"):
            result["resources"] = list_collection(obj.Ресурсы)
        if hasattr_safe(obj, "ТабличныеЧасти"):
            tps = []
            try:
                for tp in obj.ТабличныеЧасти:
                    tps.append({
                        "name": str(tp.Имя),
                        "attributes": list_collection(tp.Реквизиты),
                    })
            except pywintypes.com_error:
                pass
            if tps:
                result["tabular_sections"] = tps

        if kind in ("РегистрНакопления", "РегистрСведений", "РегистрБухгалтерии"):
            result["virtual_tables"] = virtual_tables_for(kind, obj)
        if kind == "Справочник":
            try:
                result["hierarchical"] = bool(obj.Иерархический)
            except (AttributeError, pywintypes.com_error):
                pass
        if kind == "Перечисление":
            try:
                result["values"] = [str(v.Имя) for v in obj.ЗначенияПеречисления]
            except (AttributeError, pywintypes.com_error):
                pass

        _cached_metadata_put(db_key, path, result)
        return result

    except ValueError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}
    except Exception as e:
        log.exception("describe_object failed")
        return {"error": f"Внутренняя ошибка: {e}"}


@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def list_metadata(
    metadata_type: str,
    name_filter: str | None = None,
    database: str | None = None,
) -> dict:
    """Список объектов метаданных указанной коллекции.

    Args:
        metadata_type: Имя коллекции (множественное число), например:
            "Справочники", "Документы", "РегистрыНакопления", "РегистрыСведений",
            "Перечисления", "ПланыВидовХарактеристик", "Константы", "Отчеты".
        name_filter: Подстрока для отбора по имени, регистр не учитывается.
        database: Ключ базы данных.
    """
    try:
        db_key = resolve_database(database)
        conn = get_connection(db_key)
        if not hasattr_safe(conn.Метаданные, metadata_type):
            return {
                "error": f"Коллекция метаданных не найдена: {metadata_type}",
                "database": db_key,
            }
        coll = getattr(conn.Метаданные, metadata_type)
        names = []
        f = name_filter.lower() if name_filter else None
        for o in coll:
            n = str(o.Имя)
            if f and f not in n.lower():
                continue
            names.append(n)
        names.sort()
        return {
            "database": db_key,
            "type": metadata_type,
            "count": len(names),
            "names": names,
        }
    except ValueError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}


@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def get_object_by_ref(
    uuid: str,
    type_path: str,
    database: str | None = None,
) -> dict:
    """Получает реквизиты объекта справочника или документа по UUID.

    Args:
        uuid: Уникальный идентификатор ссылки.
        type_path: "Справочник.Контрагенты" или "Документ.РеализацияТоваровУслуг".
        database: Ключ базы данных.
    """
    try:
        db_key = resolve_database(database)
        conn = get_connection(db_key)
        ref = get_ref_by_uuid(conn, type_path, uuid)
        try:
            obj = ref.ПолучитьОбъект()
        except pywintypes.com_error:
            obj = None
        if obj is None:
            return {"error": "Объект не существует или удалён", "database": db_key}

        result: dict = {"database": db_key, "_ref": uuid, "_type": type_path}
        for std in ("Код", "Наименование", "Номер", "Дата", "Проведен", "ПометкаУдаления"):
            try:
                result[std] = serialize_value(getattr(obj, std))
            except (AttributeError, pywintypes.com_error):
                pass
        try:
            for attr in obj.Метаданные().Реквизиты:
                n = str(attr.Имя)
                try:
                    result[n] = serialize_value(getattr(obj, n))
                except (AttributeError, pywintypes.com_error):
                    pass
        except (AttributeError, pywintypes.com_error):
            pass
        return result

    except ValueError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}
    except Exception as e:
        log.exception("get_object_by_ref failed")
        return {"error": f"Внутренняя ошибка: {e}"}


@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def list_databases() -> dict:
    """Возвращает список всех настроенных информационных баз с описаниями и заметками.

    Каждая база имеет короткое description и развёрнутое notes — что в этой базе
    можно найти. Используй этот инструмент когда нужно понять какая база подходит
    для конкретного вопроса.
    """
    return {
        "default_database": DB_CONFIG["default_database"],
        "databases": {
            key: {
                "description": cfg.get("description", key),
                "notes": cfg.get("notes", ""),
                "progid": cfg["progid"],
                "capabilities": {
                    "event_log": bool(cfg.get("allow_event_log", False)),
                },
            }
            for key, cfg in DB_CONFIG["databases"].items()
        },
    }


@mcp.tool(annotations=READ_ONLY_1C_TOOL)
def get_registration_log(
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
    limit: int = EVENT_LOG_DEFAULT_LIMIT,
    database: str | None = None,
) -> dict:
    """Читает отфильтрованные записи журнала регистрации 1С.

    Используй технические коды событий. Для отмены проведения документа —
    `_$Data$_.Unpost`. Инструмент только читает журнал и не изменяет базу.
    Обязательно укажи объектный фильтр (data_refs / document_numbers) или
    пару metadata_path+events: запрос только по периоду отклоняется.

    Args:
        start_datetime: начало периода, ISO "YYYY-MM-DDTHH:MM:SS" (локальное время ИБ).
        end_datetime: конец периода (включается), строго больше начала.
        events: технические коды событий (до 20). По умолчанию Post/Unpost/Update.
        metadata_path: полное имя, например "Документ.ПоступлениеДопРасходов".
        data_refs: до 50 ссылок {"_ref": "uuid", "_type": "Документ.Имя"}.
        document_numbers: до 20 номеров документов (только с metadata_path Документ.*
            и диапазоном document_date_from/document_date_to).
        document_date_from: начало диапазона даты самих документов.
        document_date_to: исключающая верхняя граница даты документов.
        successful_only: true — только зафиксированные/вне транзакции события.
        order: "asc" или "desc".
        limit: максимум строк (по умолчанию 200, hard cap 1000).
        database: ключ базы данных.

    Returns: request_id, period, resolved_objects, columns, rows, row_count,
        limit, truncated, warnings (или error_code/error).
    """
    request_id = secrets.token_hex(4)
    try:
        request = validate_event_log_request(
            start_datetime=start_datetime,
            end_datetime=end_datetime,
            events=events,
            metadata_path=metadata_path,
            data_refs=data_refs,
            document_numbers=document_numbers,
            document_date_from=document_date_from,
            document_date_to=document_date_to,
            successful_only=successful_only,
            order=order,
            limit=limit,
            max_range_days=EVENT_LOG_MAX_RANGE_DAYS,
            broad_range_days=EVENT_LOG_MAX_BROAD_RANGE_DAYS,
            max_documents=EVENT_LOG_MAX_DOCUMENTS,
            max_events=EVENT_LOG_MAX_EVENTS,
            max_refs=EVENT_LOG_MAX_REFS,
            hard_limit=EVENT_LOG_HARD_LIMIT,
        )
        db_key = resolve_database(database)

        # opt-in по базе — до установления COM-соединения
        if not DB_CONFIG["databases"][db_key].get("allow_event_log", False):
            return {
                "request_id": request_id,
                "database": db_key,
                "error_code": "EVENT_LOG_DISABLED",
                "error": "Чтение журнала регистрации не разрешено для базы.",
            }

        lock = _event_log_lock(db_key)
        if not lock.acquire(blocking=False):
            return {
                "request_id": request_id,
                "database": db_key,
                "error_code": "EVENT_LOG_BUSY",
                "error": "Другое чтение журнала этой базы уже выполняется. Повторите позже.",
            }
        try:
            conn = get_connection(db_key)

            # Разрешаем metadata_path через метаданные (не используем в запросе без валидации)
            if request.metadata_path and resolve_metadata(conn, request.metadata_path) is None:
                raise EventLogValidationError(
                    "METADATA_NOT_FOUND",
                    f"Объект метаданных не найден: {request.metadata_path}",
                )

            resolved_public: list[dict] = []
            if request.data_refs:
                resolved_refs = _resolve_data_refs(conn, request)
            elif request.document_numbers:
                resolved_refs, resolved_public = _resolve_document_numbers(conn, request)
            else:
                resolved_refs = []

            filter_1c = build_event_log_filter(conn, request, resolved_refs)
            table = conn.NewObject("ТаблицаЗначений")

            t0 = time.perf_counter()
            conn.ВыгрузитьЖурналРегистрации(
                table,
                filter_1c,
                EVENT_LOG_COLUMNS,
                "",  # текущий журнал; путь от клиента не принимается
                request.limit + 1,
            )
            rows, truncated = serialize_event_log_table(
                table, request.limit, conn
            )
            elapsed_ms = round((time.perf_counter() - t0) * 1000, 1)

            # Клиентский hard cap по размеру JSON-ответа
            if request.order == "desc":
                rows.reverse()  # сортировка по timestamp; при desc — разворот

            payload = {
                "request_id": request_id,
                "database": db_key,
                "period": request.period_payload(),
                "resolved_objects": resolved_public,
                "columns": list(EVENT_LOG_JSON_FIELDS),
                "rows": rows,
                "row_count": len(rows),
                "limit": request.limit,
                "truncated": truncated,
                "truncation_reason": "row_limit" if truncated else None,
                "execution_time_ms": elapsed_ms,
                "warnings": [],
            }

            # Лимит размера ответа: если превышен — усекаем и помечаем
            try:
                size = len(json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8"))
            except Exception:
                size = 0
            if size > EVENT_LOG_MAX_RESPONSE_BYTES:
                # Усекаем грубо: оставляем пустой список и помечаем причину
                payload["rows"] = []
                payload["row_count"] = 0
                payload["truncated"] = True
                payload["truncation_reason"] = "response_bytes"
                payload["warnings"] = [
                    "Ответ превысил лимит размера — результат усечён. Сузьте период."
                ]

            if not rows and not truncated:
                payload["warnings"] = [
                    "Записи не найдены. Проверьте период хранения и уровень регистрации событий."
                ]

            log_event_log_read(
                request_id=request_id,
                db_key=db_key,
                elapsed_ms=elapsed_ms,
                row_count=len(rows),
                truncated=truncated,
                fingerprint=filter_fingerprint(request.audit_payload()),
            )
            return payload
        finally:
            lock.release()
    except EventLogValidationError as exc:
        return {"request_id": request_id, "error_code": exc.code, "error": str(exc)}
    except EventLogAccessError as exc:
        return {
            "request_id": request_id,
            "error_code": "EVENT_LOG_ACCESS_DENIED",
            "error": str(exc),
        }
    except ValueError as exc:
        return {"request_id": request_id, "error_code": "VALIDATION_ERROR", "error": str(exc)}
    except pywintypes.com_error as exc:
        message = parse_com_error(exc)
        code = classify_event_log_com_error(message)
        if code == "EVENT_LOG_ACCESS_DENIED":
            code = "EVENT_LOG_ACCESS_DENIED"
        return {
            "request_id": request_id,
            "error_code": code,
            "error": safe_event_log_error(message),
        }
    except Exception as exc:
        log.exception("get_registration_log failed")
        return {
            "request_id": request_id,
            "error_code": "INTERNAL_ERROR",
            "error": f"Внутренняя ошибка: {type(exc).__name__}",
        }


# ---------------------------------------------------------------------------
# Постпатч описаний: дописываем _DB_BLOCK в description каждого tool'а
# в реестре FastMCP. Это гарантирует, что AI-клиент в tools/list увидит
# список баз и notes — даже без первого вызова инструмента.
# ---------------------------------------------------------------------------
def _patch_tool_descriptions():
    try:
        # FastMCP хранит инструменты во внутреннем реестре через ToolManager.
        tools = mcp._tool_manager._tools  # type: ignore[attr-defined]
        for name, tool in tools.items():
            try:
                if hasattr(tool, "description") and tool.description:
                    tool.description = tool.description.rstrip() + _DB_BLOCK
            except Exception as e:
                log.warning("Не удалось пропатчить описание %s: %s", name, e)
    except Exception as e:
        log.warning("Не удалось получить реестр инструментов FastMCP: %s", e)


_patch_tool_descriptions()


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    log.info("Стартую 1C MCP Bridge v%s", VERSION)
    log.info("Файл со списком баз: %s", find_databases_file())
    log.info("Базы: %s", list_database_keys())
    log.info("По умолчанию: %s", DB_CONFIG["default_database"])
    mcp.run()
