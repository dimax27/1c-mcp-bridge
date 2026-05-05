"""
1C MCP Bridge — сервер MCP для подключения Claude к 1С:Предприятию через COM.

v0.2.0: поддержка нескольких информационных баз. Список баз хранится в
databases.json (рядом с этим скриптом или в пути, заданном переменной
ONEC_DATABASES_FILE). Каждый из четырёх инструментов принимает опциональный
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
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP

# На Windows stdout/stderr по умолчанию в cp1251 — Claude Desktop пишет лог в
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

DEFAULT_LIMIT = int(os.environ.get("ONEC_DEFAULT_LIMIT", "1000"))
HARD_LIMIT = int(os.environ.get("ONEC_HARD_LIMIT", "10000"))


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

    # Фильтруем отключённые — Claude их вообще не должен видеть
    enabled_databases = {k: v for k, v in databases.items() if v.get("enabled", True)}
    if not enabled_databases:
        raise RuntimeError(
            f"{path}: все базы отключены (enabled=false). "
            f"Включи хотя бы одну в 1C Bridge Manager."
        )

    default_db = data.get("default_database") or next(iter(enabled_databases))
    if default_db not in enabled_databases:
        # default отключена — берём первую включённую
        default_db = next(iter(enabled_databases))

    return {
        "default_database": default_db,
        "databases": enabled_databases,
    }


# Загружаем единожды при старте; перечитывать на лету не будем — переустановка
# всё равно требует перезапуска Claude Desktop.
try:
    DB_CONFIG = load_databases()
except Exception as e:
    log.error("Не удалось загрузить databases.json: %s", e)
    log.error("Сервер запускается, но все вызовы инструментов будут падать.")
    DB_CONFIG = {"default_database": "", "databases": {}}


def list_database_keys() -> list[str]:
    return list(DB_CONFIG["databases"].keys())


def get_db_descriptions() -> str:
    """Полное описание баз для подсказки Claude в tools/list.
    Включает короткое description и длинные notes (если есть) — Claude
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
    conn_str = cfg["connection_string"]

    log.info("Подключаюсь к '%s' через %s", db_key, progid)
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
# MCP server + tools
# ---------------------------------------------------------------------------

mcp = FastMCP("1c-bridge")

# Описание баз — попадает в docstring каждого инструмента и Claude видит его
# в tools/list ДО первого вопроса от пользователя.
_DB_BLOCK = "\n\n" + get_db_descriptions() + "\n"


def _with_db_info(doc: str) -> str:
    """Дописывает информацию о базах в конец docstring."""
    return doc.rstrip() + _DB_BLOCK


@mcp.tool()
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
    limit = min(max(1, int(limit)), HARD_LIMIT)

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
                "execution_time_ms": elapsed_ms(),
            }

        columns_meta = []
        col_names = []
        for col in result.Колонки:
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


@mcp.tool()
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

        return result

    except ValueError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}
    except Exception as e:
        log.exception("describe_object failed")
        return {"error": f"Внутренняя ошибка: {e}"}


@mcp.tool()
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


@mcp.tool()
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


@mcp.tool()
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
            }
            for key, cfg in DB_CONFIG["databases"].items()
        },
    }


# ---------------------------------------------------------------------------
# execute_code — запись/изменение данных через файл-дроппер
# ---------------------------------------------------------------------------
#
# Канал: Claude → этот инструмент → JSON в inbox-папке → 1С обработка
# ИИ_Песочница читает inbox, выполняет код, пишет результат в done/errors.
# Этот инструмент возвращает результат когда обработка его создаст,
# либо отдаёт таймаут.

import uuid
import time as _time

DEFAULT_DROPPER_TIMEOUT = 60  # секунд


def get_dropper_path(db_key: str) -> Path:
    """Папка дроппера для базы. Берётся либо из конфига базы (поле dropper_path),
    либо строится по умолчанию: %LOCALAPPDATA%/1cMcpBridge/dropper/<db_key>.
    """
    cfg = DB_CONFIG["databases"].get(db_key, {})
    custom = cfg.get("dropper_path", "").strip()
    if custom:
        return Path(custom)
    local_app = os.environ.get("LOCALAPPDATA", "")
    if not local_app:
        local_app = str(Path.home() / "AppData" / "Local")
    return Path(local_app) / "1cMcpBridge" / "dropper" / db_key


@mcp.tool()
def execute_code(
    code_forward: str,
    code_backward: str,
    comment: str,
    database: str | None = None,
    timeout_seconds: int = DEFAULT_DROPPER_TIMEOUT,
    request_screenshot: bool = False,
) -> dict:
    """Выполнить произвольный код 1С (в тестовой базе!) с записью в журнал ИИ_Действия.

    Этот инструмент НЕ выполняет код напрямую через COM. Он кладёт JSON-задание
    в папку inbox дроппера, и обработка ИИ_Песочница в 1С (которая должна быть
    запущена пользователем и желательно в авто-режиме) выполнит код и положит
    результат в done или errors.

    БЕЗОПАСНОСТЬ:
    - Используй только в БАЗАХ С МАРКЕРОМ [Тест] в названии. Если в базе нет
      этого маркера, обработка покажет предупреждение, но всё равно выполнит код.
    - Каждое выполнение оборачивается в транзакцию с автоматическим rollback
      при ошибке.
    - Все действия логируются в регистр сведений ИИ_Действия в самой 1С базе.
    - code_backward — это компенсирующий код для отката. Он сохраняется в
      регистре и пользователь может вручную запустить его через обработку
      ИИ_Песочница если нужно.

    КАК ВЕРНУТЬ ВЫВОД:
    - В коде используй `ИИ_Сервер.Вывести("текст")` чтобы передать что-то
      в результат. Можно вызывать многократно, строки склеятся через перевод.

    Args:
        code_forward: Код 1С для выполнения (на русском, как в Конфигураторе).
        code_backward: Код для отката (компенсирующий). Можно пустую строку
            если откат невозможен — но тогда отметь это в comment.
        comment: Комментарий о том, что делает этот код. Попадёт в журнал.
            Например: "Заполняю ИНН для контрагентов с пустым ИНН по списку".
        database: Ключ базы из list_databases. Если None — база по умолчанию.
        timeout_seconds: Сколько ждать ответа от 1С (по умолчанию 60).
        request_screenshot: Если True — обработка попытается сохранить скриншот
            экрана после выполнения (для проверки UI-результата).

    Returns:
        Структура с полями:
            task_id: уникальный идентификатор задания
            status: "success" / "runtime_error" / "rollback_failed" / "timeout"
                / "rejected_by_safety"
            output: что вернул код через ИИ_Сервер.Вывести()
            error: текст ошибки если упало
            duration_ms: сколько выполнялось
            warnings: список предупреждений (например, что база не помечена [Тест])
    """
    if not code_forward or not code_forward.strip():
        return {"error": "code_forward пустой"}

    try:
        db_key = resolve_database(database)
    except ValueError as e:
        return {"error": str(e)}

    dropper = get_dropper_path(db_key)
    inbox = dropper / "inbox"
    done = dropper / "done"
    errors = dropper / "errors"

    # Создаём папки если нет (могут быть созданы заранее, тут на всякий случай)
    for folder in (inbox, done, errors):
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            return {
                "error": f"Не удалось создать папку дроппера {folder}: {e}",
                "hint": "Убедись что путь к дропперу настроен правильно в Manager."
            }

    task_id = str(uuid.uuid4())
    timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"{timestamp}-{task_id}.json"

    task = {
        "task_id": task_id,
        "comment": comment,
        "code_forward": code_forward,
        "code_backward": code_backward,
        "request_screenshot": request_screenshot,
        "created_at": datetime.datetime.now().isoformat(),
        "timeout_seconds": timeout_seconds,
    }

    task_file = inbox / filename
    try:
        task_file.write_text(
            json.dumps(task, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        log.info("execute_code: положил задание %s в %s", task_id, task_file)
    except Exception as e:
        return {"error": f"Не удалось записать задание в inbox: {e}"}

    # Ожидание результата — опрашиваем done и errors каждые 500мс
    result_filename = f"{timestamp}-{task_id}.result.json"
    deadline = _time.time() + timeout_seconds

    while _time.time() < deadline:
        for folder in (done, errors):
            result_path = folder / result_filename
            if result_path.exists():
                try:
                    result = json.loads(result_path.read_text(encoding="utf-8"))
                    result["database"] = db_key
                    return result
                except Exception as e:
                    return {
                        "error": f"Не удалось прочитать result-файл: {e}",
                        "task_id": task_id,
                        "database": db_key,
                    }
        _time.sleep(0.5)

    # Таймаут
    return {
        "task_id": task_id,
        "status": "timeout",
        "error": (
            f"Прошло {timeout_seconds} секунд, обработка ИИ_Песочница не вернула результат. "
            f"Проверь что обработка открыта в 1С и включён авто-режим, либо нажми "
            f"«Выполнить выделенное» вручную."
        ),
        "task_file": str(task_file),
        "database": db_key,
    }


# ---------------------------------------------------------------------------
# Постпатч описаний: дописываем _DB_BLOCK в description каждого tool'а
# в реестре FastMCP. Это гарантирует, что Claude в tools/list увидит
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
    log.info("Стартую 1C MCP Bridge v0.2.0")
    log.info("Файл со списком баз: %s", find_databases_file())
    log.info("Базы: %s", list_database_keys())
    log.info("По умолчанию: %s", DB_CONFIG["default_database"])
    mcp.run()
