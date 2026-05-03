"""
MCP-сервер для подключения Claude к базе 1С 8.2 (Управление торговлей 10.3).

Архитектура: Claude Desktop запускает этот скрипт через stdio. Инструменты,
зарегистрированные через FastMCP, вызываются по протоколу MCP (JSON-RPC поверх
stdio). Сервер держит COM-соединение с информационной базой через V82.COMConnector.

Конфигурация — через переменные окружения. См. .env.example и README.md.
"""

from __future__ import annotations

import datetime
import io
import logging
import os
import sys
import threading
import time
from typing import Any

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP

# На Windows stdout/stderr по умолчанию в cp1251 — Claude Desktop пишет лог в
# UTF-8 и ломается на кириллице. Принудительно перекодируем оба потока.
if sys.platform == "win32":
    try:
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", line_buffering=True
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", line_buffering=True
        )
    except (AttributeError, ValueError):
        # Если потоки уже обёрнуты (тесты и т.п.) — пропускаем
        pass

# Логи направляем в stderr — stdout занят протоколом MCP (JSON-RPC).
logging.basicConfig(
    level=os.environ.get("ONEC_LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("mcp-1c")

# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

CONN_STRING = os.environ.get("ONEC_CONNECTION_STRING", "").strip()

# Альтернатива — собрать строку из частей.
ONEC_MODE = os.environ.get("ONEC_MODE", "").lower()      # "file" | "server"
ONEC_FILE = os.environ.get("ONEC_FILE_PATH", "")
ONEC_SRVR = os.environ.get("ONEC_SERVER", "")
ONEC_REF = os.environ.get("ONEC_REF", "")
ONEC_USER = os.environ.get("ONEC_USER", "")
ONEC_PWD = os.environ.get("ONEC_PASSWORD", "")

DEFAULT_LIMIT = int(os.environ.get("ONEC_DEFAULT_LIMIT", "1000"))
HARD_LIMIT = int(os.environ.get("ONEC_HARD_LIMIT", "10000"))
COM_PROGID = os.environ.get("ONEC_COMCONNECTOR_PROGID", "V82.COMConnector")


def build_connection_string() -> str:
    if CONN_STRING:
        return CONN_STRING
    if ONEC_MODE == "file":
        return f'File="{ONEC_FILE}";Usr="{ONEC_USER}";Pwd="{ONEC_PWD}"'
    if ONEC_MODE == "server":
        return (
            f'Srvr="{ONEC_SRVR}";Ref="{ONEC_REF}";'
            f'Usr="{ONEC_USER}";Pwd="{ONEC_PWD}"'
        )
    raise RuntimeError(
        "Не задана конфигурация подключения. Установите ONEC_CONNECTION_STRING "
        "или (ONEC_MODE=file|server + соответствующие переменные)."
    )


# ---------------------------------------------------------------------------
# Управление COM (потокобезопасно)
# ---------------------------------------------------------------------------
#
# FastMCP может выполнять синхронные обработчики в нескольких потоках executor'а.
# COM-объекты привязаны к потоку (STA), поэтому держим CoInitialize и connection
# в thread-local. На практике пул редко превышает 1–2 потока.

_tls = threading.local()


def _ensure_com() -> None:
    if not getattr(_tls, "com_init", False):
        pythoncom.CoInitialize()
        _tls.com_init = True


def get_connection() -> Any:
    _ensure_com()
    conn = getattr(_tls, "connection", None)
    if conn is not None:
        # Простая проверка живости.
        try:
            _ = conn.Метаданные.Имя
            return conn
        except pywintypes.com_error:
            log.warning("COM-соединение умерло, переподключаюсь")
            _tls.connection = None
    log.info("Создаю COM-соединение к 1С (%s)", COM_PROGID)
    connector = win32com.client.Dispatch(COM_PROGID)
    _tls.connection = connector.Connect(build_connection_string())
    log.info("Соединение установлено: ИБ %s", _tls.connection.Метаданные.Имя)
    return _tls.connection


# ---------------------------------------------------------------------------
# Сериализация значений
# ---------------------------------------------------------------------------

EMPTY_DATE_YEAR = 1900


def serialize_value(v: Any, depth: int = 0) -> Any:
    """Превращает COM-значение в JSON-совместимое."""
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
    # Дальше идут объекты 1С: ссылки, перечисления и пр.
    if depth > 1:
        try:
            return str(v)
        except Exception:
            return f"<COM:{type(v).__name__}>"
    # Ссылка на справочник/документ?
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
    # Значение перечисления?
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
    """JSON-значение -> тип, понятный 1С."""
    if isinstance(value, dict) and "_ref" in value and "_type" in value:
        return get_ref_by_uuid(conn, value["_type"], value["_ref"])
    if isinstance(value, str):
        # Пытаемся распознать ISO-дату.
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
    """'Справочник.Контрагенты' + UUID -> ссылка на объект."""
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
# Хелперы метаданных
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
        # Fallback: ищем перебором.
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
        return ["Остатки", "Обороты", "ОстаткиИОбороты", "ОборотыДтКт", "Движения", "ДвиженияССубконто"]
    return []


def parse_com_error(e: pywintypes.com_error) -> str:
    """Достать читаемое сообщение из pywintypes.com_error."""
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


@mcp.tool()
def execute_query(
    text: str,
    parameters: dict | None = None,
    limit: int = DEFAULT_LIMIT,
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

    Returns:
        columns: схема результата [{name, type}, ...]
        rows: список объектов {колонка: значение}
        row_count, truncated, execution_time_ms
    """
    if not text or not text.strip():
        return {"error": "Пустой текст запроса"}

    limit = min(max(1, int(limit)), HARD_LIMIT)

    try:
        conn = get_connection()
        query = conn.NewObject("Запрос")
        query.Текст = text

        if parameters:
            for name, raw in parameters.items():
                try:
                    parsed = parse_parameter(raw, conn)
                    query.УстановитьПараметр(name, parsed)
                except Exception as e:
                    return {"error": f"Параметр '{name}': {e}"}

        t0 = time.perf_counter()
        result = query.Выполнить()
        elapsed_ms = lambda: round((time.perf_counter() - t0) * 1000, 1)

        try:
            empty = bool(result.Пустой())
        except (AttributeError, pywintypes.com_error):
            empty = False

        if empty:
            return {
                "columns": [],
                "rows": [],
                "row_count": 0,
                "truncated": False,
                "execution_time_ms": elapsed_ms(),
            }

        columns_meta: list[dict] = []
        col_names: list[str] = []
        for col in result.Колонки:
            n = str(col.Имя)
            col_names.append(n)
            try:
                t_str = str(col.ТипЗначения)
            except Exception:
                t_str = ""
            columns_meta.append({"name": n, "type": t_str})

        selection = result.Выбрать()
        rows: list[dict] = []
        truncated = False
        while selection.Следующий():
            if len(rows) >= limit:
                truncated = True
                break
            row: dict = {}
            for cn in col_names:
                try:
                    val = getattr(selection, cn)
                except (AttributeError, pywintypes.com_error):
                    val = None
                row[cn] = serialize_value(val)
            rows.append(row)

        return {
            "columns": columns_meta,
            "rows": rows,
            "row_count": len(rows),
            "truncated": truncated,
            "execution_time_ms": elapsed_ms(),
        }

    except pywintypes.com_error as e:
        msg = parse_com_error(e)
        log.error("Ошибка запроса: %s", msg)
        return {"error": f"Ошибка 1С: {msg}", "query_preview": text[:500]}
    except Exception as e:
        log.exception("execute_query: непредвиденная ошибка")
        return {"error": f"Внутренняя ошибка: {e}"}


@mcp.tool()
def describe_object(path: str) -> dict:
    """Возвращает структуру объекта метаданных конфигурации.

    Поддерживаются справочники, документы, регистры (накопления / сведений /
    бухгалтерии), перечисления, планы видов характеристик. Используй для разведки
    перед написанием запроса — какие у регистра измерения и ресурсы, какие
    реквизиты у документа и т.п.

    Args:
        path: Полное имя, например "РегистрНакопления.Продажи",
              "Справочник.Контрагенты", "Документ.РеализацияТоваровУслуг".
    """
    try:
        conn = get_connection()
        obj = resolve_metadata(conn, path)
        if obj is None:
            return {"error": f"Объект не найден: {path}"}

        kind = path.split(".")[0]
        result: dict = {"path": path, "kind": kind, "name": str(obj.Имя)}

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

    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}
    except Exception as e:
        log.exception("describe_object failed")
        return {"error": f"Внутренняя ошибка: {e}"}


@mcp.tool()
def list_metadata(metadata_type: str, name_filter: str | None = None) -> dict:
    """Список объектов метаданных указанной коллекции.

    Args:
        metadata_type: Имя коллекции (множественное число), например:
            "Справочники", "Документы", "РегистрыНакопления", "РегистрыСведений",
            "Перечисления", "ПланыВидовХарактеристик", "Константы", "Отчеты".
        name_filter: Подстрока для отбора по имени, регистр не учитывается.

    Returns:
        type, count, names: list[str]
    """
    try:
        conn = get_connection()
        if not hasattr_safe(conn.Метаданные, metadata_type):
            return {"error": f"Коллекция метаданных не найдена: {metadata_type}"}
        coll = getattr(conn.Метаданные, metadata_type)
        names: list[str] = []
        f = name_filter.lower() if name_filter else None
        for o in coll:
            n = str(o.Имя)
            if f and f not in n.lower():
                continue
            names.append(n)
        names.sort()
        return {"type": metadata_type, "count": len(names), "names": names}
    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}


@mcp.tool()
def get_object_by_ref(uuid: str, type_path: str) -> dict:
    """Получает реквизиты объекта справочника или документа по UUID.

    Args:
        uuid: Уникальный идентификатор ссылки.
        type_path: "Справочник.Контрагенты" или "Документ.РеализацияТоваровУслуг".
    """
    try:
        conn = get_connection()
        ref = get_ref_by_uuid(conn, type_path, uuid)
        try:
            obj = ref.ПолучитьОбъект()
        except pywintypes.com_error:
            obj = None
        if obj is None:
            return {"error": "Объект не существует или удалён"}

        result: dict = {"_ref": uuid, "_type": type_path}
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

    except pywintypes.com_error as e:
        return {"error": f"Ошибка 1С: {parse_com_error(e)}"}
    except Exception as e:
        log.exception("get_object_by_ref failed")
        return {"error": f"Внутренняя ошибка: {e}"}


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    log.info("Стартую MCP-сервер 1С (%s)", COM_PROGID)
    try:
        # Прогреваем соединение, чтобы упасть рано в случае ошибки конфига.
        get_connection()
    except Exception as e:
        log.error("Не удалось подключиться к 1С при старте: %s", e)
        log.error("Сервер всё равно запускается — попробуем при первом вызове.")
    mcp.run()
