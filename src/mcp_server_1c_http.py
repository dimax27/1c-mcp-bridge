"""
1C MCP Bridge — HTTP/SSE version for Qwen Desktop.
Usage: python mcp_server_1c_http.py [--port 8000]
Connect in Qwen: Streamable HTTP → http://127.0.0.1:8000/mcp
"""
from __future__ import annotations
import argparse, datetime, io, json, logging, os, sys, threading, time
from pathlib import Path
from typing import Any
import pythoncom, pywintypes, win32com.client
from mcp.server import MCPServer
from starlette.applications import Starlette
from starlette.responses import JSONResponse, Response
from starlette.routing import Route

logging.basicConfig(level=os.environ.get("ONEC_LOG_LEVEL","INFO"), format="%(asctime)s [%(levelname)s] %(name)s: %(message)s", stream=sys.stderr)
log = logging.getLogger("mcp-1c-http")

DEFAULT_LIMIT = int(os.environ.get("ONEC_DEFAULT_LIMIT", "1000"))
HARD_LIMIT = int(os.environ.get("ONEC_HARD_LIMIT", "10000"))
EMPTY_DATE_YEAR = 1900

def find_databases_file() -> Path:
    env_path = os.environ.get("ONEC_DATABASES_FILE", "").strip()
    if env_path: return Path(env_path)
    standard = Path(os.environ.get("PROGRAMDATA", "C:/ProgramData")) / "1cMcpBridge" / "databases.json"
    if standard.exists(): return standard
    return standard

def load_databases() -> dict:
    path = find_databases_file()
    data = json.loads(path.read_text(encoding="utf-8"))
    databases = data.get("databases", {})
    for key, cfg in databases.items():
        cfg.setdefault("description", key); cfg.setdefault("notes", ""); cfg.setdefault("enabled", True)
    enabled = {k: v for k, v in databases.items() if v.get("enabled", True)}
    default_db = data.get("default_database") or next(iter(enabled))
    if default_db not in enabled: default_db = next(iter(enabled))
    return {"default_database": default_db, "databases": enabled}

DB_CONFIG = load_databases()

def get_db_descriptions() -> str:
    lines = ["Configured databases:"]
    default = DB_CONFIG["default_database"]
    for key, cfg in DB_CONFIG["databases"].items():
        marker = " [default]" if key == default else ""
        lines.append(f"\n  • '{key}'{marker}: {cfg['description']}")
        for note_line in cfg.get("notes", "").strip().splitlines():
            lines.append(f"      {note_line}")
    return "\n".join(lines)

_DB_BLOCK = "\n\n" + get_db_descriptions()

_tls = threading.local()
def _ensure_com():
    if not getattr(_tls, "com_init", False): pythoncom.CoInitialize(); _tls.com_init = True

def get_connection(db_key: str) -> Any:
    _ensure_com()
    if not hasattr(_tls, "connections"): _tls.connections = {}
    conn = _tls.connections.get(db_key)
    if conn is not None:
        try:
            _ = conn.Метаданные.Имя; return conn
        except pywintypes.com_error:
            _tls.connections.pop(db_key, None)
    cfg = DB_CONFIG["databases"][db_key]
    connector = win32com.client.Dispatch(cfg["progid"])
    conn = connector.Connect(cfg["connection_string"])
    _tls.connections[db_key] = conn
    return conn

def resolve_database(db_param):
    if not db_param: return DB_CONFIG["default_database"]
    if db_param not in DB_CONFIG["databases"]: raise ValueError(f"Database '{db_param}' not found")
    return db_param

def serialize_value(v, depth=0):
    if v is None: return None
    if isinstance(v, bool): return v
    if isinstance(v, (int, float, str)): return v
    if isinstance(v, datetime.datetime): return None if v.year < EMPTY_DATE_YEAR else v.isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    if isinstance(v, pywintypes.TimeType):
        try: d = datetime.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second); return None if d.year < EMPTY_DATE_YEAR else d.isoformat()
        except: return str(v)
    if depth > 1: return str(v)
    try:
        uuid = str(v.УникальныйИдентификатор())
        try: type_name = str(v.Метаданные().ПолноеИмя())
        except: type_name = None
        try: presentation = str(v)
        except: presentation = None
        return {"_ref": uuid, "_type": type_name, "_presentation": presentation}
    except (AttributeError, pywintypes.com_error): pass
    try: return {"_enum": str(v.Метаданные().ПолноеИмя()), "_value": str(v)}
    except: pass
    return str(v)

def parse_parameter(value, conn):
    if isinstance(value, dict) and "_ref" in value:
        parts = value["_type"].split(".")
        kind, name = parts
        coll_map = {"Справочник": "Справочники", "Документ": "Документы", "Перечисление": "Перечисления"}
        manager = getattr(getattr(conn, coll_map[kind]), name)
        return manager.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", value["_ref"]))
    if isinstance(value, str):
        try:
            if len(value) == 10 and value[4] == "-" and value[7] == "-": return datetime.datetime.fromisoformat(value)
            if "T" in value: return datetime.datetime.fromisoformat(value.replace("Z", ""))
        except: pass
    if isinstance(value, list):
        arr = conn.NewObject("Массив")
        for item in value: arr.Добавить(parse_parameter(item, conn))
        return arr
    return value

def parse_com_error(e):
    try: return str(e.excepinfo[2]).strip() if e.excepinfo and len(e.excepinfo) > 2 else str(e)
    except: return str(e)

mcp = MCPServer("1c-bridge")

@mcp.tool()
def list_databases() -> dict:
    """Return the list of all configured 1C databases with descriptions."""
    return {"default_database": DB_CONFIG["default_database"], "databases": {key: {"description": cfg.get("description", key), "notes": cfg.get("notes", ""), "progid": cfg["progid"]} for key, cfg in DB_CONFIG["databases"].items()}}

@mcp.tool()
def execute_query(text: str, parameters: dict | None = None, limit: int = 1000, database: str | None = None) -> dict:
    """Execute a 1C query and return tabular results. Uses 1C query language (Russian keywords: ВЫБРАТЬ, ИЗ, ГДЕ)."""
    if not text: return {"error": "Empty query text"}
    limit = min(max(1, int(limit)), HARD_LIMIT)
    try:
        db_key = resolve_database(database); conn = get_connection(db_key)
        query = conn.NewObject("Запрос"); query.Текст = text
        if parameters:
            for name, raw in parameters.items():
                try: query.УстановитьПараметр(name, parse_parameter(raw, conn))
                except Exception as e: return {"error": f"Parameter '{name}': {e}", "database": db_key}
        t0 = time.perf_counter(); result = query.Выполнить(); elapsed_ms = round((time.perf_counter() - t0) * 1000, 1)
        try:
            if result.Пустой(): return {"database": db_key, "columns": [], "rows": [], "row_count": 0, "truncated": False, "execution_time_ms": elapsed_ms}
        except: pass
        columns_meta, col_names = [], []
        for col in result.Колонки:
            n = str(col.Имя); col_names.append(n)
            try: columns_meta.append({"name": n, "type": str(col.ТипЗначения)})
            except: columns_meta.append({"name": n, "type": ""})
        selection = result.Выбрать(); rows, truncated = [], False
        while selection.Следующий():
            if len(rows) >= limit: truncated = True; break
            rows.append({cn: serialize_value(getattr(selection, cn)) for cn in col_names})
        return {"database": db_key, "columns": columns_meta, "rows": rows, "row_count": len(rows), "truncated": truncated, "execution_time_ms": elapsed_ms}
    except ValueError as e: return {"error": str(e)}
    except pywintypes.com_error as e: return {"error": f"1C error: {parse_com_error(e)}", "query_preview": text[:500]}
    except Exception as e: log.exception("execute_query"); return {"error": f"Internal error: {e}"}

@mcp.tool()
def describe_object(path: str, database: str | None = None) -> dict:
    """Return metadata structure of a configuration object (Справочник, Документ, Регистр, etc.)."""
    try:
        db_key = resolve_database(database); conn = get_connection(db_key)
        parts = path.split(".")
        coll_map = {"Справочник": "Справочники", "Документ": "Документы", "РегистрНакопления": "РегистрыНакопления", "РегистрСведений": "РегистрыСведений", "РегистрБухгалтерии": "РегистрыБухгалтерии", "Перечисление": "Перечисления", "ПланВидовХарактеристик": "ПланыВидовХарактеристик"}
        kind, name = parts
        obj = None
        for o in getattr(conn.Метаданные, coll_map[kind]):
            if str(o.Имя) == name: obj = o; break
        if obj is None: return {"error": f"Object not found: {path}", "database": db_key}
        result = {"database": db_key, "path": path, "kind": kind, "name": str(obj.Имя)}
        for prop, key in (("Синоним", "synonym"), ("Комментарий", "comment")):
            try:
                v = getattr(obj, prop)
                if v: result[key] = str(v)
            except: pass
        for attr_grp in ["Реквизиты", "Измерения", "Ресурсы"]:
            try:
                items = []
                for item in getattr(obj, attr_grp):
                    entry = {"name": str(item.Имя)}
                    try:
                        t = getattr(item, "Тип")
                        if t: entry["type"] = str(t)
                    except: pass
                    items.append(entry)
                key_map = {"Реквизиты": "attributes", "Измерения": "dimensions", "Ресурсы": "resources"}
                result[key_map[attr_grp]] = items
            except: pass
        return result
    except ValueError as e: return {"error": str(e)}
    except pywintypes.com_error as e: return {"error": f"1C error: {parse_com_error(e)}"}
    except Exception as e: log.exception("describe_object"); return {"error": f"Internal error: {e}"}

@mcp.tool()
def list_metadata(metadata_type: str, name_filter: str | None = None, database: str | None = None) -> dict:
    """List metadata objects of a given collection (Справочники, Документы, РегистрыНакопления, etc.)."""
    try:
        db_key = resolve_database(database); conn = get_connection(db_key)
        coll = getattr(conn.Метаданные, metadata_type)
        names = [str(o.Имя) for o in coll if not name_filter or name_filter.lower() in str(o.Имя).lower()]
        names.sort()
        return {"database": db_key, "type": metadata_type, "count": len(names), "names": names}
    except ValueError as e: return {"error": str(e)}
    except pywintypes.com_error as e: return {"error": f"1C error: {parse_com_error(e)}"}

@mcp.tool()
def get_object_by_ref(uuid: str, type_path: str, database: str | None = None) -> dict:
    """Get object details by UUID. type_path: 'Справочник.Контрагенты' or 'Документ.РеализацияТоваровУслуг'."""
    try:
        db_key = resolve_database(database); conn = get_connection(db_key)
        parts = type_path.split(".")
        coll_map = {"Справочник": "Справочники", "Документ": "Документы"}
        manager = getattr(getattr(conn, coll_map[parts[0]]), parts[1])
        ref = manager.ПолучитьСсылку(conn.NewObject("УникальныйИдентификатор", uuid))
        try: obj = ref.ПолучитьОбъект()
        except: return {"error": "Object does not exist or deleted", "database": db_key}
        result = {"database": db_key, "_ref": uuid, "_type": type_path}
        for std in ("Код", "Наименование", "Номер", "Дата", "Проведен", "ПометкаУдаления"):
            try: result[std] = serialize_value(getattr(obj, std))
            except: pass
        return result
    except ValueError as e: return {"error": str(e)}
    except pywintypes.com_error as e: return {"error": f"1C error: {parse_com_error(e)}"}
    except Exception as e: log.exception("get_object_by_ref"); return {"error": f"Internal error: {e}"}

def _patch_tool_descriptions():
    try:
        for name, tool in mcp._tool_manager._tools.items():
            if hasattr(tool, "description") and tool.description:
                tool.description = tool.description.rstrip() + _DB_BLOCK
    except Exception as e: log.warning("Patch failed: %s", e)
_patch_tool_descriptions()

# --- HTTP server ---
async def handle_mcp(request):
    from mcp.server.streamable_http import StreamableHTTPServerTransport
    transport = StreamableHTTPServerTransport("/mcp")
    async with transport.connect() as (read_stream, write_stream):
        await mcp.run(read_stream, write_stream)
    return Response(status_code=200)

async def health(request):
    return JSONResponse({"status": "ok", "databases": list(DB_CONFIG["databases"].keys())})

def create_app():
    return Starlette(routes=[Route("/mcp", endpoint=handle_mcp, methods=["GET", "POST"]), Route("/health", endpoint=health)])

def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--port", type=int, default=8000); parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()
    log.info("Starting 1C MCP Bridge HTTP on %s:%d", args.host, args.port)
    log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))
    import uvicorn; uvicorn.run(create_app(), host=args.host, port=args.port, log_level="info")

if __name__ == "__main__": main()
