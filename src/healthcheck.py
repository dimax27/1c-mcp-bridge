"""Единый MCP health-check для HTTP-сервера моста.

Одна и та же проверка используется установщиком (install.ps1) и служебными
скриптами (restart_http_server.ps1), чтобы проверки не расходились:

  * сервер отвечает на streamable-http /mcp/<токен>;
  * tools/list возвращает ВСЕ пять инструментов моста;
  * list_databases вызывается без ошибки.

Режим `--com` дополнительно проверяет COM-подключение и метаданные каждой
включённой базы: вызывает list_metadata с заведомо несуществующим фильтром и
проверяет, что сервер вернул правильный тип коллекции без ошибки.

Запуск:   python healthcheck.py [--com] [--database <key>]
          (ONEC_HTTP_TOKEN — в переменной окружения, порт — ONEC_HTTP_PORT)

Код:      0 = проверка пройдена;
          1 = сервер недоступен / нет токена / инструменты не те;
          2 = list_databases вернул ошибку или список баз пуст/некорректен;
          3 = COM-проба (--com) не прошла хотя бы для одной базы.
Вывод — ASCII-маркеры (HEALTH_OK / HEALTH_TOOLS_MISSING / HEALTH_COM_FAIL...),
имена баз и ошибок экранируются (ensure_ascii), безопасно для журналов.
"""

import argparse
import asyncio
import json
import os
import traceback

EXPECTED_TOOLS = {
    "describe_object",
    "execute_query",
    "get_object_by_ref",
    "list_databases",
    "list_metadata",
}

# "Справочники" через unicode-escape: ASCII-safe, не зависит от кодировки консоли.
METADATA_DIRECTORIES = (
    "\u0421\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a\u0438"
)
PROBE_FILTER = "__1C_BRIDGE_CONNECTION_PROBE__"


def _result_text(result) -> str:
    return "".join(
        getattr(item, "text", "") or ""
        for item in getattr(result, "content", []) or []
    )


def ascii_json(value) -> str:
    """Компактный JSON с ensure_ascii: ключи баз и ошибки всегда ASCII."""
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"))


async def _probe_com(client, databases: list[str]) -> list[str]:
    """Проверяет COM/метаданные каждой базы. Возвращает список неудач."""
    failures: list[str] = []
    for db in databases:
        try:
            res = await client.call_tool(
                "list_metadata",
                {
                    "metadata_type": METADATA_DIRECTORIES,
                    "name_filter": PROBE_FILTER,
                    "database": db,
                },
            )
            text = _result_text(res)
            payload = json.loads(text) if text else {}
            ok = (
                not getattr(res, "is_error", False)
                and not payload.get("error")
                and payload.get("type") == METADATA_DIRECTORIES
            )
        except Exception:  # noqa: BLE001 — любая ошибка = провал пробы
            ok = False
        print(f"HEALTH_COM database={ascii_json(db)} status={'OK' if ok else 'FAIL'}")
        if not ok:
            failures.append(db)
    return failures


def check_databases_payload(payload: dict) -> tuple[int, str, list[str]]:
    """Валидация ответа list_databases.

    Возвращает (код, маркер, список ключей баз): 0 — всё корректно;
    2 — список баз пуст/не является объектом/нет default_database.
    """
    databases_payload = payload.get("databases")
    if not isinstance(databases_payload, dict):
        return 2, "HEALTH_DATABASES_INVALID", []
    if not databases_payload:
        return 2, "HEALTH_NO_DATABASES", []
    if payload.get("default_database") not in databases_payload:
        return 2, "HEALTH_DEFAULT_DATABASE_INVALID", []
    return 0, "", sorted(databases_payload.keys())


async def run(args) -> int:
    token = os.environ.get("ONEC_HTTP_TOKEN", "").strip()
    if not token:
        print("HEALTH_NO_TOKEN")
        return 1
    port = os.environ.get("ONEC_HTTP_PORT", "8000").strip() or "8000"
    from mcp import Client

    url = f"http://127.0.0.1:{port}/mcp/{token}"
    try:
        async with Client(url) as client:
            tools = await client.list_tools()
            names = {t.name for t in tools.tools}
            missing = EXPECTED_TOOLS - names
            if missing:
                print("HEALTH_TOOLS_MISSING:", ",".join(sorted(missing)))
                return 1
            result = await client.call_tool("list_databases", {})
            if getattr(result, "is_error", False):
                print("HEALTH_LIST_DATABASES_ERROR")
                return 2
            text = _result_text(result)
            payload = json.loads(text) if text else {}
            code, marker, databases = check_databases_payload(payload)
            if code != 0:
                print(marker)
                return code

            if args.com:
                selected = [args.database] if args.database else databases
                if not selected:
                    print("HEALTH_COM_NO_DATABASES")
                    return 2
                if args.database and args.database not in databases:
                    print(
                        "HEALTH_DATABASE_NOT_FOUND database="
                        + ascii_json(args.database)
                    )
                    return 3
                failures = await _probe_com(client, selected)
                if failures:
                    print(
                        "HEALTH_COM_FAIL databases="
                        + ascii_json(failures)
                    )
                    return 3
                print("HEALTH_COM_OK")
                return 0

            print(f"HEALTH_OK tools={len(names)} databases={len(databases)}")
            return 0
    except Exception as exc:  # noqa: BLE001 — причина ошибки выводится в консоль
        # stdout остаётся ASCII-маркером; полные детали (в т.ч. внутренние
        # ошибки, а не только сетевое подключение) — в stderr
        print("HEALTH_CONNECT_ERROR:", type(exc).__name__)
        traceback.print_exception(exc)
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="healthcheck.py",
        description="MCP health-check моста 1C (5 инструментов + list_databases; "
        "с --com ещё и COM/метаданные каждой базы).",
    )
    parser.add_argument(
        "--com",
        action="store_true",
        help="проверить COM-подключение и метаданные каждой включённой базы",
    )
    parser.add_argument(
        "--database",
        metavar="KEY",
        help="в режиме --com проверить только одну базу",
    )
    args = parser.parse_args()
    if args.database and not args.com:
        parser.error("--database требует --com")
    return asyncio.run(run(args))


if __name__ == "__main__":
    raise SystemExit(main())
