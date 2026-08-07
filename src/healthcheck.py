"""Единый MCP health-check для HTTP-сервера моста.

Одна и та же проверка используется установщиком (install.ps1) и служебными
скриптами (restart_http_server.ps1), чтобы проверки не расходились:

  * сервер отвечает на streamable-http /mcp/<токен>;
  * tools/list возвращает ВСЕ пять инструментов моста;
  * list_databases вызывается без ошибки.

Запуск:   python healthcheck.py     (ONEC_HTTP_TOKEN — в переменной окружения)
Код:      0 = проверка пройдена; 1 = сервер недоступен/инструменты не те;
          2 = list_databases вернул ошибку.
Вывод — ASCII-маркеры (HEALTH_OK / HEALTH_TOOLS_MISSING / ...) для журналов.
"""

import asyncio
import os

EXPECTED_TOOLS = {
    "describe_object",
    "execute_query",
    "get_object_by_ref",
    "list_databases",
    "list_metadata",
}


async def run() -> int:
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
    except Exception as exc:  # noqa: BLE001 — причина ошибки выводится в консоль
        print("HEALTH_CONNECT_ERROR:", type(exc).__name__)
        return 1
    print(f"HEALTH_OK tools={len(names)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(run()))
