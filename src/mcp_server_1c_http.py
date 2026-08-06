"""
1C MCP Bridge — HTTP server for MCP-клиентов с песочницей (Qwen, ChatGPT/Codex).

Security: binds to 127.0.0.1. Uses a random path prefix if
ONEC_HTTP_TOKEN is set (e.g. /mcp/<token> instead of /mcp).
"""
import logging
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from mcp_server_1c import DB_CONFIG, log, mcp

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--insecure-host", action="store_true")
    parser.add_argument(
        "--log-file",
        default="",
        help="путь к файлу журнала (лог mcp-1c пишется туда, помимо stderr)",
    )
    args = parser.parse_args()

    if args.host not in ("127.0.0.1", "localhost", "::1") and not args.insecure_host:
        parser.error(f"Binding to {args.host} requires --insecure-host flag.")

    if args.log_file:
        handler = logging.FileHandler(args.log_file, encoding="utf-8")
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )
        log.addHandler(handler)

    # Random path prefix if token is set (defense-in-depth for localhost).
    # Сам токен в лог не выводим — он является частью URL MCP.
    token = os.environ.get("ONEC_HTTP_TOKEN", "").strip()
    path_prefix = f"/mcp/{token}" if token else "/mcp"

    if token:
        log.info(
            "Starting 1C MCP Bridge HTTP on %s:%d (secret path /mcp/<token>)",
            args.host,
            args.port,
        )
    else:
        log.info("Starting 1C MCP Bridge HTTP on %s:%d/mcp", args.host, args.port)
    log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))
    try:
        mcp.run(
            transport="streamable-http",
            host=args.host,
            port=args.port,
            streamable_http_path=path_prefix,
        )
    except OSError as exc:
        # Например, порт уже занят — раньше ошибка была невидима (VBS без окна).
        log.critical("Не удалось запустить HTTP-сервер: %s", exc)
        sys.exit(1)
