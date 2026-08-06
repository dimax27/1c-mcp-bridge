"""
1C MCP Bridge — HTTP server for MCP-клиентов с песочницей (Qwen, ChatGPT/Codex).

Security: binds to 127.0.0.1. Uses a random path prefix if
ONEC_HTTP_TOKEN is set (e.g. /mcp/<token> instead of /mcp).

Журналирование устроено так, чтобы ошибки не терялись в скрытом окне VBS:
bootstrap-логгер настраивается ДО импорта основного сервера, поэтому даже
ModuleNotFoundError / битые зависимости / ошибки databases.json попадают в
--log-file. Uvicorn сам выходит (SystemExit) при занятом порте — перехватываем.
"""

import argparse
import logging
import os
from logging.handlers import RotatingFileHandler

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

# Общий handler: сначала подключается к bootstrap-логгеру, после успешного
# импорта — к основному "mcp-1c". Один экземпляр на файл — без конфликтов ротации.
_file_handler: RotatingFileHandler | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--insecure-host", action="store_true")
    parser.add_argument(
        "--log-file",
        default="",
        help="путь к журналу (ротация 5 МБ x 3); ошибки запуска пишутся сюда",
    )
    return parser.parse_args()


def setup_bootstrap_logging(log_file: str) -> logging.Logger:
    """Журнал доступен ещё до импорта mcp_server_1c."""
    bootstrap = logging.getLogger("mcp-1c.bootstrap")
    global _file_handler
    if not log_file:
        return bootstrap
    try:
        _file_handler = RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
    except OSError as exc:
        bootstrap.error("Не удалось открыть журнал %s: %s", log_file, exc)
        return bootstrap
    _file_handler.setFormatter(logging.Formatter(LOG_FORMAT))
    bootstrap.addHandler(_file_handler)
    bootstrap.setLevel(os.environ.get("ONEC_LOG_LEVEL", "INFO"))
    return bootstrap


def main() -> int:
    args = parse_args()
    bootstrap = setup_bootstrap_logging(args.log_file)

    if args.host not in ("127.0.0.1", "localhost", "::1") and not args.insecure_host:
        bootstrap.error("Binding to %s requires --insecure-host flag.", args.host)
        return 2

    try:
        from mcp_server_1c import DB_CONFIG, log, mcp
    except Exception:  # noqa: BLE001 — ловим любую ошибку импорта (историческая поломка)
        # Историческая поломка проекта: не хватало модулей — ошибка должна
        # быть видна в журнале, а не только в скрытом stderr.
        bootstrap.exception("Не удалось импортировать MCP-сервер (mcp_server_1c)")
        return 1

    if _file_handler is not None:
        log.addHandler(_file_handler)

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
    except SystemExit as exc:
        # uvicorn сам выходит при невозможности привязаться к порту и т.п.
        code = exc.code if isinstance(exc.code, int) else 1
        log.critical("HTTP-сервер завершился при запуске/работе, код=%s", code)
        return code
    except Exception:  # noqa: BLE001 — записываем причину и выходим с кодом ошибки
        log.exception("HTTP-сервер аварийно завершился")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
