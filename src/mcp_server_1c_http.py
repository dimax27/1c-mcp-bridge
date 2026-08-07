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
    """Журнал доступен ещё до импорта mcp_server_1c.

    Файловый handler подключаем сразу к обоим логгерам — и bootstrap, и
    основному "mcp-1c": ошибки чтения databases.json, перехваченные внутри
    модуля при импорте, попадают в файл, а не только в скрытый stderr.
    """
    bootstrap = logging.getLogger("mcp-1c.bootstrap")
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
    level = os.environ.get("ONEC_LOG_LEVEL", "INFO")
    # Файловый handler вешаем ТОЛЬКО на "mcp-1c": "mcp-1c.bootstrap" — его
    # дочерний логгер (propagate=True), поэтому сообщения bootstrap попадают
    # в тот же файл, но без дублирования.
    logger = logging.getLogger("mcp-1c")
    if not any(
        isinstance(h, RotatingFileHandler) and h.baseFilename == _file_handler.baseFilename
        for h in logger.handlers
    ):
        logger.addHandler(_file_handler)
    logger.setLevel(level)
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

    # Файловый handler уже подключён к "mcp-1c" в setup_bootstrap_logging.
    log.setLevel(os.environ.get("ONEC_LOG_LEVEL", "INFO"))

    # Random path prefix if token is set (defense-in-depth for localhost).
    # Сам токен в лог не выводим — он является частью URL MCP.
    token = os.environ.get("ONEC_HTTP_TOKEN", "").strip()
    if not token:
        # VBS-лаунчер и Планировщик задают переменную; fallback для прямого
        # запуска (python mcp_server_1c_http.py) — читаем из стандартного
        # файла токена (путь можно переопределить через ONEC_HTTP_TOKEN_FILE).
        token_file = os.environ.get(
            "ONEC_HTTP_TOKEN_FILE",
            os.path.join(
                os.environ.get("ProgramData", "C:\\ProgramData"),
                "1cMcpBridge",
                ".http_token",
            ),
        )
        try:
            with open(token_file, encoding="utf-8") as fh:
                token = fh.read().strip()
        except OSError:
            pass
    if not token:
        log.warning(
            "ONEC_HTTP_TOKEN не задан и файл токена не найден — "
            "сервер запущен БЕЗ секретного пути (/mcp отвечает 404)!"
        )
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
