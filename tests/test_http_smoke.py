"""HTTP smoke test: real streamable-http server + MCP client through the token path.

Поднимает mcp_server_1c_http.py на свободном порту с временным databases.json
и проверяет полный цикл MCP-транспорта: initialize → tools/list → call_tool.

Не требует 1С: tools/list и list_databases работают без COM-подключения.

Требует пакет `mcp` (входит в requirements.txt, ставится в CI). Локально можно
запустить интерпретатором из .venv моста: `python -m pytest tests/test_http_smoke.py -q`.
"""

import asyncio
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

import pytest

mcp = pytest.importorskip("mcp")  # нужен клиент mcp (входит в requirements.txt)

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "src"
SERVER = SRC / "mcp_server_1c_http.py"
EXPECTED_TOOLS = [
    "describe_object",
    "execute_query",
    "get_object_by_ref",
    "list_databases",
    "list_metadata",
]

_TOKEN = "SmokeToken123"


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _wait_port(port: int, timeout: float = 20.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        with socket.socket() as s:
            s.settimeout(1.0)
            try:
                s.connect(("127.0.0.1", port))
                return True
            except OSError:
                time.sleep(0.3)
    return False


@pytest.fixture
def http_server():
    with tempfile.TemporaryDirectory(prefix="1c-bridge-smoke-") as td:
        db = Path(td) / "databases.json"
        db.write_text(
            json.dumps(
                {
                    "version": 1,
                    "default_database": "smoke",
                    "databases": {
                        "smoke": {
                            "description": "smoke test (без 1С)",
                            "progid": "V83.COMConnector",
                            "connection_string": 'Srvr="localhost";Ref="smoke"',
                            "dll_path": "",
                        }
                    },
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        port = _free_port()
        env = dict(os.environ)
        env["ONEC_DATABASES_FILE"] = str(db)
        env["ONEC_HTTP_TOKEN"] = _TOKEN

        proc = subprocess.Popen(
            [sys.executable, str(SERVER), "--port", str(port)],
            cwd=str(SRC),
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        try:
            assert _wait_port(port), "HTTP-сервер не поднялся за отведённое время"
            token_url = f"http://127.0.0.1:{port}/mcp/{_TOKEN}"
            plain_url = f"http://127.0.0.1:{port}/mcp"
            yield token_url, plain_url
        finally:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()


def _run_probe(token_url: str):
    """С retry: первый connect может прийти до полного старта uvicorn."""

    async def probe() -> tuple:
        from mcp import Client

        last: Exception | None = None
        for _ in range(15):
            try:
                async with Client(token_url) as client:
                    result = await client.list_tools()
                    names = sorted(t.name for t in result.tools)
                    call = await client.call_tool("list_databases", {})
                    return names, call
            except Exception as exc:  # noqa: BLE001 — retry на этапе старта
                last = exc
                await asyncio.sleep(1)
        raise AssertionError(f"не удалось подключиться к серверу: {last}")

    return asyncio.run(probe())


def test_http_smoke_token_path(http_server):
    token_url, _plain = http_server
    names, call = _run_probe(token_url)

    assert names == EXPECTED_TOOLS, f"tools/list вернул: {names}"

    assert not getattr(call, "is_error", False), "list_databases вернул ошибку"
    text = "".join(
        getattr(item, "text", "") or ""
        for item in getattr(call, "content", []) or []
    )
    payload = json.loads(text)
    assert "smoke" in payload.get("databases", {}), f"нет базы smoke: {payload}"


def test_http_smoke_plain_path_rejected(http_server):
    """Без токена маршрут /mcp не должен существовать (защита path-токеном)."""
    _token_url, plain_url = http_server
    req = urllib.request.Request(
        plain_url,
        data=json.dumps(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": {},
                    "clientInfo": {"name": "smoke", "version": "1"},
                },
            }
        ).encode(),
        headers={"Content-Type": "application/json", "MCP-Protocol-Version": "2025-06-18"},
        method="POST",
    )
    with pytest.raises(urllib.error.HTTPError) as excinfo:
        urllib.request.urlopen(req, timeout=10)
    assert excinfo.value.code == 404
