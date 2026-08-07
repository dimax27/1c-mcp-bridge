"""HTTP smoke test: real streamable-http server + MCP client through the token path.

Поднимает mcp_server_1c_http.py на свободном порту с временным databases.json
и проверяет полный цикл MCP-транспорта: initialize → tools/list → call_tool,
а также систему журналирования (--log-file), поведение при занятом порте и
read-only аннотации инструментов.

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
import time
import urllib.error
import urllib.request
from pathlib import Path

import pytest

pytest.importorskip("mcp")  # нужен клиент mcp (входит в requirements.txt)

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
_LOG_LINE_START = "Starting 1C MCP Bridge HTTP"
_LOG_LINE_PORT_BUSY = "HTTP-сервер завершился при запуске/работе, код="


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
def start_server(tmp_path):
    """Фабрика запуска сервера: возвращает контекст с url, путём лога и proc."""
    procs: list[subprocess.Popen] = []
    blockers: list[socket.socket] = []

    def _start(
        *,
        log_file: Path | None = None,
        occupy_port: bool = False,
        port: int | None = None,
        databases: dict | None = None,
        extra_env: dict | None = None,
        token: str | None = _TOKEN,
    ) -> dict:
        if databases is None:
            databases = {
                "smoke": {
                    "description": "smoke test (без 1С)",
                    "progid": "V83.COMConnector",
                    "connection_string": 'Srvr="localhost";Ref="smoke"',
                    "dll_path": "",
                }
            }
        default_db = "smoke" if "smoke" in databases else next(iter(databases), "")
        db = tmp_path / f"databases-{len(procs)}.json"
        db.write_text(
            json.dumps(
                {
                    "version": 1,
                    "default_database": default_db,
                    "databases": databases,
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        port = port or _free_port()
        blocker = None
        if occupy_port:
            blocker = socket.socket()
            blocker.bind(("127.0.0.1", port))
            blocker.listen(1)
            blockers.append(blocker)

        env = dict(os.environ)
        env["ONEC_DATABASES_FILE"] = str(db)
        if token is not None:
            env["ONEC_HTTP_TOKEN"] = token
        if extra_env:
            env.update(extra_env)

        cmd = [sys.executable, str(SERVER), "--port", str(port)]
        if log_file is not None:
            cmd += ["--log-file", str(log_file)]

        proc = subprocess.Popen(
            cmd,
            cwd=str(SRC),
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        procs.append(proc)
        if not occupy_port:
            assert _wait_port(port), "HTTP-сервер не поднялся за отведённое время"
        return {
            "proc": proc,
            "port": port,
            "script": str(SERVER),
            "token_url": f"http://127.0.0.1:{port}/mcp/{env.get('ONEC_HTTP_TOKEN', _TOKEN)}",
            "plain_url": f"http://127.0.0.1:{port}/mcp",
            "log_file": log_file,
        }

    yield _start

    for p in procs:
        if p.poll() is None:
            p.terminate()
            try:
                p.wait(timeout=5)
            except subprocess.TimeoutExpired:
                p.kill()
    for b in blockers:
        b.close()


def _probe_tools(token_url: str) -> tuple:
    """list_tools + list_databases с retry на этап старта сервера."""

    async def run() -> tuple:
        from mcp import Client

        last: Exception | None = None
        for _ in range(15):
            try:
                async with Client(token_url) as client:
                    result = await client.list_tools()
                    names = sorted(t.name for t in result.tools)
                    annotations = {
                        t.name: t.annotations for t in result.tools
                    }
                    call = await client.call_tool("list_databases", {})
                    return names, annotations, call
            except Exception as exc:  # noqa: BLE001 — retry на этапе старта
                last = exc
                await asyncio.sleep(1)
        raise AssertionError(f"не удалось подключиться к серверу: {last}")

    return asyncio.run(run())


def _call_tool_text(call) -> str:
    return "".join(
        getattr(item, "text", "") or ""
        for item in getattr(call, "content", []) or []
    )


def test_http_smoke_token_path(start_server):
    ctx = start_server()
    names, annotations, call = _probe_tools(ctx["token_url"])

    assert names == EXPECTED_TOOLS, f"tools/list вернул: {names}"

    # инструменты помечены read-only — клиент видит подсказку
    for name in EXPECTED_TOOLS:
        ann = annotations.get(name)
        assert ann is not None and getattr(ann, "read_only_hint", False) is True, (
            f"инструмент {name} не помечен read_only_hint=True"
        )

    assert not getattr(call, "is_error", False), "list_databases вернул ошибку"
    payload = json.loads(_call_tool_text(call))
    assert "smoke" in payload.get("databases", {}), f"нет базы smoke: {payload}"


def test_token_fallback_from_file(start_server, tmp_path):
    """Если ONEC_HTTP_TOKEN не задан, сервер читает токен из файла."""
    token_file = tmp_path / ".http_token"
    fallback_token = "FALLBACK99"
    token_file.write_text(fallback_token, encoding="utf-8")

    ctx = start_server(
        extra_env={"ONEC_HTTP_TOKEN_FILE": str(token_file)},
        token=None,  # не задаём токен через env — сервер сам прочитает файл
    )
    url = f"http://127.0.0.1:{ctx['port']}/mcp/{fallback_token}"
    names, _, _ = _probe_tools(url)
    assert "list_databases" in names, f"tools/list: {names}"
    # /mcp без токена — всё ещё 404 (защита на месте)
    with pytest.raises(urllib.error.HTTPError) as excinfo:
        urllib.request.urlopen(ctx["plain_url"], timeout=10)
    assert excinfo.value.code == 404


def test_http_smoke_plain_path_rejected(start_server):
    """Без токена маршрут /mcp не должен существовать (защита path-токеном)."""
    ctx = start_server()
    req = urllib.request.Request(
        ctx["plain_url"],
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


def test_http_log_created_and_token_never_in_log(start_server, tmp_path):
    log_file = tmp_path / "http-server.log"
    ctx = start_server(log_file=log_file)

    _probe_tools(ctx["token_url"])

    assert log_file.exists(), "http-server.log не создан"
    text = log_file.read_text(encoding="utf-8", errors="replace")
    assert _LOG_LINE_START in text, f"в логе нет строки запуска: {text[:300]}"
    assert _TOKEN not in text, "path-токен попал в журнал!"


def test_occupied_port_is_logged(start_server, tmp_path):
    log_file = tmp_path / "http-server.log"
    ctx = start_server(log_file=log_file, occupy_port=True)

    try:
        code = ctx["proc"].wait(timeout=30)
    except subprocess.TimeoutExpired:
        ctx["proc"].kill()
        pytest.fail("сервер не завершился при занятом порте")

    assert code != 0, f"ожидался ненулевой код возврата, получен {code}"
    assert log_file.exists(), "http-server.log не создан"
    text = log_file.read_text(encoding="utf-8", errors="replace")
    assert _LOG_LINE_PORT_BUSY in text, (
        f"в логе нет записи об ошибке старта: {text[-500:]}"
    )


def _decode_process_output(data: bytes | None) -> str:
    """Текст только для диагностики; ASCII-маркеры проверяются по байтам."""
    return (data or b"").decode("utf-8", errors="replace")


@pytest.mark.skipif(sys.platform != "win32", reason="нужен PowerShell/Windows")
def test_shared_healthcheck_script(start_server):
    """src/healthcheck.py — единая проверка (5 инструментов + list_databases):
    успех с верным токеном, отказ с неверным."""
    ctx = start_server()
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    ok = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py")],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    assert ok.returncode == 0, f"healthcheck упал: {ok.stdout[-400:]}\n{ok.stderr[-400:]}"
    assert "HEALTH_OK" in ok.stdout, ok.stdout[-400:]

    env["ONEC_HTTP_TOKEN"] = "WRONGTOKEN"
    bad = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py")],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    assert bad.returncode != 0, "healthcheck с неверным токеном не должен проходить"


def test_shared_healthcheck_com_probe_fails_on_unreachable_db(start_server):
    """healthcheck.py --com: для недоступной базы (smoke без 1С) COM-проба
    честно падает с кодом 3, вывод — ASCII."""
    ctx = start_server()
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py"), "--com"],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
        env=env,
    )
    # сервер MCP отвечает, но база 'smoke' не подключается к 1С -> код 3
    assert res.returncode == 3, f"ожидался код 3: rc={res.returncode}\n{res.stdout[-600:]}"
    assert 'HEALTH_COM database="smoke" status=FAIL' in res.stdout, res.stdout[-600:]
    assert 'HEALTH_COM_FAIL databases=["smoke"]' in res.stdout, res.stdout[-600:]
    # краткая причина COM-ошибки — в stderr (санитизированная)
    assert "HEALTH_COM_DETAIL" in res.stderr, res.stderr[-600:]
    # вывод — только ASCII (имена/ошибки экранируются)
    res.stdout.encode("ascii")  # не должно упасть с UnicodeEncodeError


def test_healthcheck_redacts_token_from_traceback(start_server):
    """Traceback при HTTP-ошибке (404) не должен содержать токен."""
    ctx = start_server()
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = "WRONGTOKEN"  # 404 -> traceback с URL и токеном
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py")],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    combined = res.stdout + res.stderr
    assert "WRONGTOKEN" not in combined, "токен утёк в вывод healthcheck"
    assert res.returncode != 0


def test_healthcheck_cyrillic_db_key_ascii(start_server):
    """Ключ базы с кириллицей экранируется ensure_ascii: вывод остаётся ASCII."""
    ctx = start_server(
        databases={
            "Бухгалтерия": {
                "description": "БП (без 1С)",
                "progid": "V83.COMConnector",
                "connection_string": 'Srvr="localhost";Ref="buh"',
                "dll_path": "",
            }
        }
    )
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py"), "--com"],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
        env=env,
    )
    assert res.returncode == 3, f"rc={res.returncode}\n{res.stdout[-600:]}"
    # в stdout — escape-последовательность, а не «сырая» кириллица
    assert "Бухгалтерия" not in res.stdout
    assert r"\u0411\u0443\u0445\u0433\u0430\u043b\u0442\u0435\u0440\u0438\u044f" in res.stdout
    res.stdout.encode("ascii")


def test_healthcheck_empty_databases_fails(start_server):
    """Пустой список баз (сервер поднялся, но баз нет) — это ошибка (код 2)."""
    ctx = start_server(databases={})
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py")],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    assert res.returncode == 2, f"rc={res.returncode}\n{res.stdout[-600:]}"
    assert "HEALTH_NO_DATABASES" in res.stdout, res.stdout[-600:]


def test_healthcheck_database_requires_com(start_server):
    """--database без --com отклоняется парсером (код 2)."""
    ctx = start_server()
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py"), "--database", "smoke"],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    assert res.returncode == 2, f"rc={res.returncode}\n{res.stdout[-300:]}\n{res.stderr[-300:]}"
    assert "--database" in res.stderr


def test_healthcheck_unknown_database_fails(start_server):
    """--com --database <несуществующий> -> HEALTH_DATABASE_NOT_FOUND, код 3."""
    ctx = start_server()
    env = dict(os.environ)
    env["ONEC_HTTP_TOKEN"] = _TOKEN
    env["ONEC_HTTP_PORT"] = str(ctx["port"])

    res = subprocess.run(
        [sys.executable, str(REPO / "src" / "healthcheck.py"), "--com", "--database", "NOPE"],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
        env=env,
    )
    assert res.returncode == 3, f"rc={res.returncode}\n{res.stdout[-600:]}"
    assert "HEALTH_DATABASE_NOT_FOUND" in res.stdout, res.stdout[-600:]


def test_installer_stops_only_target_server(start_server, tmp_path):
    """stop_http_server.ps1 с ExpectedScriptPath останавливает ТОЛЬКО процесс
    этой установки.

    Установщик вызывает скрипт в самом начале, чтобы старый сервер не держал
    порт и файлы venv во время переустановки. При этом:
      - decoy с тем же именем скрипта из другого каталога — выживает;
      - посторонний процесс на другом порту — выживает;
      - целевой сервер этой установки — останавливается.
    """
    # свободные порты вместо хардкода: на CI-раннерах 18731/18732 могут быть заняты
    target_port = _free_port()
    other_port = _free_port()
    ctx = start_server(port=target_port)
    assert ctx["proc"].poll() is None, "целевой сервер должен быть запущен"

    # decoy: «тот же скрипт» из другого каталога — просто спящий процесс
    decoy_dir = tmp_path / "other-install"
    decoy_dir.mkdir()
    decoy_script = decoy_dir / "mcp_server_1c_http.py"
    decoy_script.write_text("import time; time.sleep(120)\n", encoding="utf-8")
    decoy = subprocess.Popen([sys.executable, str(decoy_script)], cwd=str(decoy_dir))

    # второй decoy: имя БЕЗ _http — проверяем, что улучшенный regex ловит оба варианта
    decoy_no_http = decoy_dir / "mcp_server_1c.py"
    decoy_no_http.write_text("import time; time.sleep(120)\n", encoding="utf-8")
    decoy2 = subprocess.Popen(
        [sys.executable, str(decoy_no_http)], cwd=str(decoy_dir)
    )

    # посторонний процесс на другом порту (не мост)
    foreign = subprocess.Popen(
        [sys.executable, "-m", "http.server", str(other_port), "--bind", "127.0.0.1"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    try:
        stop_script = REPO / "installer" / "stop_http_server.ps1"
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(stop_script),
                "-Port",
                str(target_port),
                "-ExpectedScriptPath",
                ctx["script"],
                "-ExpectedPythonPath",
                sys.executable,
            ],
            capture_output=True,
            timeout=90,
            check=False,
        )
        # PowerShell может писать в UTF-8 при cp1252-локали раннера:
        # берём байты, маркеры проверяем по ASCII, декодируем только для сообщений
        stdout = result.stdout or b""
        stderr = result.stderr or b""
        stdout_text = _decode_process_output(stdout)
        stderr_text = _decode_process_output(stderr)

        assert result.returncode == 0, (
            f"stop_http_server.ps1 упал: rc={result.returncode}\n"
            f"stdout={stdout_text[-600:]}\nstderr={stderr_text[-600:]}"
        )
        marker = f"PORT_{target_port}_FREE".encode("ascii")
        assert marker in stdout, stdout_text[-600:]

        # целевой процесс завершён
        try:
            code = ctx["proc"].wait(timeout=15)
        except subprocess.TimeoutExpired:
            ctx["proc"].kill()
            pytest.fail("stop_http_server.ps1 не остановил целевой сервер")
        assert code is not None

        # decoy и посторонний процесс выжили
        assert decoy.poll() is None, (
            "decoy-процесс (тот же скрипт из другого каталога) не должен останавливаться"
        )
        assert decoy2.poll() is None, (
            "decoy2 (mcp_server_1c.py из другого каталога) не должен останавливаться"
        )
        assert foreign.poll() is None, (
            "посторонний процесс на другом порту не должен останавливаться"
        )
    finally:
        for p in (decoy, decoy2, foreign):
            p.kill()
            try:
                p.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass

    # порт должен освободиться (соединение отклоняется)
    with socket.socket() as s:
        s.settimeout(2)
        try:
            s.connect(("127.0.0.1", ctx["port"]))
            pytest.fail("порт всё ещё принимает соединения")
        except OSError:
            pass  # ожидаемо: порт свободен
