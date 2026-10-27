"""Tests for the Inno Setup installer packaging.

Pure Python, no COM/1C required.

Regression guard for: v0.5.0 shipped `mcp_server_1c.py` importing
`config`, `credentials`, `query_timeout` — but `installer/setup.iss`
[Files] section didn't package those modules, so the installed server
crashed on startup with `ModuleNotFoundError`.

Rule: every module in `src/` must be installed to `{app}` by setup.iss.
"""

import re
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
ISS = REPO / "installer" / "setup.iss"
SRC = REPO / "src"

_SRC_ENTRY = re.compile(
    r'Source:\s*"\.\.\\src\\([\w.]+\.py)"\s*;\s*DestDir:\s*"\{app\}"'
)


def _read_iss() -> str:
    raw = ISS.read_bytes()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _packaged_src_modules() -> set[str]:
    return set(_SRC_ENTRY.findall(_read_iss()))


def test_every_src_module_is_packaged():
    """Каждый src/*.py обязан попадать в {app} при установке."""
    on_disk = {p.name for p in SRC.glob("*.py")}
    missing = on_disk - _packaged_src_modules()
    assert not missing, (
        "Эти модули есть в src/, но не включены в [Files] setup.iss: "
        f"{sorted(missing)}. Без них установленный сервер падает с "
        "ModuleNotFoundError."
    )


def test_packaged_modules_exist_on_disk():
    """Обратная проверка: записи в setup.iss не должны быть битыми."""
    on_disk = {p.name for p in SRC.glob("*.py")}
    stale = _packaged_src_modules() - on_disk
    assert not stale, f"В [Files] setup.iss указаны отсутствующие файлы: {sorted(stale)}"


def test_critical_modules_present():
    """Явная проверка модулей, импортируемых mcp_server_1c.py при старте."""
    packaged = _packaged_src_modules()
    for name in ("config.py", "credentials.py", "query_timeout.py",
                 "mcp_server_1c.py", "mcp_server_1c_http.py", "clients_config.py"):
        assert name in packaged, f"Модуль {name} не попал в установщик"


INSTALL_PS1 = REPO / "installer" / "install.ps1"

# У Tee-Object параметр -Encoding есть только в PowerShell 7. Установщик
# (и ярлыки) запускают Windows PowerShell 5.1, где такого параметра нет.
_TEE_ENCODING = re.compile(r"Tee-Object[^\r\n]*-Encoding", re.IGNORECASE)


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def test_install_ps1_uses_only_powershell51_parameters():
    """install.ps1 не должен вызывать `Tee-Object -Encoding` (регресс v0.6.1).

    Установщик запускает install.ps1 скрыто в Windows PowerShell 5.1.
    Первый же вызов Log падает с ParameterBindingException, пост-установочный
    этап молча не выполняется: нет .venv, databases.json, токена и задачи
    Планировщика — мост не запускается, а в логах пусто.
    """
    text = _read_text(INSTALL_PS1)
    offenders = _TEE_ENCODING.findall(text)
    assert not offenders, (
        "install.ps1 использует `Tee-Object -Encoding` (есть только в PowerShell 7), "
        f"из-за чего установка падает на PowerShell 5.1: {offenders}. "
        "Пишите лог через `Write-Host` + `Add-Content -Encoding UTF8`."
    )

