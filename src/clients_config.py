"""
MCP Desktop Clients Configuration.

Defines supported MCP desktop clients, their config file paths, and
auto-detection logic. Used by the installer, uninstaller, manager, and
diagnostic tools to support Claude Desktop, Qwen Desktop, and Kimi Desktop.

Adding a new client:
    Append a dict to KNOWN_CLIENTS below. The installer and Manager will
    pick it up automatically.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Known MCP desktop clients
# ---------------------------------------------------------------------------
# Each entry:
#   id          : stable internal id, used in env vars (e.g. ONEC_QWEN_CONFIG)
#   name        : human-readable name
#   appdata_dir : folder name under %APPDATA% where the client keeps its files
#   config_name : filename of the MCP servers config JSON
#   exe_names   : candidate exe names for detection under %LOCALAPPDATA%
#   mcp_key     : JSON key for MCP servers (default: "mcpServers"; Qwen uses "mcp_config")
# ---------------------------------------------------------------------------

KNOWN_CLIENTS: list[dict[str, Any]] = [
    {
        "id": "claude",
        "name": "Claude Desktop",
        "appdata_dir": "Claude",
        "config_name": "claude_desktop_config.json",
        "exe_names": [
            "AnthropicClaude/Claude.exe",
            "Programs/claude-desktop/Claude.exe",
            "Programs/Claude/Claude.exe",
        ],
        "download_url": "https://claude.ai/download",
    },
    {
        "id": "qwen",
        "name": "Qwen Desktop",
        "appdata_dir": "Qwen",
        "config_name": "settings.json",
        "mcp_key": "mcp_config",  # Qwen uses "mcp_config" not "mcpServers"
        "exe_names": [
            "Programs/Qwen/Qwen.exe",
            "Programs/qwen-desktop/Qwen.exe",
            "Programs/qwen/Qwen.exe",
            "QwenDesktop/Qwen.exe",
        ],
        "download_url": "https://www.qianwenai.com/agents/qwen",
    },
    {
        "id": "kimi",
        "name": "Kimi Desktop",
        "appdata_dir": "kimi-desktop",
        "config_name": "mcp_config.json",
        "exe_names": [
            "Programs/kimi-desktop/Kimi.exe",
            "Programs/Kimi/Kimi.exe",
            "Programs/kimi/Kimi.exe",
            "KimiDesktop/Kimi.exe",
        ],
        "download_url": "https://kimi.moonshot.cn",
    },
    {
        "id": "reasonix",
        "name": "Reasonix",
        # Uses .mcp.json in workspace dirs, not a single appdata config.
        # Primary target: global-workspace under %APPDATA%\reasonix.
        "appdata_dir": "reasonix",
        "config_name": ".mcp.json",
        "config_subdir": "global-workspace",  # config is at appdata_dir/config_subdir/config_name
        "exe_names": [],
        "download_url": "https://reasonix.ai",
    },
    {
        "id": "chatgpt",
        "name": "ChatGPT Desktop",
        # MSIX-приложение OpenAI Codex хранит конфиг в %USERPROFILE%\.codex\config.toml.
        # Оно запускает stdio-MCP-серверы внутри Windows-песочницы, где COM-коннектор
        # 1С и сеть до сервера 1С недоступны, поэтому ChatGPT настраивается на
        # streamable-http сервер моста на 127.0.0.1:8000 (процесс вне песочницы).
        "appdata_dir": ".codex",
        "config_name": "config.toml",
        "config_format": "toml",      # не JSON: правится точечно, сохраняя остальные секции
        "config_base_env": "USERPROFILE",  # путь = %USERPROFILE%\.codex\config.toml
        "exe_names": [],
        "download_url": "https://chatgpt.com/download",
    },
]


def _appdata() -> Path:
    return Path(os.environ.get("APPDATA", os.path.expandvars("%APPDATA%")))


def _localappdata() -> Path:
    return Path(
        os.environ.get("LOCALAPPDATA", os.path.expandvars("%LOCALAPPDATA%"))
    )


def _client_base_dir(client: dict) -> Path:
    """Base directory for a client's config.

    Standard clients use %APPDATA%/appdata_dir. Clients with
    config_base_env (e.g. ChatGPT -> <USERPROFILE>/.codex) resolve the
    env var and append appdata_dir to it.
    """
    base_env = client.get("config_base_env")
    if base_env:
        root = Path(os.environ.get(base_env, ""))
        return root / client["appdata_dir"]
    return _appdata() / client["appdata_dir"]


def client_config_path(client: dict) -> Path:
    """Return the full path to a client's MCP config file.

    Standard clients use %APPDATA%/appdata_dir/config_name.
    Clients with config_subdir use %APPDATA%/appdata_dir/config_subdir/config_name
    (e.g. Reasonix: %APPDATA%/reasonix/global-workspace/.mcp.json).
    ChatGPT: %USERPROFILE%/.codex/config.toml.
    """
    base = _client_base_dir(client)
    subdir = client.get("config_subdir")
    if subdir:
        return base / subdir / client["config_name"]
    return base / client["config_name"]


def detect_installed_clients() -> list[dict]:
    """Return list of client dicts that appear to be installed.

    Detection strategy (tried in order):
    1. Config file exists.
    2. Executable found under %LOCALAPPDATA%.
    3. Client directory exists under %APPDATA%.
    """
    installed: list[dict] = []
    for c in KNOWN_CLIENTS:
        # Allow per-client override of the config path
        env_var = f"ONEC_{c['id'].upper()}_CONFIG"
        env_path = os.environ.get(env_var, "").strip()
        if env_path:
            if Path(env_path).exists():
                installed.append(c)
                continue

        # Check if the config file already exists
        config_path = client_config_path(c)
        if config_path.exists():
            installed.append(c)
            continue

        # Check for executables (desktop clients only)
        local = _localappdata()
        for exe in c.get("exe_names", []):
            if (local / exe).exists():
                installed.append(c)
                break
        else:
            # Check if the client dir exists at all
            client_dir = _client_base_dir(c)
            if client_dir.exists():
                installed.append(c)

    return installed


def read_client_config(client: dict) -> dict:
    """Read and parse a client's MCP config file.

    Returns empty dict if file doesn't exist.
    Raises ValueError if file exists but can't be parsed (corrupt JSON).
    """
    path = client_config_path(client)
    if not path.exists():
        return {}
    import json

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        raise ValueError(f"Failed to parse {path}: {e}") from e


def write_client_config(client: dict, config: dict) -> None:
    """Write a client's MCP config file atomically (UTF-8 without BOM).

    Copies existing config to .bak, writes to unique temp file,
    then atomically replaces the original via os.replace().
    """
    import json
    import os
    import shutil
    import tempfile

    path = client_config_path(client)
    path.parent.mkdir(parents=True, exist_ok=True)

    raw = json.dumps(config, ensure_ascii=False, indent=2)
    backup = path.with_suffix(path.suffix + ".bak")

    # Backup existing config first (so we never lose it)
    if path.exists():
        shutil.copy2(path, backup)

    # Write to unique temp file in the same directory
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=str(path.parent),
    )
    tmp_path = Path(tmp_name)

    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(raw)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(tmp_path, path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise


def _mcp_key(client: dict) -> str:
    """Return the JSON key for MCP servers in this client's config."""
    return client.get("mcp_key", "mcpServers")


def update_mcp_servers(client: dict, server_name: str, server_entry: dict) -> bool:
    """Add or update an MCP server entry in a client's config.

    Returns True if the config was created or updated.
    """
    config = read_client_config(client)
    key = _mcp_key(client)

    if key not in config:
        config[key] = {}
    if not isinstance(config[key], dict):
        config[key] = {}

    existing = config[key].get(server_name)
    if existing == server_entry:
        return False  # No change needed

    config[key][server_name] = server_entry
    write_client_config(client, config)
    return True


def remove_mcp_server(client: dict, server_name: str) -> bool:
    """Remove an MCP server entry from a client's config.

    Returns True if the entry was found and removed.
    """
    config = read_client_config(client)
    key = _mcp_key(client)
    servers = config.get(key)
    if not isinstance(servers, dict):
        return False
    if server_name not in servers:
        return False
    del servers[server_name]
    write_client_config(client, config)
    return True


def get_clients_overview() -> str:
    """Return a human-readable overview of all clients and their status."""
    installed = {c["id"] for c in detect_installed_clients()}
    lines = ["Поддерживаемые MCP-клиенты для рабочего стола:"]
    for c in KNOWN_CLIENTS:
        status = "✓ найден" if c["id"] in installed else "✗ не найден"
        lines.append(f"  {c['name']}: {status}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Codex (ChatGPT Desktop) config.toml patching
# ---------------------------------------------------------------------------

def patch_codex_config_toml(config_path: Path, server_name: str, url: str) -> bool:
    """Set `[mcp_servers.<server_name>]` to `{ enabled = true, url = "..." }`.

    ChatGPT Desktop (OpenAI Codex, MSIX) хранит MCP-серверы в
    config.toml в домашней папке пользователя (по умолчанию
    <USERPROFILE>/.codex/config.toml). Файл содержит много чужих секций
    (plugins, marketplaces, auth и т.п.), поэтому правка делается точечно
    по тексту, остальные секции и комментарии сохраняются.

    Старые ключи stdio-конфигурации (command/args/env) и подсекция
    `[mcp_servers.<server_name>.env]` удаляются — вместо них подставляется
    url streamable-http сервера (процесс моста вне песочницы MSIX).

    Returns True if the file was changed, False if already up to date.
    If url is None, the whole section (including sub-tables) is removed.
    """
    import json
    import os
    import re
    import shutil
    import tempfile

    import tomllib

    if not re.fullmatch(r"[\w.\-]+", server_name):
        raise ValueError(f"Недопустимое имя MCP-сервера: {server_name!r}")

    section = f"[mcp_servers.{server_name}]"
    sub_prefix = f"[mcp_servers.{server_name}."
    # json.dumps даёт корректную TOML basic-string (экранирует " и \)
    url_line = f'url = {json.dumps(url)}'

    if config_path.exists():
        # newline="" — иначе Python-чтение нормализует CRLF в LF,
        # и мы не сможем сохранить стиль переводов строк пользователя
        with open(config_path, encoding="utf-8", newline="") as f:
            original = f.read()
    else:
        original = ""

    if original.strip():
        try:
            tomllib.loads(original)
        except tomllib.TOMLDecodeError as e:
            raise ValueError(f"config.toml не является валидным TOML: {e}") from e

    newline = "\r\n" if "\r\n" in original else "\n"
    lines = original.splitlines()
    found_section = False
    out: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not found_section and stripped == section:
            found_section = True
            # пропускаем тело секции
            i += 1
            while i < len(lines):
                if lines[i].strip().startswith("["):
                    break
                i += 1
            # пропускаем подсекции [mcp_servers.<name>.*] (например .env)
            while i < len(lines) and lines[i].strip().startswith(sub_prefix):
                i += 1
                while i < len(lines) and not lines[i].strip().startswith("["):
                    i += 1
            if url is not None:
                out.append(section)
                out.append("enabled = true")
                out.append(url_line)
            out.append("")
            continue
        out.append(line)
        i += 1

    if not found_section:
        if url is None:
            return False  # удалять нечего
        if out and out[-1].strip() != "":
            out.append("")
        out.append(section)
        out.append("enabled = true")
        out.append(url_line)
        out.append("")

    result = newline.join(out).rstrip() + newline
    # контроль: результат должен оставаться валидным TOML
    try:
        tomllib.loads(result)
    except tomllib.TOMLDecodeError as e:
        raise ValueError(f"Результат патча невалиден (url содержит недопустимые символы?): {e}") from e

    if original == result:
        return False  # уже настроено

    config_path.parent.mkdir(parents=True, exist_ok=True)
    if config_path.exists():
        shutil.copy2(config_path, config_path.with_suffix(config_path.suffix + ".bak"))

    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{config_path.name}.", suffix=".tmp", dir=str(config_path.parent)
    )
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(result)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(tmp_path, config_path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
    return True


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        prog="clients_config.py",
        description="Настройка MCP-клиентов моста 1C.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_patch = sub.add_parser(
        "patch-codex",
        help="Задать [mcp_servers.<name>] url в config.toml ChatGPT Desktop",
    )
    p_patch.add_argument("--path", required=True, help="путь к config.toml")
    p_patch.add_argument("--server", default="1c-bridge", help="имя MCP-сервера")
    p_patch.add_argument("--url", required=True, help="URL streamable-http сервера")

    p_remove = sub.add_parser(
        "remove-codex",
        help="Удалить [mcp_servers.<name>] из config.toml ChatGPT Desktop",
    )
    p_remove.add_argument("--path", required=True, help="путь к config.toml")
    p_remove.add_argument("--server", default="1c-bridge", help="имя MCP-сервера")

    args = parser.parse_args()
    if args.cmd == "patch-codex":
        patch_codex_config_toml(Path(args.path), args.server, args.url)
        print(f"OK: [mcp_servers.{args.server}] url={args.url} -> {args.path}")
    elif args.cmd == "remove-codex":
        changed = patch_codex_config_toml(Path(args.path), args.server, url=None)
        print(f"{'OK' if changed else 'NOT FOUND'}: [mcp_servers.{args.server}] удалён -> {args.path}")
    else:  # pragma: no cover
        parser.error(f"unknown command: {args.cmd}")
