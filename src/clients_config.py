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
#   download_url: where to get the app
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
        "config_name": "mcp_config.json",
        "exe_names": [
            "Programs/qwen-desktop/Qwen.exe",
            "Programs/Qwen/Qwen.exe",
            "Programs/qwen/Qwen.exe",
            "QwenDesktop/Qwen.exe",
        ],
        "download_url": "https://www.qianwenai.com/agents/qwen",
    },
    {
        "id": "kimi",
        "name": "Kimi Desktop",
        "appdata_dir": "Kimi",
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
]


def _appdata() -> Path:
    return Path(os.environ.get("APPDATA", os.path.expandvars("%APPDATA%")))


def _localappdata() -> Path:
    return Path(
        os.environ.get("LOCALAPPDATA", os.path.expandvars("%LOCALAPPDATA%"))
    )


def client_config_path(client: dict) -> Path:
    """Return the full path to a client's MCP config file.

    Standard clients use %APPDATA%/appdata_dir/config_name.
    Clients with config_subdir use %APPDATA%/appdata_dir/config_subdir/config_name
    (e.g. Reasonix: %APPDATA%/reasonix/global-workspace/.mcp.json).
    """
    base = _appdata() / client["appdata_dir"]
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
            client_dir = _appdata() / c["appdata_dir"]
            if client_dir.exists():
                installed.append(c)

    return installed


def read_client_config(client: dict) -> dict:
    """Read and parse a client's MCP config file. Returns empty dict on failure."""
    path = client_config_path(client)
    if not path.exists():
        return {}
    import json

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def write_client_config(client: dict, config: dict) -> None:
    """Write a client's MCP config file (UTF-8 without BOM)."""
    import json

    path = client_config_path(client)
    path.parent.mkdir(parents=True, exist_ok=True)
    # UTF-8 without BOM — same as Claude expects
    raw = json.dumps(config, ensure_ascii=False, indent=2)
    path.write_text(raw, encoding="utf-8")


def update_mcp_servers(client: dict, server_name: str, server_entry: dict) -> bool:
    """Add or update an MCP server entry in a client's config.

    Returns True if the config was created or updated.
    """
    config = read_client_config(client)

    if "mcpServers" not in config:
        config["mcpServers"] = {}
    if not isinstance(config["mcpServers"], dict):
        config["mcpServers"] = {}

    existing = config["mcpServers"].get(server_name)
    if existing == server_entry:
        return False  # No change needed

    config["mcpServers"][server_name] = server_entry
    write_client_config(client, config)
    return True


def remove_mcp_server(client: dict, server_name: str) -> bool:
    """Remove an MCP server entry from a client's config.

    Returns True if the entry was found and removed.
    """
    config = read_client_config(client)
    servers = config.get("mcpServers")
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
