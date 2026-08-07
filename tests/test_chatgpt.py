"""Tests for ChatGPT Desktop (Codex) MCP configuration.

Pure Python, no COM/1C required.

Covers: KNOWN_CLIENTS entry for chatgpt, config path resolution under
USERPROFILE/.codex/config.toml, and the TOML patcher used by the
installer (patch_codex_config_toml) that switches the 1c-bridge entry
from stdio to the streamable-http URL (the MSIX sandbox breaks stdio
COM servers).
"""

import sys
import tempfile
import tomllib
from pathlib import Path

import pytest

SRC = str(Path(__file__).resolve().parent.parent / "src")
sys.path.insert(0, SRC)

from clients_config import (
    KNOWN_CLIENTS,
    client_config_path,
    patch_codex_config_toml,
)


def _chatgpt() -> dict:
    for c in KNOWN_CLIENTS:
        if c["id"] == "chatgpt":
            return c
    raise AssertionError("chatgpt not in KNOWN_CLIENTS")


class TestChatGptClientEntry:
    def test_chatgpt_in_known_clients(self):
        c = _chatgpt()
        assert c["name"] == "ChatGPT Desktop"
        assert c["config_format"] == "toml"
        assert c["config_name"] == "config.toml"
        assert c["config_base_env"] == "USERPROFILE"

    def test_config_path_under_userprofile(self, monkeypatch):
        monkeypatch.setenv("USERPROFILE", r"C:\Users\Test")
        p = client_config_path(_chatgpt())
        assert p == Path(r"C:\Users\Test\.codex\config.toml")


class TestPatchCodexConfigToml:
    def _tmp(self, text: str) -> Path:
        d = tempfile.mkdtemp(prefix="codex-test-")
        p = Path(d) / "config.toml"
        # newline="" — писать как есть, без трансляции переводов строк Windows
        with open(p, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        return p

    STDIO_CFG = r"""# comment at top
[marketplaces.openai-bundled]
last_updated = "2026-08-06T11:41:36Z"

[mcp_servers.node_repl]
args = []
command = 'C:\node.exe'

[mcp_servers.1c-bridge]
enabled = true
command = "C:\\Program Files\\1cMcpBridge\\.venv\\Scripts\\python.exe"
args = [ "C:\\Program Files\\1cMcpBridge\\mcp_server_1c.py" ]

[mcp_servers.1c-bridge.env]
ONEC_DATABASES_FILE = "C:\\ProgramData\\1cMcpBridge\\databases.json"

[mcp_servers.bookstack]
enabled = true
command = "npx"

[plugins."browser@openai-bundled"]
enabled = true

[desktop]
fontSize = 14
"""

    def test_replaces_stdio_with_url_and_keeps_other_sections(self):
        p = self._tmp(self.STDIO_CFG)
        changed = patch_codex_config_toml(p, "1c-bridge",
                                          "http://127.0.0.1:8000/mcp/TOKEN")
        assert changed is True

        data = tomllib.loads(p.read_text(encoding="utf-8"))
        entry = data["mcp_servers"]["1c-bridge"]
        assert entry == {"enabled": True,
                         "url": "http://127.0.0.1:8000/mcp/TOKEN"}
        # старая stdio-подсекция env удалена
        assert "env" not in data["mcp_servers"]["1c-bridge"]
        # соседние серверы и прочие секции целы
        assert data["mcp_servers"]["node_repl"]["command"] == "C:\\node.exe"
        assert data["mcp_servers"]["bookstack"]["command"] == "npx"
        assert data["plugins"]["browser@openai-bundled"]["enabled"] is True
        assert data["desktop"]["fontSize"] == 14

    def test_idempotent(self):
        p = self._tmp(self.STDIO_CFG)
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        again = patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        assert again is False

    def test_appends_section_if_missing(self):
        p = self._tmp('[desktop]\nfontSize = 14\n')
        changed = patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        assert changed is True
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        assert data["mcp_servers"]["1c-bridge"]["url"] == "http://127.0.0.1:8000/mcp/TOKEN"
        assert data["desktop"]["fontSize"] == 14

    def test_creates_new_file(self):
        p = Path(tempfile.mkdtemp(prefix="codex-test-")) / "config.toml"
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        assert data["mcp_servers"]["1c-bridge"]["url"] == "http://127.0.0.1:8000/mcp/TOKEN"

    def test_invalid_toml_raises(self):
        p = self._tmp("this is not toml [[[")
        with pytest.raises(ValueError):
            patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")

    def test_creates_backup(self):
        p = self._tmp(self.STDIO_CFG)
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        assert p.with_suffix(".toml.bak").exists()
        # в бэкапе — исходный stdio-конфиг
        bak = tomllib.loads(p.with_suffix(".toml.bak").read_text(encoding="utf-8"))
        assert "command" in bak["mcp_servers"]["1c-bridge"]

    def test_remove_section(self):
        p = self._tmp(self.STDIO_CFG)
        changed = patch_codex_config_toml(p, "1c-bridge", url=None)
        assert changed is True
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        assert "1c-bridge" not in data["mcp_servers"]
        # соседние секции целы
        assert data["mcp_servers"]["node_repl"]["command"] == "C:\\node.exe"
        assert data["plugins"]["browser@openai-bundled"]["enabled"] is True

    def test_remove_missing_section_is_noop(self):
        p = self._tmp('[desktop]\nfontSize = 14\n')
        assert patch_codex_config_toml(p, "1c-bridge", url=None) is False

    def test_crlf_preserved(self):
        # фикстура может сама содержать CRLF (файл .py на Windows), нормализуем
        cfg = self.STDIO_CFG.replace("\r\n", "\n").replace("\n", "\r\n")
        p = self._tmp(cfg)
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/TOKEN")
        raw = p.read_bytes()
        assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b"")
        data = tomllib.loads(raw.decode("utf-8"))
        assert data["mcp_servers"]["1c-bridge"]["url"] == "http://127.0.0.1:8000/mcp/TOKEN"

    def test_url_with_quotes_escaped(self):
        p = self._tmp('[desktop]\nfontSize = 14\n')
        tricky = 'http://127.0.0.1:8000/mcp/a"b\\c'
        patch_codex_config_toml(p, "1c-bridge", tricky)
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        assert data["mcp_servers"]["1c-bridge"]["url"] == tricky

    def test_bad_server_name_rejected(self):
        p = self._tmp('')
        with pytest.raises(ValueError):
            patch_codex_config_toml(p, 'bad name; inject = "x"', "http://x")

    def test_user_settings_preserved(self):
        """Патчер не трогает настройки пользователя: timeouts, approval, tools.*."""
        cfg = (
            "[mcp_servers.1c-bridge]\n"
            'enabled = true\n'
            'tool_timeout_sec = 180\n'
            'required = true\n'
            'enabled_tools = ["list_databases", "execute_query"]\n'
            'default_tools_approval_mode = "writes"\n'
            'url = "http://127.0.0.1:8000/mcp/OLD"\n'
            'command = "C:\\\\Program Files\\\\1cMcpBridge\\\\.venv\\\\Scripts\\\\python.exe"\n'
            'args = ["C:\\\\Program Files\\\\1cMcpBridge\\\\mcp_server_1c.py"]\n'
            "\n"
            "[mcp_servers.1c-bridge.env]\n"
            'ONEC_DATABASES_FILE = "C:\\\\ProgramData\\\\1cMcpBridge\\\\databases.json"\n'
            "\n"
            "[mcp_servers.1c-bridge.tools.execute_query]\n"
            'approval_mode = "prompt"\n'
            "\n"
            "[desktop]\n"
            "fontSize = 14\n"
        )
        p = self._tmp(cfg)
        changed = patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/NEW")
        assert changed is True
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        srv = data["mcp_servers"]["1c-bridge"]
        # url обновлён, транспорт удалён
        assert srv["url"] == "http://127.0.0.1:8000/mcp/NEW"
        assert "command" not in srv and "args" not in srv
        # пользовательские настройки на месте
        assert srv["enabled"] is True
        assert srv["tool_timeout_sec"] == 180
        assert srv["required"] is True
        assert srv["enabled_tools"] == ["list_databases", "execute_query"]
        assert srv["default_tools_approval_mode"] == "writes"
        # подсекция .env удалена, а .tools.execute_query сохранена
        assert "env" not in data["mcp_servers"]["1c-bridge"]
        assert data["mcp_servers"]["1c-bridge"]["tools"]["execute_query"]["approval_mode"] == "prompt"

    def test_environs_subtable_not_touched(self):
        """Подсекция с похожим именем (.environs) не удаляется патчером."""
        cfg = (
            "[mcp_servers.1c-bridge]\n"
            'url = "http://127.0.0.1:8000/mcp/OLD"\n'
            "\n"
            "[mcp_servers.1c-bridge.env]\n"
            'ONEC_DATABASES_FILE = "C:\\\\x"\n'
            "\n"
            "[mcp_servers.1c-bridge.environs]\n"
            "foo = 1\n"
        )
        p = self._tmp(cfg)
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/NEW")
        data = tomllib.loads(p.read_text(encoding="utf-8"))
        srv = data["mcp_servers"]["1c-bridge"]
        assert srv["url"] == "http://127.0.0.1:8000/mcp/NEW"
        assert "env" not in srv
        assert data["mcp_servers"]["1c-bridge"]["environs"]["foo"] == 1

    def test_multiline_args_removed_cleanly(self):
        """Многострочное значение args = [...] удаляется целиком, результат
        остаётся валидным TOML."""
        cfg = (
            "[mcp_servers.1c-bridge]\n"
            'enabled = true\n'
            'url = "http://127.0.0.1:8000/mcp/OLD"\n'
            'args = [\n'
            '  "C:\\\\Program Files\\\\1cMcpBridge\\\\mcp_server_1c.py",\n'
            '  "--flag",\n'
            ']\n'
            "\n"
            "[desktop]\n"
            "fontSize = 14\n"
        )
        p = self._tmp(cfg)
        patch_codex_config_toml(p, "1c-bridge", "http://127.0.0.1:8000/mcp/NEW")
        raw = p.read_text(encoding="utf-8")
        # результат валиден и не содержит остатков массива
        data = tomllib.loads(raw)
        srv = data["mcp_servers"]["1c-bridge"]
        assert srv["url"] == "http://127.0.0.1:8000/mcp/NEW"
        assert srv["enabled"] is True
        assert "args" not in srv
        assert "--flag" not in raw

    def test_cli_stdout_redacts_token(self):
        """CLI не выводит токен в stdout."""
        import subprocess
        import sys

        p = self._tmp("[desktop]\nfontSize = 14\n")
        url = "http://127.0.0.1:8000/mcp/SECRETTOKEN123"
        script = Path(__file__).resolve().parent.parent / "src" / "clients_config.py"
        proc = subprocess.run(
            [sys.executable, str(script), "patch-codex", "--path", str(p), "--server", "1c-bridge", "--url", url],
            capture_output=True,
            text=True,
            check=False,
        )
        assert proc.returncode == 0, proc.stderr
        assert "SECRETTOKEN123" not in proc.stdout
        assert "url=http://127.0.0.1:8000/mcp/<token>" in proc.stdout
