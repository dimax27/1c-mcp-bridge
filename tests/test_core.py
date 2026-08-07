"""Tests for 1C MCP Bridge — pure Python, no COM/1C required."""

import json
import os
import sys
import tempfile
from pathlib import Path

import pytest

SRC = str(Path(__file__).resolve().parent.parent / "src")
sys.path.insert(0, SRC)


# ---------------------------------------------------------------------------
# clients_config
# ---------------------------------------------------------------------------

from clients_config import (
    KNOWN_CLIENTS,
    _mcp_key,
    client_config_path,
    read_client_config,
    remove_mcp_server,
    update_mcp_servers,
    write_client_config,
)


class TestClientConfig:
    def test_known_clients_structure(self):
        required = {"id", "name", "appdata_dir", "config_name", "download_url"}
        for c in KNOWN_CLIENTS:
            missing = required - set(c.keys())
            assert not missing, f"Client {c.get('id', '?')} missing {missing}"

    def test_client_config_path_claude(self):
        claude = KNOWN_CLIENTS[0]
        p = client_config_path(claude)
        assert p.name == "claude_desktop_config.json"
        assert "Claude" in str(p)

    def test_client_config_path_qwen(self):
        qwen = KNOWN_CLIENTS[1]
        p = client_config_path(qwen)
        assert p.name == "settings.json"
        assert "Qwen" in str(p)

    def test_mcp_key_default(self):
        assert _mcp_key(KNOWN_CLIENTS[0]) == "mcpServers"

    def test_mcp_key_qwen(self):
        assert _mcp_key(KNOWN_CLIENTS[1]) == "mcp_config"

    def test_read_missing_config(self):
        fake = {"id": "test", "name": "Test", "appdata_dir": "Nope", "config_name": "nope.json"}
        assert read_client_config(fake) == {}

    def test_read_corrupt_config(self):
        d = tempfile.mkdtemp()
        try:
            p = Path(d) / "bad.json"
            p.write_text("{not valid", encoding="utf-8")
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "bad.json"}
            with pytest.raises(ValueError, match="Failed to parse"):
                read_client_config(fake)
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_write_and_read_roundtrip(self):
        d = tempfile.mkdtemp()
        try:
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "test.json"}
            config = {"mcpServers": {"test-server": {"command": "echo"}}}
            write_client_config(fake, config)
            result = read_client_config(fake)
            assert result == config
            bak = Path(d) / "test.json.bak"
            assert not bak.exists()
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_write_creates_backup(self):
        d = tempfile.mkdtemp()
        try:
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "test.json"}
            c1 = {"mcpServers": {"old": {"command": "old"}}}
            c2 = {"mcpServers": {"new": {"command": "new"}}}
            write_client_config(fake, c1)
            write_client_config(fake, c2)
            result = read_client_config(fake)
            assert result == c2
            bak = Path(d) / "test.json.bak"
            assert bak.exists()
            assert json.loads(bak.read_text(encoding="utf-8")) == c1
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_update_mcp_servers(self):
        d = tempfile.mkdtemp()
        try:
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "test.json"}
            entry = {"command": "npx", "args": ["test"]}
            assert update_mcp_servers(fake, "myserver", entry) is True
            assert read_client_config(fake)["mcpServers"]["myserver"] == entry
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_update_no_change_returns_false(self):
        d = tempfile.mkdtemp()
        try:
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "test.json"}
            entry = {"command": "npx"}
            update_mcp_servers(fake, "myserver", entry)
            assert update_mcp_servers(fake, "myserver", entry) is False
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_remove_mcp_server(self):
        d = tempfile.mkdtemp()
        try:
            fake = {"id": "test", "name": "Test", "appdata_dir": d, "config_name": "test.json"}
            update_mcp_servers(fake, "myserver", {"command": "x"})
            assert remove_mcp_server(fake, "myserver") is True
            assert remove_mcp_server(fake, "myserver") is False
            assert "myserver" not in read_client_config(fake).get("mcpServers", {})
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)


# ---------------------------------------------------------------------------
# Config (env validation + limits)
# ---------------------------------------------------------------------------

from config import (
    DEFAULT_LIMIT,
    HARD_LIMIT,
    MAX_COLUMNS,
    MAX_PARAMETERS,
    MAX_QUERY_LENGTH,
    positive_env_int,
)


class TestPositiveEnvInt:
    def test_default(self):
        old = os.environ.pop("_TEST_VAL", None)
        try:
            assert positive_env_int("_TEST_VAL", 200, 10000) == 200
        finally:
            if old is not None:
                os.environ["_TEST_VAL"] = old

    def test_custom(self):
        os.environ["_TEST_VAL"] = "500"
        try:
            assert positive_env_int("_TEST_VAL", 100, 1000) == 500
        finally:
            del os.environ["_TEST_VAL"]

    def test_out_of_range(self):
        os.environ["_TEST_VAL"] = "0"
        try:
            with pytest.raises(RuntimeError, match="must be 1"):
                positive_env_int("_TEST_VAL", 100, 1000)
        finally:
            del os.environ["_TEST_VAL"]

    def test_non_integer(self):
        os.environ["_TEST_VAL"] = "abc"
        try:
            with pytest.raises(RuntimeError, match="must be an integer"):
                positive_env_int("_TEST_VAL", 100, 1000)
        finally:
            del os.environ["_TEST_VAL"]


class TestLimits:
    def test_defaults_positive(self):
        assert DEFAULT_LIMIT > 0
        assert HARD_LIMIT > DEFAULT_LIMIT
        assert MAX_QUERY_LENGTH > 100
        assert MAX_COLUMNS > 1
        assert MAX_PARAMETERS > 1

    def test_hard_limit_bounds_default(self):
        """HARD_LIMIT >= DEFAULT_LIMIT by construction."""
        assert HARD_LIMIT >= DEFAULT_LIMIT
