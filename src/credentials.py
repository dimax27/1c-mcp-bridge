"""
Credential encryption via Windows DPAPI (Data Protection API).

Provides transparent encryption of 1C passwords stored in databases.json.
Backward-compatible: plaintext Pwd= in connection_string still works.
Migration from plaintext to encrypted happens automatically on first load.
"""

from __future__ import annotations

import base64
import logging
import re
from typing import Any

log = logging.getLogger("mcp-1c.credentials")

# ---------------------------------------------------------------------------
# DPAPI wrapper — uses Windows CryptProtectData / CryptUnprotectData
# Falls back gracefully if pywin32 crypto is unavailable
# ---------------------------------------------------------------------------

try:
    import win32crypt  # type: ignore[import-untyped]

    def _encrypt(plaintext: str) -> str:
        """Encrypt a string with DPAPI (current user scope)."""
        blob = win32crypt.CryptProtectData(
            plaintext.encode("utf-16-le"),
            "1c-mcp-bridge",
            None, None, None,
            0,  # CRYPTPROTECT_LOCAL_MACHINE not set = current user only
        )
        return base64.b64encode(blob).decode("ascii")

    def _decrypt(b64blob: str) -> str:
        """Decrypt a DPAPI-encrypted base64 blob."""
        blob = base64.b64decode(b64blob)
        return (
            win32crypt.CryptUnprotectData(blob, None, None, None, 0)[1]
            .decode("utf-16-le")
        )

    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    log.warning(
        "pywin32 crypto not available — DPAPI disabled, passwords remain plaintext"
    )

    def _encrypt(plaintext: str) -> str:
        raise RuntimeError("pywin32 crypto not available")

    def _decrypt(b64blob: str) -> str:
        raise RuntimeError("pywin32 crypto not available")


# ---------------------------------------------------------------------------
# Connection string handling
# ---------------------------------------------------------------------------

_PWD_RE = re.compile(r';?Pwd\s*=\s*"([^"]*)"', re.IGNORECASE)
_USR_RE = re.compile(r';?Usr\s*=\s*"([^"]*)"', re.IGNORECASE)


def extract_credential(conn_str: str) -> tuple[str, str | None]:
    """Extract username and password from a 1C connection string.

    Returns (clean_conn_str, encrypted_blob_or_none).
    """
    m = _PWD_RE.search(conn_str)
    if not m:
        return conn_str, None
    pwd = m.group(1)
    clean = _PWD_RE.sub("", conn_str)  # remove Pwd=... entirely
    if not CRYPTO_AVAILABLE:
        return conn_str, None  # keep as-is
    blob = _encrypt(pwd)
    return clean, blob


def build_conn_str(db_cfg: dict[str, Any]) -> str:
    """Build a complete connection string from a database config dict.

    Supports both plaintext (legacy) and DPAPI-encrypted credentials.
    """
    conn_str = db_cfg.get("connection_string", "")
    if not conn_str:
        raise ValueError("connection_string is required")

    credential = db_cfg.get("credential")
    if not credential:
        return conn_str  # plaintext, legacy

    provider = credential.get("provider", "")
    blob = credential.get("blob", "")
    if provider != "dpapi-current-user" or not blob:
        return conn_str

    if not CRYPTO_AVAILABLE:
        log.warning("DPAPI not available, cannot decrypt credentials")
        raise RuntimeError("DPAPI not available")

    pwd = _decrypt(blob)
    # Remove any existing Pwd= and add the decrypted one
    conn_str = _PWD_RE.sub("", conn_str)
    if conn_str.endswith(";"):
        conn_str += f'Pwd="{pwd}"'
    else:
        conn_str += f';Pwd="{pwd}"'
    return conn_str


def migrate_to_encrypted(db_cfg: dict[str, Any]) -> bool:
    """Migrate a plaintext Pwd= to DPAPI-encrypted credential.

    Returns True if migration happened, False if already encrypted or not needed.
    """
    conn_str = db_cfg.get("connection_string", "")
    m = _PWD_RE.search(conn_str)
    if not m:
        return False  # no password to encrypt

    if not CRYPTO_AVAILABLE:
        return False  # can't encrypt

    clean, blob = extract_credential(conn_str)
    if blob is None:
        return False

    db_cfg["connection_string"] = clean
    db_cfg["credential"] = {"provider": "dpapi-current-user", "blob": blob}
    return True
