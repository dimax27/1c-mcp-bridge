"""Тесты логики Manager: сохранение DPAPI-credential при редактировании базы.

manager.py импортирует tkinter — на Windows он доступен, окно не создаётся.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "installer"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from manager import assemble_db_config, resolve_connection_string

LEGACY_CFG = {
    "enabled": True,
    "description": "УТ 10.3",
    "progid": "V83.COMConnector",
    "connection_string": (
        "Srvr=192.168.0.35;Ref=ut10;Usr=\"Кувыкин Д.А.\";Pwd=\"1с\""
    ),
}

DPAPI_CFG = {
    "enabled": True,
    "description": "УТ 10.3",
    "progid": "V83.COMConnector",
    "connection_string": "Srvr=192.168.0.35;Ref=ut10;Usr=\"Кувыкин Д.А.\"",
    "notes": "",
    "credential": {"provider": "dpapi-current-user", "blob": b"\x01\x02\x03"},
}


def test_assemble_preserves_credential_and_strips_pwd():
    cfg = assemble_db_config(
        DPAPI_CFG,
        enabled=True,
        description="Новое описание",
        progid="V83.COMConnector",
        connection_string='Srvr=192.168.0.35;Ref=ut10;Usr="Кувыкин Д.А.";Pwd="другое"',
        notes="изменили описание",
        password_modified=False,
    )
    assert cfg["description"] == "Новое описание"
    assert cfg["notes"] == "изменили описание"
    # credential сохранён, Pwd= вырезан из строки
    assert cfg["credential"] == DPAPI_CFG["credential"]
    assert ";Pwd=" not in cfg["connection_string"]
    assert cfg["connection_string"].endswith('Usr="Кувыкин Д.А."')


def test_assemble_new_password_replaces_credential():
    cfg = assemble_db_config(
        DPAPI_CFG,
        enabled=True,
        description="УТ",
        progid="V83.COMConnector",
        connection_string='Srvr=192.168.0.35;Ref=ut10;Usr="Кувыкин Д.А.";Pwd="новый"',
        notes="",
        password_modified=True,
    )
    assert "credential" not in cfg
    assert 'Pwd="новый"' in cfg["connection_string"]


def test_assemble_new_base_has_no_credential():
    cfg = assemble_db_config(
        None,
        enabled=True,
        description="Новая",
        progid="V83.COMConnector",
        connection_string='Srvr=1.2.3.4;Ref=x;Usr="u";Pwd="p"',
        notes="",
    )
    assert "credential" not in cfg


def test_assemble_dll_path_preserved():
    cfg = assemble_db_config(
        None,
        enabled=True,
        description="",
        progid="V83.COMConnector",
        connection_string="Srvr=1.2.3.4;Ref=x",
        notes="",
        dll_path=r"C:\Program Files\1cv8\bin\comcntr.dll",
    )
    assert cfg["dll_path"] == r"C:\Program Files\1cv8\bin\comcntr.dll"


def test_resolve_plaintext_passthrough():
    conn = 'Srvr=1.2.3.4;Ref=x;Usr="u";Pwd="p"'
    assert resolve_connection_string(LEGACY_CFG, conn, False) == conn


def test_resolve_dpapi_decrypts_password():
    """Реальный DPAPI-круг через migrate_to_encrypted (как при загрузке базы):
    шифруем пароль и проверяем, что resolve_connection_string возвращает
    строку с расшифрованным Pwd=."""
    pytest.importorskip("win32crypt")
    from credentials import build_conn_str, migrate_to_encrypted

    secret = "Секретное-1с"
    cfg = {"connection_string": f'Srvr=1.2.3.4;Ref=x;Usr="u";Pwd="{secret}"'}
    assert migrate_to_encrypted(cfg) is True
    assert cfg["credential"]["provider"] == "dpapi-current-user"
    assert "Pwd=" not in cfg["connection_string"]

    # расшифровка работает (build_conn_str — то, что вызывает test_connection)
    assert f'Pwd="{secret}"' in build_conn_str(cfg)

    # через resolve_connection_string Manager: пароль не меняли → расшифровывается
    resolved = resolve_connection_string(cfg, cfg["connection_string"], False)
    assert f'Pwd="{secret}"' in resolved

    # пароль меняли → строка берётся как есть (без расшифровки)
    new_conn = 'Srvr=1.2.3.4;Ref=x;Usr="u";Pwd="new"'
    assert resolve_connection_string(cfg, new_conn, True) == new_conn
