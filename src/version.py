"""Версия 1C MCP Bridge.

Единый источник версии для сервера и диагностики.
При сборке установщика (GitHub Actions / Inno Setup) значение запекается
из имени тега в src/version.py; при запуске из исходников без APP_VERSION
честно показывается "dev" — чтобы версия исходников никогда не «отставала»
от опубликованного релиза.
"""

import os

VERSION = os.environ.get("APP_VERSION", "dev").lstrip("v")
