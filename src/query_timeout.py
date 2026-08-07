"""
COM query timeout protection.

Since Python threads cannot be safely interrupted during COM calls,
true timeout requires process isolation (planned for v0.6.0).

For now: log warnings for slow queries and enforce hard response limits.
"""

from __future__ import annotations

import hashlib
import logging
import time

log = logging.getLogger("mcp-1c.timeout")

# Warn if query takes longer than this (seconds)
SLOW_QUERY_WARNING = 30.0


def check_slow(start_time: float, db_key: str, query_text: str) -> None:
    """Log a warning if the query has been running too long.

    Текст запроса может содержать внутренние имена, номера документов и
    строковые литералы — в журнал пишем только SHA-256 хэш. Полный превью
    доступен на уровне DEBUG (по умолчанию выключен).
    """
    elapsed = time.perf_counter() - start_time
    if elapsed > SLOW_QUERY_WARNING:
        digest = hashlib.sha256(query_text.encode("utf-8", "replace")).hexdigest()[:16]
        log.warning(
            "Slow query on '%s': %.1fs (query hash %s, preview hidden — use DEBUG to show)",
            db_key,
            elapsed,
            digest,
        )
        if log.isEnabledFor(logging.DEBUG):
            log.debug("Slow query text on '%s': %s...", db_key, query_text[:200])
