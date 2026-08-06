"""
COM query timeout protection.

Since Python threads cannot be safely interrupted during COM calls,
true timeout requires process isolation (planned for v0.6.0).

For now: log warnings for slow queries and enforce hard response limits.
"""

from __future__ import annotations

import logging
import time

log = logging.getLogger("mcp-1c.timeout")

# Warn if query takes longer than this (seconds)
SLOW_QUERY_WARNING = 30.0


def check_slow(start_time: float, db_key: str, query_preview: str) -> None:
    """Log a warning if the query has been running too long."""
    elapsed = time.perf_counter() - start_time
    if elapsed > SLOW_QUERY_WARNING:
        log.warning(
            "Slow query on '%s': %.1fs — %s...",
            db_key,
            elapsed,
            query_preview[:200],
        )
