"""Shared configuration — no COM imports, safe to import in tests."""

import os


def positive_env_int(name: str, default: int, maximum: int) -> int:
    """Validate integer environment variable with range check."""
    raw = os.environ.get(name, str(default))
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{name} must be an integer, got: {raw!r}") from exc
    if not 1 <= value <= maximum:
        raise RuntimeError(f"{name} must be 1..{maximum}, got: {value}")
    return value


HARD_LIMIT = positive_env_int("ONEC_HARD_LIMIT", 10000, 100000)
DEFAULT_LIMIT = positive_env_int("ONEC_DEFAULT_LIMIT", 1000, HARD_LIMIT)
MAX_QUERY_LENGTH = positive_env_int("ONEC_MAX_QUERY_LENGTH", 10000, 1000000)
MAX_COLUMNS = positive_env_int("ONEC_MAX_COLUMNS", 200, 10000)
MAX_PARAMETERS = positive_env_int("ONEC_MAX_PARAMETERS", 50, 1000)

# ---------------------------------------------------------------------------
# Лимиты журнала регистрации (get_registration_log)
# ---------------------------------------------------------------------------

EVENT_LOG_DEFAULT_LIMIT = positive_env_int(
    "ONEC_EVENT_LOG_DEFAULT_LIMIT", 200, 1000
)
EVENT_LOG_HARD_LIMIT = positive_env_int(
    "ONEC_EVENT_LOG_HARD_LIMIT", 1000, 10000
)
EVENT_LOG_MAX_RANGE_DAYS = positive_env_int(
    "ONEC_EVENT_LOG_MAX_RANGE_DAYS", 93, 366
)
EVENT_LOG_MAX_BROAD_RANGE_DAYS = positive_env_int(
    "ONEC_EVENT_LOG_MAX_BROAD_RANGE_DAYS", 7, 31
)
EVENT_LOG_MAX_DOCUMENTS = positive_env_int(
    "ONEC_EVENT_LOG_MAX_DOCUMENTS", 20, 100
)
EVENT_LOG_MAX_EVENTS = positive_env_int(
    "ONEC_EVENT_LOG_MAX_EVENTS", 20, 100
)
EVENT_LOG_MAX_REFS = positive_env_int(
    "ONEC_EVENT_LOG_MAX_REFS", 50, 500
)
EVENT_LOG_MAX_RESPONSE_BYTES = positive_env_int(
    "ONEC_EVENT_LOG_MAX_RESPONSE_BYTES", 5 * 1024 * 1024, 50 * 1024 * 1024
)

# DEFAULT <= HARD проверяется при импорте, как и для существующих лимитов запросов.
if EVENT_LOG_DEFAULT_LIMIT > EVENT_LOG_HARD_LIMIT:
    raise RuntimeError(
        "ONEC_EVENT_LOG_DEFAULT_LIMIT must not exceed ONEC_EVENT_LOG_HARD_LIMIT"
    )
