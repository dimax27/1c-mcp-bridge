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
