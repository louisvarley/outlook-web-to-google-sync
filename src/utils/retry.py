"""Exponential backoff retry decorator for HTTP API calls."""
from __future__ import annotations

import functools
import logging
import random
import time
from typing import Callable, Optional

logger = logging.getLogger("calendar_sync")

_RETRYABLE_HTTP_CODES = {429, 500, 502, 503, 504}


def _classify_exception(exc: Exception) -> tuple[bool, Optional[int]]:
    """Return (should_retry, retry_after_seconds_or_None)."""
    # requests.HTTPError
    try:
        import requests

        if isinstance(exc, requests.HTTPError):
            resp = exc.response
            status = resp.status_code if resp is not None else 0
            if status in _RETRYABLE_HTTP_CODES:
                retry_after: Optional[int] = None
                if resp is not None:
                    header = resp.headers.get("Retry-After")
                    if header:
                        try:
                            retry_after = int(header)
                        except ValueError:
                            pass
                return True, retry_after
            return False, None
    except ImportError:
        pass

    # Google API HttpError
    try:
        from googleapiclient.errors import HttpError  # type: ignore

        if isinstance(exc, HttpError):
            if int(exc.resp.status) in _RETRYABLE_HTTP_CODES:
                header = exc.resp.get("retry-after")
                retry_after = int(header) if header else None
                return True, retry_after
            return False, None
    except ImportError:
        pass

    return False, None


def with_retry(max_attempts: int = 3, base_delay: float = 1.0) -> Callable:
    """Decorator that retries on retryable HTTP errors with exponential backoff + jitter."""

    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc: Optional[Exception] = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as exc:
                    retryable, retry_after = _classify_exception(exc)
                    if not retryable or attempt == max_attempts:
                        raise
                    last_exc = exc
                    if retry_after is not None:
                        delay = float(retry_after)
                    else:
                        delay = base_delay * (2 ** (attempt - 1)) + random.uniform(0, 1)
                    logger.warning(
                        "Attempt %d/%d failed (%s). Retrying in %.1fs…",
                        attempt,
                        max_attempts,
                        exc,
                        delay,
                    )
                    time.sleep(delay)
            raise last_exc  # type: ignore[misc]

        return wrapper

    return decorator
