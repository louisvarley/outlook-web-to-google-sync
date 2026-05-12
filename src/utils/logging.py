"""Rich-based structured logger with dated log files."""
from __future__ import annotations

import logging
from datetime import date
from pathlib import Path
from typing import Optional

from rich.console import Console
from rich.logging import RichHandler
from rich.theme import Theme

_CONSOLE_THEME = Theme(
    {
        "log.created": "green bold",
        "log.updated": "yellow bold",
        "log.deleted": "red bold",
        "log.skipped": "dim",
        "log.error": "red reverse bold",
        "log.conflict": "magenta bold",
    }
)

_console = Console(theme=_CONSOLE_THEME)
_initialised = False


def init_logger(log_dir: str = "logs") -> logging.Logger:
    """Initialise the logger (idempotent — safe to call multiple times)."""
    global _initialised
    logger = logging.getLogger("calendar_sync")
    if _initialised:
        return logger

    logger.setLevel(logging.DEBUG)

    # Rich console handler
    console_handler = RichHandler(
        console=_console,
        show_path=False,
        markup=True,
        rich_tracebacks=True,
    )
    console_handler.setLevel(logging.INFO)
    logger.addHandler(console_handler)

    # File handler with full detail
    log_path = Path(log_dir)
    log_path.mkdir(parents=True, exist_ok=True)
    log_file = log_path / f"sync-{date.today().isoformat()}.log"
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s  %(levelname)-8s  %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S",
        )
    )
    logger.addHandler(file_handler)
    _initialised = True
    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger("calendar_sync")


def log_action(action: str, subject: str, direction: str = "", extra: str = "") -> None:
    """Log a sync action with a consistent format."""
    logger = get_logger()
    parts = [f"[{action}]"]
    if direction:
        parts.append(direction)
    parts.append(subject)
    if extra:
        parts.append(f"({extra})")
    logger.info(" ".join(parts))
