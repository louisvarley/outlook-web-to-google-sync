"""Configuration management — load/save config.json."""
from __future__ import annotations

import copy
import json
import stat
from pathlib import Path
from typing import Any

CONFIG_PATH = Path(__file__).parent.parent / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "microsoft": {
        "client_id": "",
        "tenant_id": "",
        "client_secret": "",
        "account_type": "personal",  # "personal" | "work"
        "calendar_id": "",
        "sync_categories": [],
        "token_cache": {},
    },
    "google": {
        "client_id": "",
        "client_secret": "",
        "calendar_id": "",
        "color_filter": [],  # list of int colorIds; empty = sync all
        "token": {},
    },
    "sync": {
        "lookback_days": 30,
        "lookahead_days": 365,
        "initial_lookback_days": 365,
        "interval_minutes": 15,
        "state_db_path": "sync_state.db",
        "log_dir": "logs",
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge override into a copy of base."""
    result = copy.deepcopy(base)
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config() -> dict[str, Any]:
    """Load config from disk, merging with defaults for any missing keys."""
    if not CONFIG_PATH.exists():
        return copy.deepcopy(DEFAULT_CONFIG)
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        on_disk = json.load(f)
    return _deep_merge(DEFAULT_CONFIG, on_disk)


def save_config(config: dict[str, Any]) -> None:
    """Persist config to disk. Never logs credential values."""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)
    # Restrict file permissions (owner read/write only)
    try:
        CONFIG_PATH.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except Exception:
        pass  # Best-effort on Windows


def config_exists() -> bool:
    return CONFIG_PATH.exists()
