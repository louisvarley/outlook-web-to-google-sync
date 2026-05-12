"""SQLite sync-state tracking: Outlook event ID ↔ Google event ID pairs.

Schema
------
sync_pairs:
  - outlook_event_id / google_event_id  UNIQUE — primary keys for each side
  - *_last_modified                     ISO-8601 strings of the last known mtime
  - last_sync_time                      when this pair was last written
  - event_type                          'single' | 'master' | 'exception'
  - outlook_series_master_id            populated for exception instances (Outlook)
  - google_recurring_event_id           populated for exception instances (Google)

sync_metadata:
  - key/value store (e.g. last_sync_time)
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Generator, Optional


@dataclass
class SyncPair:
    outlook_event_id: str
    google_event_id: str
    outlook_last_modified: Optional[str] = None
    google_last_modified: Optional[str] = None
    last_sync_time: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )
    event_type: str = "single"  # "single" | "master" | "exception"
    outlook_series_master_id: Optional[str] = None
    google_recurring_event_id: Optional[str] = None


class SyncStateDB:
    def __init__(self, db_path: str = "sync_state.db") -> None:
        self._path = Path(db_path)
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self.init_db()

    @contextmanager
    def _conn(self) -> Generator[sqlite3.Connection, None, None]:
        conn = sqlite3.connect(self._path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def init_db(self) -> None:
        with self._conn() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS sync_pairs (
                    id                        INTEGER PRIMARY KEY AUTOINCREMENT,
                    outlook_event_id          TEXT NOT NULL,
                    google_event_id           TEXT NOT NULL,
                    outlook_last_modified     TEXT,
                    google_last_modified      TEXT,
                    last_sync_time            TEXT NOT NULL,
                    event_type                TEXT NOT NULL DEFAULT 'single',
                    outlook_series_master_id  TEXT,
                    google_recurring_event_id TEXT,
                    UNIQUE(outlook_event_id),
                    UNIQUE(google_event_id)
                );

                CREATE TABLE IF NOT EXISTS sync_metadata (
                    key   TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                """
            )

    # ── Lookups ───────────────────────────────────────────────────────────────

    def get_pair_by_outlook_id(self, outlook_id: str) -> Optional[SyncPair]:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM sync_pairs WHERE outlook_event_id = ?",
                (outlook_id,),
            ).fetchone()
        return _row_to_pair(row) if row else None

    def get_pair_by_google_id(self, google_id: str) -> Optional[SyncPair]:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM sync_pairs WHERE google_event_id = ?",
                (google_id,),
            ).fetchone()
        return _row_to_pair(row) if row else None

    def get_all_pairs(self) -> list[SyncPair]:
        with self._conn() as conn:
            rows = conn.execute("SELECT * FROM sync_pairs").fetchall()
        return [_row_to_pair(r) for r in rows]

    # ── Mutations ─────────────────────────────────────────────────────────────

    def upsert_pair(self, pair: SyncPair) -> None:
        with self._conn() as conn:
            conn.execute(
                """
                INSERT INTO sync_pairs
                    (outlook_event_id, google_event_id, outlook_last_modified,
                     google_last_modified, last_sync_time, event_type,
                     outlook_series_master_id, google_recurring_event_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(outlook_event_id) DO UPDATE SET
                    google_event_id           = excluded.google_event_id,
                    outlook_last_modified     = excluded.outlook_last_modified,
                    google_last_modified      = excluded.google_last_modified,
                    last_sync_time            = excluded.last_sync_time,
                    event_type                = excluded.event_type,
                    outlook_series_master_id  = excluded.outlook_series_master_id,
                    google_recurring_event_id = excluded.google_recurring_event_id
                """,
                (
                    pair.outlook_event_id,
                    pair.google_event_id,
                    pair.outlook_last_modified,
                    pair.google_last_modified,
                    pair.last_sync_time,
                    pair.event_type,
                    pair.outlook_series_master_id,
                    pair.google_recurring_event_id,
                ),
            )

    def delete_pair_by_outlook_id(self, outlook_id: str) -> None:
        with self._conn() as conn:
            conn.execute(
                "DELETE FROM sync_pairs WHERE outlook_event_id = ?", (outlook_id,)
            )

    def delete_pair_by_google_id(self, google_id: str) -> None:
        with self._conn() as conn:
            conn.execute(
                "DELETE FROM sync_pairs WHERE google_event_id = ?", (google_id,)
            )

    # ── Metadata ──────────────────────────────────────────────────────────────

    def get_metadata(self, key: str) -> Optional[str]:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT value FROM sync_metadata WHERE key = ?", (key,)
            ).fetchone()
        return row["value"] if row else None

    def set_metadata(self, key: str, value: str) -> None:
        with self._conn() as conn:
            conn.execute(
                "INSERT INTO sync_metadata (key, value) VALUES (?, ?)"
                " ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, value),
            )

    def pair_count(self) -> int:
        with self._conn() as conn:
            row = conn.execute("SELECT COUNT(*) AS n FROM sync_pairs").fetchone()
        return row["n"]


def _row_to_pair(row: sqlite3.Row) -> SyncPair:
    return SyncPair(
        outlook_event_id=row["outlook_event_id"],
        google_event_id=row["google_event_id"],
        outlook_last_modified=row["outlook_last_modified"],
        google_last_modified=row["google_last_modified"],
        last_sync_time=row["last_sync_time"],
        event_type=row["event_type"],
        outlook_series_master_id=row["outlook_series_master_id"],
        google_recurring_event_id=row["google_recurring_event_id"],
    )
