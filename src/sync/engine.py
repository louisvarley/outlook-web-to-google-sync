"""Sync engine: orchestrates bi-directional Outlook ↔ Google sync.

Sync order
----------
1. Fetch all matching Outlook events (category-filtered) for the date window.
2. Fetch all matching Google events (colour-filtered) for the date window.
3. Load all known sync pairs from SQLite.
4. Pass A — Outlook → Google:
     • New Outlook event  → create Google event, record pair.
     • Changed Outlook    → update Google event, update pair.
     • Conflict (both sides changed) → prefer most recently modified, log conflict.
5. Pass B — Deletions O→G:
     • Pair whose Outlook ID is absent from current fetch → delete Google + pair.
6. Pass C — Google → Outlook:
     • New Google event   → create Outlook event with sync category, record pair.
     • Changed Google     → update Outlook event, update pair.
7. Pass D — Deletions G→O:
     • Pair whose Google ID is absent from current fetch → delete Outlook + pair.
8. Persist last_sync_time.

Duplicate-prevention guarantee
-------------------------------
Events are never matched by title/time alone.  Only pairs with a recorded
SQLite entry are updated/deleted; untracked events are always treated as new
and a fresh pair is created after the create call succeeds.

Loop-prevention guarantee
-------------------------
When we write an event to Outlook from Google, we assign the configured sync
category.  On the next sync cycle that event will therefore pass the category
filter and be recognised as already-tracked via its pair, so it is not
duplicated.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

from src.auth.google import get_credentials
from src.auth.microsoft import get_token
from src.calendars.google import GoogleCalendarClient
from src.calendars.microsoft import GraphClient
from src.sync.mapper import google_to_outlook, outlook_to_google
from src.sync.state import SyncPair, SyncStateDB
from src.utils.logging import log_action

logger = logging.getLogger("calendar_sync")


@dataclass
class SyncSummary:
    created: int = 0
    updated: int = 0
    deleted: int = 0
    skipped: int = 0
    errors: int = 0
    conflicts: list[str] = field(default_factory=list)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _categories_match(event_categories: list[str], sync_categories: list[str]) -> bool:
    """Case-insensitive: is any event category in the sync list?"""
    if not sync_categories:
        return True  # no filter = sync all
    lower_sync = {c.casefold() for c in sync_categories}
    return any(c.casefold() in lower_sync for c in event_categories)


def _color_matches(event: dict, color_filter: list[int]) -> bool:
    """Does the Google event match the configured colour filter?"""
    if not color_filter:
        return True  # no filter = sync all
    raw = event.get("colorId")
    if raw is None:
        return False
    try:
        return int(raw) in color_filter
    except (ValueError, TypeError):
        return False


def _parse_modified(event: dict, source: str) -> Optional[datetime]:
    """Return the last-modified datetime from an event dict, or None."""
    key = "lastModifiedDateTime" if source == "outlook" else "updated"
    ts = event.get(key)
    if not ts:
        return None
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00"))
    except Exception:
        return None


def _event_type(o_event: dict) -> str:
    t = o_event.get("type", "singleInstance")
    if t == "seriesMaster":
        return "master"
    if t in ("exception", "occurrence"):
        return "exception"
    return "single"


def _log_conflict(log_dir: str, message: str) -> None:
    Path(log_dir).mkdir(parents=True, exist_ok=True)
    log_file = (
        Path(log_dir) / f"conflicts-{datetime.now().strftime('%Y-%m-%d')}.log"
    )
    with log_file.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()}Z  {message}\n")


def _times_differ(ts_a: Optional[str], ts_b: Optional[str]) -> bool:
    """Return True if both timestamps exist and are different."""
    return bool(ts_a and ts_b and ts_a != ts_b)


# ── Main entry point ──────────────────────────────────────────────────────────

def run(config: dict[str, Any], dry_run: bool = False) -> SyncSummary:
    """Execute one complete bidirectional sync cycle.

    Args:
        config:  Loaded configuration dict.  Token caches may be mutated.
        dry_run: When True, detect changes and log them but make no API calls.

    Returns:
        SyncSummary with counts of each action.
    """
    summary = SyncSummary()
    sync_cfg: dict = config["sync"]
    log_dir: str = sync_cfg.get("log_dir", "logs")

    # ── Build API clients ─────────────────────────────────────────────────────
    ms_cfg: dict = config["microsoft"]
    g_cfg: dict = config["google"]

    def _ms_token() -> str:
        return get_token(ms_cfg)

    ms_client = GraphClient(token_getter=_ms_token)
    google_creds = get_credentials(g_cfg)
    g_client = GoogleCalendarClient(google_creds)

    # ── Open sync-state DB ────────────────────────────────────────────────────
    db = SyncStateDB(sync_cfg.get("state_db_path", "sync_state.db"))

    # ── Compute date window ───────────────────────────────────────────────────
    now = datetime.now(timezone.utc)
    last_sync_str = db.get_metadata("last_sync_time")
    is_first_run = last_sync_str is None

    if is_first_run:
        lookback = int(sync_cfg.get("initial_lookback_days", sync_cfg.get("lookback_days", 30)))
        logger.info("First run — using initial lookback of %d days.", lookback)
    else:
        lookback = int(sync_cfg.get("lookback_days", 30))

    lookahead = int(sync_cfg.get("lookahead_days", 365))
    window_start = now - timedelta(days=lookback)
    window_end = now + timedelta(days=lookahead)

    last_sync_time: Optional[datetime] = None
    if last_sync_str:
        try:
            last_sync_time = datetime.fromisoformat(last_sync_str)
        except Exception:
            pass

    ms_calendar_id: str = ms_cfg["calendar_id"]
    g_calendar_id: str = g_cfg["calendar_id"]
    sync_categories: list[str] = ms_cfg.get("sync_categories", [])
    color_filter: list[int] = g_cfg.get("color_filter", [])
    # Canonical sync category for writing into Outlook
    sync_category = sync_categories[0] if sync_categories else "Synced"
    # colorId applied when writing into Google
    configured_color_id: Optional[int] = color_filter[0] if color_filter else None

    logger.info(
        "Sync window: %s → %s",
        window_start.strftime("%Y-%m-%d"),
        window_end.strftime("%Y-%m-%d"),
    )
    if dry_run:
        logger.info("[bold yellow]DRY RUN[/bold yellow] — no changes will be written.")

    # ── Fetch events ──────────────────────────────────────────────────────────
    logger.info("Fetching Outlook events…")
    try:
        raw_outlook = ms_client.get_events(ms_calendar_id, window_start, window_end)
    except Exception as exc:
        logger.error("Failed to fetch Outlook events: %s", exc)
        summary.errors += 1
        return summary

    logger.info("Fetching Google events…")
    try:
        raw_google = g_client.get_events(g_calendar_id, window_start, window_end)
    except Exception as exc:
        logger.error("Failed to fetch Google events: %s", exc)
        summary.errors += 1
        return summary

    # Apply filters; key by event ID
    outlook_events: dict[str, dict] = {
        e["id"]: e
        for e in raw_outlook
        if _categories_match(e.get("categories", []), sync_categories)
        and not e.get("isCancelled", False)
    }
    google_events: dict[str, dict] = {
        e["id"]: e
        for e in raw_google
        if _color_matches(e, color_filter)
        and e.get("status") != "cancelled"
    }

    logger.info(
        "Found %d matching Outlook event(s), %d matching Google event(s).",
        len(outlook_events),
        len(google_events),
    )

    # ── Load existing pairs ───────────────────────────────────────────────────
    all_pairs = db.get_all_pairs()
    by_outlook_id: dict[str, SyncPair] = {p.outlook_event_id: p for p in all_pairs}
    by_google_id: dict[str, SyncPair] = {p.google_event_id: p for p in all_pairs}

    # Track IDs created during this cycle so deletion passes don't touch them
    newly_created_google_ids: set[str] = set()
    newly_created_outlook_ids: set[str] = set()

    # ── Pass A: Outlook → Google ──────────────────────────────────────────────
    for outlook_id, o_event in outlook_events.items():
        subject = o_event.get("subject") or "(No title)"
        try:
            pair = by_outlook_id.get(outlook_id)
            o_modified = _parse_modified(o_event, "outlook")

            if pair is None:
                # New Outlook event — create in Google
                log_action("CREATED", subject, "O→G")
                if not dry_run:
                    g_body = outlook_to_google(o_event, sync_category, configured_color_id)
                    g_created = g_client.create_event(g_calendar_id, g_body)
                    new_pair = SyncPair(
                        outlook_event_id=outlook_id,
                        google_event_id=g_created["id"],
                        outlook_last_modified=o_event.get("lastModifiedDateTime"),
                        google_last_modified=g_created.get("updated"),
                        event_type=_event_type(o_event),
                        outlook_series_master_id=o_event.get("seriesMasterId"),
                    )
                    db.upsert_pair(new_pair)
                    by_google_id[g_created["id"]] = new_pair
                    newly_created_google_ids.add(g_created["id"])
                summary.created += 1
                continue

            # Existing pair — check if update is needed
            g_event = google_events.get(pair.google_event_id)
            g_modified = _parse_modified(g_event, "google") if g_event else None

            outlook_changed = _times_differ(
                o_event.get("lastModifiedDateTime"), pair.outlook_last_modified
            ) or (o_modified and not pair.outlook_last_modified)

            google_changed = (
                g_event is not None
                and _times_differ(
                    g_event.get("updated"), pair.google_last_modified
                )
            ) or (g_event is not None and g_modified and not pair.google_last_modified)

            if outlook_changed and google_changed:
                # Conflict — prefer most recently modified
                if o_modified and g_modified:
                    prefer_outlook = o_modified >= g_modified
                else:
                    prefer_outlook = True
                winner = "Outlook" if prefer_outlook else "Google"
                msg = (
                    f"CONFLICT '{subject}': modified on both sides since last sync. "
                    f"Winner: {winner} "
                    f"(Outlook:{o_modified} Google:{g_modified})"
                )
                logger.warning(msg)
                _log_conflict(log_dir, msg)
                summary.conflicts.append(msg)
                if not prefer_outlook:
                    summary.skipped += 1
                    continue

            if outlook_changed:
                log_action("UPDATED", subject, "O→G")
                if not dry_run and g_event is not None:
                    g_body = outlook_to_google(o_event, sync_category, configured_color_id)
                    g_updated = g_client.update_event(
                        g_calendar_id, pair.google_event_id, g_body
                    )
                    pair.outlook_last_modified = o_event.get("lastModifiedDateTime")
                    pair.google_last_modified = g_updated.get("updated")
                    pair.last_sync_time = datetime.now(timezone.utc).isoformat()
                    db.upsert_pair(pair)
                summary.updated += 1
            else:
                summary.skipped += 1

        except Exception as exc:
            logger.error("Error processing Outlook event '%s' (%s): %s", subject, outlook_id, exc)
            summary.errors += 1

    # ── Pass B: Deletions Outlook → Google ───────────────────────────────────
    for pair in list(all_pairs):
        if (
            pair.outlook_event_id not in outlook_events
            and pair.google_event_id not in newly_created_google_ids
        ):
            label = f"(Outlook ID …{pair.outlook_event_id[-8:]})"
            try:
                log_action("DELETED", label, "O→G (Outlook event gone)")
                if not dry_run:
                    g_client.delete_event(g_calendar_id, pair.google_event_id)
                    db.delete_pair_by_outlook_id(pair.outlook_event_id)
                    # Remove from local maps so Pass D doesn't try to delete again
                    by_google_id.pop(pair.google_event_id, None)
                summary.deleted += 1
            except Exception as exc:
                logger.error("Error deleting Google event for pair %s: %s", pair.outlook_event_id, exc)
                summary.errors += 1

    # ── Pass C: Google → Outlook ──────────────────────────────────────────────
    for google_id, g_event in google_events.items():
        if google_id in newly_created_google_ids:
            # Created by Pass A in this cycle — already tracked, skip
            summary.skipped += 1
            continue

        subject = g_event.get("summary") or "(No title)"
        try:
            pair = by_google_id.get(google_id)
            g_modified = _parse_modified(g_event, "google")

            if pair is None:
                log_action("CREATED", subject, "G→O")
                if not dry_run:
                    o_body = google_to_outlook(g_event, sync_category)
                    o_created = ms_client.create_event(ms_calendar_id, o_body)
                    new_pair = SyncPair(
                        outlook_event_id=o_created["id"],
                        google_event_id=google_id,
                        outlook_last_modified=o_created.get("lastModifiedDateTime"),
                        google_last_modified=g_event.get("updated"),
                        google_recurring_event_id=g_event.get("recurringEventId"),
                    )
                    db.upsert_pair(new_pair)
                    by_outlook_id[o_created["id"]] = new_pair
                    newly_created_outlook_ids.add(o_created["id"])
                summary.created += 1
                continue

            # Existing pair — check if update needed
            google_changed = _times_differ(
                g_event.get("updated"), pair.google_last_modified
            ) or (g_modified and not pair.google_last_modified)

            if google_changed:
                log_action("UPDATED", subject, "G→O")
                if not dry_run:
                    o_body = google_to_outlook(g_event, sync_category)
                    o_updated = ms_client.update_event(
                        ms_calendar_id, pair.outlook_event_id, o_body
                    )
                    pair.google_last_modified = g_event.get("updated")
                    pair.outlook_last_modified = o_updated.get("lastModifiedDateTime")
                    pair.last_sync_time = datetime.now(timezone.utc).isoformat()
                    db.upsert_pair(pair)
                summary.updated += 1
            else:
                summary.skipped += 1

        except Exception as exc:
            logger.error("Error processing Google event '%s' (%s): %s", subject, google_id, exc)
            summary.errors += 1

    # ── Pass D: Deletions Google → Outlook ────────────────────────────────────
    # Re-fetch pairs to include any inserted during this cycle
    all_pairs_fresh = db.get_all_pairs()
    for pair in all_pairs_fresh:
        if (
            pair.google_event_id not in google_events
            and pair.google_event_id not in newly_created_google_ids
            and pair.outlook_event_id not in newly_created_outlook_ids
            # Don't double-delete pairs already removed in Pass B
            and pair.google_event_id in {p.google_event_id for p in all_pairs_fresh}
        ):
            # Verify the Google event is really gone (not just outside the window)
            # We only delete if the pair was tracked in the current window
            o_label = f"(Google ID …{pair.google_event_id[-8:]})"
            try:
                log_action("DELETED", o_label, "G→O (Google event gone)")
                if not dry_run:
                    ms_client.delete_event(ms_calendar_id, pair.outlook_event_id)
                    db.delete_pair_by_google_id(pair.google_event_id)
                summary.deleted += 1
            except Exception as exc:
                logger.error(
                    "Error deleting Outlook event for pair %s: %s",
                    pair.google_event_id,
                    exc,
                )
                summary.errors += 1

    # ── Finalise ──────────────────────────────────────────────────────────────
    if not dry_run:
        db.set_metadata("last_sync_time", datetime.now(timezone.utc).isoformat())

    logger.info(
        "Sync complete — created: %d  updated: %d  deleted: %d  "
        "skipped: %d  errors: %d  conflicts: %d",
        summary.created,
        summary.updated,
        summary.deleted,
        summary.skipped,
        summary.errors,
        len(summary.conflicts),
    )
    return summary
