"""Google Calendar API client.

Events are fetched with singleEvents=True so recurring instances are expanded
into individual items (consistent with how Graph calendarView works).

All mutating operations and the event list call are wrapped with the retry
decorator to handle transient 429/5xx responses.
"""
from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any, Optional

from google.oauth2.credentials import Credentials  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.errors import HttpError  # type: ignore

from src.utils.retry import with_retry

logger = logging.getLogger("calendar_sync")

_PAGE_SIZE = 250


class GoogleCalendarClient:
    """Thin wrapper around the Google Calendar v3 REST API."""

    def __init__(self, credentials: Credentials) -> None:
        self._service = build(
            "calendar", "v3", credentials=credentials, cache_discovery=False
        )

    # ── Calendars ─────────────────────────────────────────────────────────────

    def list_calendars(self) -> list[dict]:
        """Return all calendars accessible to the authenticated user."""
        calendars: list[dict] = []
        page_token: Optional[str] = None
        while True:
            resp = (
                self._service.calendarList()
                .list(maxResults=250, pageToken=page_token)
                .execute()
            )
            calendars.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return calendars

    # ── Events ────────────────────────────────────────────────────────────────

    @with_retry(max_attempts=3)
    def get_events(
        self,
        calendar_id: str,
        start: datetime,
        end: datetime,
    ) -> list[dict]:
        """Fetch all event instances within [start, end].

        singleEvents=True expands recurring events into individual instances,
        which mirrors the behaviour of Graph calendarView.
        """
        events: list[dict] = []
        page_token: Optional[str] = None
        time_min = start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        time_max = end.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        while True:
            resp = (
                self._service.events()
                .list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=True,
                    orderBy="startTime",
                    maxResults=_PAGE_SIZE,
                    pageToken=page_token,
                    showDeleted=False,
                )
                .execute()
            )
            events.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return events

    @with_retry(max_attempts=3)
    def create_event(self, calendar_id: str, event_body: dict) -> dict:
        return (
            self._service.events()
            .insert(calendarId=calendar_id, body=event_body)
            .execute()
        )

    @with_retry(max_attempts=3)
    def update_event(self, calendar_id: str, event_id: str, event_body: dict) -> dict:
        return (
            self._service.events()
            .update(calendarId=calendar_id, eventId=event_id, body=event_body)
            .execute()
        )

    @with_retry(max_attempts=3)
    def delete_event(self, calendar_id: str, event_id: str) -> None:
        try:
            self._service.events().delete(
                calendarId=calendar_id, eventId=event_id
            ).execute()
        except HttpError as exc:
            # 410 Gone means the event is already deleted — treat as success
            if int(exc.resp.status) == 410:
                logger.debug("Google event %s already deleted (410 Gone).", event_id)
                return
            raise
