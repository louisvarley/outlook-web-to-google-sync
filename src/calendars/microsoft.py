"""Microsoft Graph API calendar client.

All mutating operations (create/update/delete) and reads are wrapped with the
retry decorator so transient 429/5xx responses are handled transparently.

Important: event queries use the /calendarView endpoint (not /me/events) because
only calendarView correctly expands recurring event instances within a date range.
"""
from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any, Callable, Generator, Optional

import requests

from src.utils.retry import with_retry

logger = logging.getLogger("calendar_sync")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_PAGE_SIZE = 100

# Fields fetched for each event — keep the $select tight to minimise payload
_EVENT_SELECT = (
    "id,subject,body,start,end,location,categories,isAllDay,"
    "recurrence,isOnlineMeeting,onlineMeetingUrl,onlineMeeting,"
    "lastModifiedDateTime,organizer,attendees,isCancelled,"
    "type,seriesMasterId,originalStart"
)


class GraphClient:
    """Thin wrapper around the Microsoft Graph calendar REST API."""

    def __init__(self, token_getter: Callable[[], str]) -> None:
        self._get_token = token_getter

    # ── HTTP helpers ──────────────────────────────────────────────────────────

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
            # Ask Graph to return datetimes in UTC
            "Prefer": 'outlook.timezone="UTC"',
        }

    @with_retry(max_attempts=3)
    def _get(self, url: str, params: Optional[dict] = None) -> dict:
        resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
        if not resp.ok:
            raise requests.HTTPError(
                f"{resp.status_code} Client Error: {resp.reason} for url: {resp.url} — {resp.text}",
                response=resp,
            )
        return resp.json()

    @with_retry(max_attempts=3)
    def _post(self, url: str, body: dict) -> dict:
        resp = requests.post(url, headers=self._headers(), json=body, timeout=30)
        resp.raise_for_status()
        return resp.json()

    @with_retry(max_attempts=3)
    def _patch(self, url: str, body: dict) -> dict:
        resp = requests.patch(url, headers=self._headers(), json=body, timeout=30)
        resp.raise_for_status()
        return resp.json()

    @with_retry(max_attempts=3)
    def _delete(self, url: str) -> None:
        resp = requests.delete(url, headers=self._headers(), timeout=30)
        if resp.status_code == 404:
            logger.debug("DELETE %s returned 404 (already gone).", url)
            return
        resp.raise_for_status()

    def _paginate(
        self,
        url: str,
        params: Optional[dict] = None,
    ) -> Generator[dict, None, None]:
        """Yield all items across paginated Graph responses (handles @odata.nextLink)."""
        next_url: Optional[str] = url
        current_params = params
        while next_url:
            data = self._get(next_url, params=current_params)
            # Params apply only to the first request; subsequent pages use nextLink
            current_params = None
            for item in data.get("value", []):
                yield item
            next_url = data.get("@odata.nextLink")

    # ── Calendars ─────────────────────────────────────────────────────────────

    def list_calendars(self) -> list[dict]:
        return list(self._paginate(f"{_GRAPH_BASE}/me/calendars", {"$top": _PAGE_SIZE}))

    # ── Events ────────────────────────────────────────────────────────────────

    def get_events(
        self,
        calendar_id: str,
        start: datetime,
        end: datetime,
    ) -> list[dict]:
        """Fetch all events within [start, end] using calendarView.

        calendarView must be used (not /events) to correctly expand recurring
        event instances into individual occurrences within the date range.
        """
        url = f"{_GRAPH_BASE}/me/calendars/{calendar_id}/calendarView"
        params = {
            "startDateTime": start.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "endDateTime": end.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "$top": _PAGE_SIZE,
            "$select": _EVENT_SELECT,
        }
        return list(self._paginate(url, params))

    def create_event(self, calendar_id: str, event_body: dict) -> dict:
        return self._post(f"{_GRAPH_BASE}/me/calendars/{calendar_id}/events", event_body)

    def update_event(self, calendar_id: str, event_id: str, event_body: dict) -> dict:
        return self._patch(
            f"{_GRAPH_BASE}/me/calendars/{calendar_id}/events/{event_id}",
            event_body,
        )

    def delete_event(self, calendar_id: str, event_id: str) -> None:
        self._delete(f"{_GRAPH_BASE}/me/calendars/{calendar_id}/events/{event_id}")

    # ── Categories ────────────────────────────────────────────────────────────

    def get_master_categories(self) -> list[dict]:
        """Fetch the user's Outlook master category list."""
        try:
            data = self._get(f"{_GRAPH_BASE}/me/outlook/masterCategories")
            return data.get("value", [])
        except Exception as exc:
            logger.warning("Could not fetch master categories: %s", exc)
            return []

    def get_event_categories(
        self,
        calendar_id: str,
        start: datetime,
        end: datetime,
    ) -> list[str]:
        """Return all unique category names used in events within the date range."""
        events = self.get_events(calendar_id, start, end)
        seen: set[str] = set()
        for event in events:
            for cat in event.get("categories", []):
                seen.add(cat)
        return sorted(seen, key=str.casefold)
