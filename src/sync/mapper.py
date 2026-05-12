"""Event field mapping between Outlook (Graph API) and Google Calendar.

Handles:
  - Timezone-aware datetime conversion (pytz)
  - All-day events (date-only format)
  - HTML body stripping (html2text)
  - Teams / online meeting URL preservation
  - Full bidirectional RRULE ↔ Graph recurrence conversion
  - "(No title)" fallback for events with no subject/summary

Google color ID reference
-------------------------
1 Tomato  2 Flamingo  3 Tangerine  4 Banana   5 Sage
6 Basil   7 Peacock   8 Blueberry  9 Lavender 10 Grape  11 Graphite
"""
from __future__ import annotations

import logging
import re
from datetime import datetime
from typing import Any, Optional

import html2text as _html2text_mod
import pytz

logger = logging.getLogger("calendar_sync")

# ── HTML → plain-text converter ───────────────────────────────────────────────
_H2T = _html2text_mod.HTML2Text()
_H2T.ignore_links = False
_H2T.ignore_images = True
_H2T.body_width = 0  # no forced line-wrapping

GOOGLE_COLOR_MAP: dict[int, str] = {
    1: "Tomato",
    2: "Flamingo",
    3: "Tangerine",
    4: "Banana",
    5: "Sage",
    6: "Basil",
    7: "Peacock",
    8: "Blueberry",
    9: "Lavender",
    10: "Grape",
    11: "Graphite",
}

# ── Day-name translation tables ───────────────────────────────────────────────
_GRAPH_DAY_TO_ABBR: dict[str, str] = {
    "sunday": "SU",
    "monday": "MO",
    "tuesday": "TU",
    "wednesday": "WE",
    "thursday": "TH",
    "friday": "FR",
    "saturday": "SA",
}
_ABBR_TO_GRAPH_DAY: dict[str, str] = {v: k for k, v in _GRAPH_DAY_TO_ABBR.items()}

_GRAPH_PATTERN_TO_FREQ: dict[str, str] = {
    "daily": "DAILY",
    "weekly": "WEEKLY",
    "absoluteMonthly": "MONTHLY",
    "relativeMonthly": "MONTHLY",
    "absoluteYearly": "YEARLY",
    "relativeYearly": "YEARLY",
}

_INDEX_TO_GRAPH: dict[int, str] = {
    1: "first",
    2: "second",
    3: "third",
    4: "fourth",
    -1: "last",
}
_GRAPH_TO_INDEX: dict[str, int] = {v: k for k, v in _INDEX_TO_GRAPH.items()}


# ── Datetime helpers ──────────────────────────────────────────────────────────

def _parse_graph_datetime(dt_obj: dict) -> datetime:
    """Convert a Graph dateTime object {dateTime, timeZone} to a UTC-aware datetime."""
    dt_str: str = dt_obj.get("dateTime", "")
    tz_name: str = dt_obj.get("timeZone", "UTC")

    try:
        tz = pytz.timezone(tz_name)
    except pytz.UnknownTimeZoneError:
        tz = pytz.UTC

    # Strip trailing fractional seconds beyond microseconds; remove trailing 'Z'
    dt_str_clean = re.sub(r"(\.\d{6})\d+", r"\1", dt_str.rstrip("Z"))
    try:
        dt = datetime.fromisoformat(dt_str_clean)
    except ValueError:
        # Fallback: strip everything after seconds
        dt = datetime.fromisoformat(dt_str_clean[:19])

    if dt.tzinfo is None:
        dt = tz.localize(dt)
    return dt.astimezone(pytz.UTC)


def _html_to_text(html_content: str) -> str:
    if not html_content:
        return ""
    return _H2T.handle(html_content).strip()


# ── Outlook → Google ──────────────────────────────────────────────────────────

def outlook_to_google(
    outlook_event: dict[str, Any],
    sync_category: str,
    color_id: Optional[int],
) -> dict[str, Any]:
    """Convert a Graph API event dict to a Google Calendar event body."""
    subject = outlook_event.get("subject") or "(No title)"

    # Body
    body_obj = outlook_event.get("body") or {}
    raw_content = body_obj.get("content", "")
    if body_obj.get("contentType", "").lower() == "html":
        description = _html_to_text(raw_content)
    else:
        description = raw_content

    # Append Teams / Zoom meeting URL if present and not already in body
    online_url: Optional[str] = outlook_event.get("onlineMeetingUrl")
    if not online_url and outlook_event.get("isOnlineMeeting"):
        join_info = outlook_event.get("onlineMeeting") or {}
        online_url = join_info.get("joinUrl")
    if online_url and online_url not in description:
        description = (description + f"\n\nJoin meeting: {online_url}").strip()

    # Start / end
    is_all_day: bool = bool(outlook_event.get("isAllDay", False))
    start_obj = outlook_event.get("start") or {}
    end_obj = outlook_event.get("end") or {}

    if is_all_day:
        start_dt = _parse_graph_datetime(start_obj)
        end_dt = _parse_graph_datetime(end_obj)
        g_start: dict[str, str] = {"date": start_dt.strftime("%Y-%m-%d")}
        g_end: dict[str, str] = {"date": end_dt.strftime("%Y-%m-%d")}
    else:
        start_dt = _parse_graph_datetime(start_obj)
        end_dt = _parse_graph_datetime(end_obj)
        tz_name = start_obj.get("timeZone", "UTC")
        try:
            tz = pytz.timezone(tz_name)
            g_start = {"dateTime": start_dt.astimezone(tz).isoformat(), "timeZone": tz_name}
            g_end = {"dateTime": end_dt.astimezone(tz).isoformat(), "timeZone": tz_name}
        except pytz.UnknownTimeZoneError:
            g_start = {"dateTime": start_dt.isoformat(), "timeZone": "UTC"}
            g_end = {"dateTime": end_dt.isoformat(), "timeZone": "UTC"}

    location: str = (outlook_event.get("location") or {}).get("displayName", "")

    google_event: dict[str, Any] = {
        "summary": subject,
        "start": g_start,
        "end": g_end,
    }
    if description:
        google_event["description"] = description
    if location:
        google_event["location"] = location
    if color_id is not None:
        google_event["colorId"] = str(color_id)

    # Recurrence
    recurrence = outlook_event.get("recurrence")
    if recurrence:
        rrule = recurrence_graph_to_rrule(recurrence)
        if rrule:
            google_event["recurrence"] = [rrule]

    return google_event


# ── Google → Outlook ──────────────────────────────────────────────────────────

def google_to_outlook(
    google_event: dict[str, Any],
    sync_category: str,
) -> dict[str, Any]:
    """Convert a Google Calendar event dict to a Graph API event body."""
    summary = google_event.get("summary") or "(No title)"
    description = google_event.get("description", "")
    location = google_event.get("location", "")

    start_obj = google_event.get("start") or {}
    end_obj = google_event.get("end") or {}

    is_all_day = "date" in start_obj and "dateTime" not in start_obj

    if is_all_day:
        start_date = start_obj["date"]
        end_date = end_obj["date"]
        graph_start = {"dateTime": f"{start_date}T00:00:00", "timeZone": "UTC"}
        graph_end = {"dateTime": f"{end_date}T00:00:00", "timeZone": "UTC"}
    else:
        start_dt = datetime.fromisoformat(start_obj["dateTime"])
        end_dt = datetime.fromisoformat(end_obj["dateTime"])
        tz_name = start_obj.get("timeZone") or end_obj.get("timeZone") or "UTC"
        graph_start = {"dateTime": start_dt.isoformat(), "timeZone": tz_name}
        graph_end = {"dateTime": end_dt.isoformat(), "timeZone": tz_name}

    outlook_event: dict[str, Any] = {
        "subject": summary,
        "body": {"contentType": "text", "content": description},
        "start": graph_start,
        "end": graph_end,
        "isAllDay": is_all_day,
        "categories": [sync_category],
    }

    if location:
        outlook_event["location"] = {"displayName": location}

    # Recurrence: find the first RRULE in the list
    for item in google_event.get("recurrence", []):
        if item.startswith("RRULE:"):
            start_date_str = start_obj.get("date") or (
                start_obj.get("dateTime", "")[:10]
            )
            graph_recurrence = recurrence_rrule_to_graph(item, start_date_str)
            if graph_recurrence:
                outlook_event["recurrence"] = graph_recurrence
            break

    return outlook_event


# ── Recurrence: Graph pattern/range → RRULE ───────────────────────────────────

def recurrence_graph_to_rrule(recurrence: dict[str, Any]) -> str:
    """Convert a Graph recurrence object to an RRULE string (e.g. 'RRULE:FREQ=WEEKLY;BYDAY=MO')."""
    pattern: dict = recurrence.get("pattern") or {}
    range_: dict = recurrence.get("range") or {}

    pattern_type: str = pattern.get("type", "")
    freq = _GRAPH_PATTERN_TO_FREQ.get(pattern_type)
    if not freq:
        logger.warning("Unknown Graph recurrence pattern type: %r", pattern_type)
        return ""

    interval: int = int(pattern.get("interval", 1))
    parts = [f"FREQ={freq}"]

    if interval > 1:
        parts.append(f"INTERVAL={interval}")

    # BYDAY
    days_of_week: list[str] = pattern.get("daysOfWeek") or []
    if days_of_week:
        index_word: str = pattern.get("index", "")
        index_num = _GRAPH_TO_INDEX.get(index_word, 0)
        byday_parts: list[str] = []
        for day in days_of_week:
            abbr = _GRAPH_DAY_TO_ABBR.get(day.lower(), day[:2].upper())
            if index_num and pattern_type in ("relativeMonthly", "relativeYearly"):
                byday_parts.append(f"{index_num}{abbr}")
            else:
                byday_parts.append(abbr)
        parts.append(f"BYDAY={','.join(byday_parts)}")

    # BYMONTHDAY (absoluteMonthly / absoluteYearly)
    day_of_month = pattern.get("dayOfMonth")
    if day_of_month and pattern_type in ("absoluteMonthly", "absoluteYearly"):
        parts.append(f"BYMONTHDAY={day_of_month}")

    # BYMONTH (yearly patterns)
    month = pattern.get("month")
    if month and pattern_type in ("absoluteYearly", "relativeYearly"):
        parts.append(f"BYMONTH={month}")

    # Range
    range_type: str = range_.get("type", "")
    if range_type == "endDate":
        end_date = range_.get("endDate", "").replace("-", "")
        if end_date:
            parts.append(f"UNTIL={end_date}T000000Z")
    elif range_type == "numbered":
        count = range_.get("numberOfOccurrences", 0)
        if count:
            parts.append(f"COUNT={count}")
    # "noEnd" → omit UNTIL/COUNT

    return "RRULE:" + ";".join(parts)


# ── Recurrence: RRULE → Graph pattern/range ───────────────────────────────────

def recurrence_rrule_to_graph(
    rrule_str: str, start_date_str: str
) -> Optional[dict[str, Any]]:
    """Convert an RRULE string to a Graph recurrence object.

    Returns None if the RRULE cannot be mapped (e.g. unsupported FREQ).
    """
    rule_body = rrule_str.removeprefix("RRULE:")
    props: dict[str, str] = {}
    for part in rule_body.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            props[k.strip().upper()] = v.strip()

    freq = props.get("FREQ", "")
    interval = int(props.get("INTERVAL", "1"))
    byday = props.get("BYDAY", "")
    bymonthday = props.get("BYMONTHDAY", "")
    bymonth = props.get("BYMONTH", "")
    until = props.get("UNTIL", "")
    count = props.get("COUNT", "")

    # Determine Graph pattern type
    if freq == "DAILY":
        pattern_type = "daily"
    elif freq == "WEEKLY":
        pattern_type = "weekly"
    elif freq == "MONTHLY":
        # relativeMonthly uses a positional BYDAY like "1MO" or "-1FR"
        if byday and re.match(r"^-?\d+[A-Z]{2}", byday.split(",")[0].strip()):
            pattern_type = "relativeMonthly"
        else:
            pattern_type = "absoluteMonthly"
    elif freq == "YEARLY":
        if byday and re.match(r"^-?\d+[A-Z]{2}", byday.split(",")[0].strip()):
            pattern_type = "relativeYearly"
        else:
            pattern_type = "absoluteYearly"
    else:
        logger.warning("Unsupported RRULE FREQ: %r — recurrence will not be synced.", freq)
        return None

    pattern: dict[str, Any] = {"type": pattern_type, "interval": interval}

    # daysOfWeek + optional index
    if byday:
        days: list[str] = []
        index_num = 0
        for part in byday.split(","):
            m = re.match(r"^(-?\d+)?([A-Z]{2})$", part.strip())
            if m:
                if m.group(1):
                    index_num = int(m.group(1))
                abbr = m.group(2)
                days.append(_ABBR_TO_GRAPH_DAY.get(abbr, abbr.lower()))
        pattern["daysOfWeek"] = days
        if index_num and pattern_type in ("relativeMonthly", "relativeYearly"):
            pattern["index"] = _INDEX_TO_GRAPH.get(index_num, "first")

    # dayOfMonth
    if bymonthday:
        try:
            pattern["dayOfMonth"] = int(bymonthday.split(",")[0])
        except ValueError:
            pass

    # month (yearly only)
    if bymonth:
        try:
            pattern["month"] = int(bymonth.split(",")[0])
        except ValueError:
            pass

    # Range
    range_: dict[str, Any] = {"startDate": start_date_str}
    if until:
        # UNTIL is YYYYMMDD or YYYYMMDDTHHmmssZ
        d = until.replace("Z", "")[:8]
        range_["type"] = "endDate"
        range_["endDate"] = f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    elif count:
        range_["type"] = "numbered"
        range_["numberOfOccurrences"] = int(count)
    else:
        range_["type"] = "noEnd"

    return {"pattern": pattern, "range": range_}
