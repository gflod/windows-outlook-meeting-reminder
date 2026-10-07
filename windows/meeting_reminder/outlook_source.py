"""Read-only access to the classic Outlook default calendar over COM.

The new Outlook (``olk.exe``) has no COM object model; classic Outlook must be
running. Every COM call here must happen on one thread that has called
``pythoncom.CoInitialize`` (see ``poller.PollWorker``).
"""

from __future__ import annotations

import logging
from datetime import UTC, datetime, timedelta
from typing import Any

import pywintypes
import win32com.client

from meeting_reminder.models import LOOKAHEAD_MIN, CalendarEvent
from meeting_reminder.titles import banner_title

log = logging.getLogger(__name__)

OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_CLASS = 26
OL_RECIPIENT_RESOURCE = 3
#: olMeetingCanceled, olMeetingReceivedAndCanceled
CANCELLED_STATUSES = frozenset({5, 7})
#: Reading Recipients can trigger Outlook's security prompt on machines without
#: recognised antivirus. Set False to use subjects only.
USE_ATTENDEE_NAMES = True
#: Look slightly into the past so an event whose start just passed the poll
#: boundary is still seen.
LOOKBEHIND = timedelta(minutes=1)

_DASL_FMT = "%Y-%m-%d %H:%M"


class OutlookNotRunning(RuntimeError):
    """Classic Outlook is not running, so there is nothing to attach to."""


def dasl_window_query(start_utc: datetime, end_utc: datetime) -> str:
    """Build a DASL filter on appointment start, in UTC.

    DASL is used instead of ``[Start] >= '...'`` because ``Restrict`` parses
    that form using the Windows regional date format and silently matches
    nothing on en-AU machines.
    """
    lo = start_utc.astimezone(UTC).strftime(_DASL_FMT)
    hi = end_utc.astimezone(UTC).strftime(_DASL_FMT)
    field = '"urn:schemas:calendar:dtstart"'
    return f"@SQL={field} >= '{lo}' AND {field} <= '{hi}'"


def _as_utc(value: datetime) -> datetime:
    """Label a COM ``*UTC`` date as UTC (pywin32 may attach a tzinfo or not)."""
    return value.replace(tzinfo=UTC)


def attendee_names(item: Any, current_user: str) -> list[str]:
    """Display names of people on the item, excluding the user and rooms."""
    me = current_user.casefold()
    names: list[str] = []
    for recipient in item.Recipients:
        if recipient.Type == OL_RECIPIENT_RESOURCE:
            continue
        name = (recipient.Name or "").strip()
        if name and name.casefold() != me and name not in names:
            names.append(name)
    return names


def map_appointment(
    item: Any, current_user: str, use_attendee_names: bool = USE_ATTENDEE_NAMES
) -> CalendarEvent | None:
    """Convert an Outlook AppointmentItem (or stand-in) to a CalendarEvent.

    Returns:
        None for all-day and cancelled items, which never alert.
    """
    if item.AllDayEvent or item.MeetingStatus in CANCELLED_STATUSES:
        return None
    start = _as_utc(item.StartUTC)
    end = _as_utc(item.EndUTC)
    base_id = item.GlobalAppointmentID or item.EntryID
    names = attendee_names(item, current_user) if use_attendee_names else []
    return CalendarEvent(
        id=f"{base_id}|{start.isoformat()}",
        title=banner_title(item.Subject, names),
        start=start,
        end=end,
    )


class OutlookCalendarSource:
    """Lists upcoming events from the running classic Outlook."""

    def __init__(self, use_attendee_names: bool = USE_ATTENDEE_NAMES) -> None:
        self._use_attendee_names = use_attendee_names
        self._app: Any = None

    def _attach(self) -> Any:
        if self._app is None:
            try:
                # GetActiveObject, not Dispatch: Dispatch would launch a
                # hidden Outlook if none is running.
                self._app = win32com.client.GetActiveObject("Outlook.Application")
            except pywintypes.com_error as exc:
                raise OutlookNotRunning("Classic Outlook is not running") from exc
        return self._app

    def fetch_upcoming(self, now: datetime) -> list[CalendarEvent]:
        """Return non-all-day, non-cancelled events starting in the window.

        Raises:
            OutlookNotRunning: Outlook could not be attached to.
            pywintypes.com_error: Outlook failed mid-read; the next call
                re-attaches.
        """
        lo, hi = now - LOOKBEHIND, now + timedelta(minutes=LOOKAHEAD_MIN)
        app = self._attach()
        try:
            session = app.Session
            current_user = session.CurrentUser.Name or ""
            items = session.GetDefaultFolder(OL_FOLDER_CALENDAR).Items
            # Order matters: recurrences only expand on a list sorted by Start.
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            matches = items.Restrict(dasl_window_query(lo, hi))
            events: list[CalendarEvent] = []
            # Count is meaningless with IncludeRecurrences; walk the list.
            item = matches.GetFirst()
            while item is not None:
                if getattr(item, "Class", None) == OL_APPOINTMENT_CLASS:
                    event = map_appointment(
                        item, current_user, self._use_attendee_names
                    )
                    if event is not None and lo <= event.start <= hi:
                        events.append(event)
                item = matches.GetNext()
        except pywintypes.com_error:
            self._app = None
            raise
        return events
