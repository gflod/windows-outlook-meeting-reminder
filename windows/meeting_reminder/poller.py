"""Alert decisions (pure) and the background polling worker."""

from __future__ import annotations

import logging
import math
from collections import OrderedDict
from collections.abc import Callable, Iterable
from datetime import UTC, datetime

import pythoncom
import pywintypes
from PySide6.QtCore import QObject, QTimer, Signal, Slot

from meeting_reminder.models import (
    ALERT_MINUTES_BEFORE,
    NOTIFIED_CAP,
    POLL_INTERVAL_S,
    CalendarEvent,
    CalendarSource,
)
from meeting_reminder.outlook_source import OutlookNotRunning

log = logging.getLogger(__name__)

STATUS_CONNECTED = "Outlook connected"
STATUS_NOT_RUNNING = "Outlook not running — open Outlook"


class AlertTracker:
    """Decides which events should alert now, once each.

    An event is due when it starts between ``minutes_before - 1`` and
    ``minutes_before + 1`` minutes from now; the two-minute window guarantees
    at least one hit with a 30 s poll.
    """

    def __init__(
        self, minutes_before: float = ALERT_MINUTES_BEFORE, cap: int = NOTIFIED_CAP
    ) -> None:
        self._low = minutes_before - 1
        self._high = minutes_before + 1
        self._cap = cap
        self._notified: OrderedDict[str, None] = OrderedDict()

    def due(
        self, events: Iterable[CalendarEvent], now: datetime
    ) -> list[tuple[CalendarEvent, int]]:
        """Return newly due events with whole minutes until start (rounded up)."""
        fired: list[tuple[CalendarEvent, int]] = []
        for event in events:
            minutes = (event.start - now).total_seconds() / 60
            if not self._low <= minutes <= self._high or event.id in self._notified:
                continue
            self._notified[event.id] = None
            if len(self._notified) > self._cap:
                self._notified.popitem(last=False)
            fired.append((event, math.ceil(minutes)))
        return fired


class PollWorker(QObject):
    """Polls a calendar source on its own thread and emits due events.

    The source is built inside the worker thread because COM objects are bound
    to the thread that created them.
    """

    due = Signal(object, int)
    status = Signal(str)

    def __init__(
        self,
        source_factory: Callable[[], CalendarSource],
        interval_s: int = POLL_INTERVAL_S,
    ) -> None:
        super().__init__()
        self._source_factory = source_factory
        self._interval_ms = interval_s * 1000
        self._tracker = AlertTracker()
        self._source: CalendarSource | None = None
        self._timer: QTimer | None = None
        self._last_status = ""

    @Slot()
    def start(self) -> None:
        """Initialise COM on this thread, poll now, then on a timer."""
        pythoncom.CoInitialize()
        self._source = self._source_factory()
        self._timer = QTimer(self)
        self._timer.setInterval(self._interval_ms)
        self._timer.timeout.connect(self.poll)
        self._timer.start()
        self.poll()

    @Slot()
    def stop(self) -> None:
        """Stop polling and release COM. Call on the worker thread."""
        if self._timer is not None:
            self._timer.stop()
        self._source = None
        pythoncom.CoUninitialize()

    @Slot()
    def poll(self) -> None:
        """Fetch upcoming events once and emit any that are due."""
        if self._source is None:
            return
        now = datetime.now(UTC)
        try:
            events = self._source.fetch_upcoming(now)
        except OutlookNotRunning:
            self._set_status(STATUS_NOT_RUNNING)
            return
        except pywintypes.com_error as exc:
            log.warning("Outlook read failed: %s", exc)
            self._set_status(f"Error: {exc.strerror or 'Outlook read failed'}")
            return
        self._set_status(STATUS_CONNECTED)
        log.debug("poll: %d event(s) in window", len(events))
        for event, minutes in self._tracker.due(events, now):
            log.info("Alert: event starting in %d min", minutes)
            self.due.emit(event, minutes)

    def _set_status(self, text: str) -> None:
        if text != self._last_status:
            self._last_status = text
            log.info("Status: %s", text)
            self.status.emit(text)
