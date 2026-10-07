"""Shared value types and tunable constants."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from typing import Protocol

#: Minutes before a meeting's start at which the airplane flies.
ALERT_MINUTES_BEFORE = 5
#: Seconds between calendar polls.
POLL_INTERVAL_S = 30
#: How far ahead each poll looks, in minutes.
LOOKAHEAD_MIN = 60
#: Maximum number of already-alerted event ids remembered.
NOTIFIED_CAP = 200


@dataclass(frozen=True)
class CalendarEvent:
    """One calendar occurrence, with timezone-aware UTC start and end.

    Attributes:
        id: Unique per occurrence (recurring meetings share an Outlook id, so
            the start time is part of this key).
        title: Text shown on the banner, before " in N min".
        start: Occurrence start, aware UTC.
        end: Occurrence end, aware UTC.
    """

    id: str
    title: str
    start: datetime
    end: datetime


class Speed(str, Enum):
    """Preset plane speeds; the value is the setting key."""

    SLOW = "slow"
    NORMAL = "normal"
    FAST = "fast"

    @property
    def flight_seconds(self) -> float:
        """Seconds the plane takes to cross the screen."""
        return {Speed.SLOW: 22.0, Speed.NORMAL: 14.0, Speed.FAST: 8.0}[self]

    @property
    def label(self) -> str:
        """Menu label."""
        return self.value.capitalize()


class CalendarSource(Protocol):
    """Anything that can list upcoming events."""

    def fetch_upcoming(self, now: datetime) -> list[CalendarEvent]:
        """Return events starting within the lookahead window from ``now``."""
        ...
