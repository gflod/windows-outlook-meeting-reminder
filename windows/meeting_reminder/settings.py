"""Persisted user preferences."""

from __future__ import annotations

from PySide6.QtCore import QSettings

from meeting_reminder.models import Speed

_SPEED_KEY = "speed"


def _settings() -> QSettings:
    return QSettings("MeetingReminder", "MeetingReminder")


def load_speed() -> Speed:
    """Saved plane speed, defaulting to normal."""
    try:
        return Speed(_settings().value(_SPEED_KEY, Speed.NORMAL.value))
    except ValueError:
        return Speed.NORMAL


def save_speed(speed: Speed) -> None:
    """Persist the plane speed."""
    _settings().setValue(_SPEED_KEY, speed.value)
