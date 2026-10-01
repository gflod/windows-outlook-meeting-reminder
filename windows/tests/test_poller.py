from datetime import UTC, datetime, timedelta

import pytest

from meeting_reminder.models import CalendarEvent
from meeting_reminder.poller import AlertTracker

NOW = datetime(2026, 10, 1, 0, 0, tzinfo=UTC)


def event(minutes_away: float, id_: str = "e1") -> CalendarEvent:
    start = NOW + timedelta(minutes=minutes_away)
    return CalendarEvent(id_, "T", start, start + timedelta(minutes=30))


@pytest.mark.parametrize(
    ("minutes", "fires"),
    [(3.99, False), (4.0, True), (5.0, True), (6.0, True), (6.01, False), (-1, False)],
)
def test_window_edges(minutes, fires):
    assert bool(AlertTracker().due([event(minutes)], NOW)) is fires


@pytest.mark.parametrize(("minutes", "shown"), [(4.1, 5), (5.0, 5), (5.5, 6)])
def test_minutes_round_up(minutes, shown):
    [(_, got)] = AlertTracker().due([event(minutes)], NOW)
    assert got == shown


def test_fires_once_across_polls():
    tracker = AlertTracker()
    assert len(tracker.due([event(5.5)], NOW)) == 1
    later = NOW + timedelta(seconds=30)
    assert tracker.due([event(5.5)], later) == []


def test_recurring_occurrences_fire_separately():
    # Same Outlook id, different occurrence starts -> different event ids.
    tracker = AlertTracker()
    first = event(5, "gaid|2026-10-01T00:05:00+00:00")
    assert len(tracker.due([first], NOW)) == 1
    next_day = NOW + timedelta(days=1)
    second = CalendarEvent(
        "gaid|2026-10-02T00:05:00+00:00",
        "T",
        next_day + timedelta(minutes=5),
        next_day + timedelta(minutes=35),
    )
    assert len(tracker.due([second], next_day)) == 1


def test_eviction_at_cap():
    tracker = AlertTracker(cap=2)
    for i in range(3):
        tracker.due([event(5, f"e{i}")], NOW)
    # e0 was evicted, so it can fire again; e2 is still remembered.
    assert len(tracker.due([event(5, "e0")], NOW)) == 1
    assert tracker.due([event(5, "e2")], NOW) == []
