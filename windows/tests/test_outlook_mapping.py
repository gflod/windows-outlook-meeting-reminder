from datetime import UTC, datetime
from types import SimpleNamespace

import pytest

from meeting_reminder.outlook_source import (
    OL_RECIPIENT_RESOURCE,
    dasl_window_query,
    map_appointment,
)

ME = "Georgina Doyle"


def recipient(name: str, type_: int = 1) -> SimpleNamespace:
    return SimpleNamespace(Name=name, Type=type_)


def appt(**overrides) -> SimpleNamespace:
    fields = dict(
        AllDayEvent=False,
        MeetingStatus=1,
        # pywin32 hands back naive or UTC-tagged datetimes; both must map to UTC.
        StartUTC=datetime(2026, 10, 1, 0, 5),
        EndUTC=datetime(2026, 10, 1, 0, 35),
        GlobalAppointmentID="GAID",
        EntryID="ENTRY",
        Subject="Standup",
        Recipients=[],
    )
    fields.update(overrides)
    return SimpleNamespace(**fields)


def test_maps_basic_fields():
    ev = map_appointment(appt(), ME)
    assert ev.title == "Standup"
    assert ev.start == datetime(2026, 10, 1, 0, 5, tzinfo=UTC)
    assert ev.end.tzinfo is UTC
    assert ev.id == "GAID|2026-10-01T00:05:00+00:00"


def test_entry_id_fallback():
    assert map_appointment(appt(GlobalAppointmentID=""), ME).id.startswith("ENTRY|")


@pytest.mark.parametrize(
    "overrides", [{"AllDayEvent": True}, {"MeetingStatus": 5}, {"MeetingStatus": 7}]
)
def test_skips_all_day_and_cancelled(overrides):
    assert map_appointment(appt(**overrides), ME) is None


def test_attendees_exclude_self_rooms_and_duplicates():
    recips = [
        recipient(ME),
        recipient("georgina doyle"),
        recipient("Boardroom", OL_RECIPIENT_RESOURCE),
        recipient("Alex"),
        recipient("Alex"),
        recipient("Sam", 2),
    ]
    assert map_appointment(appt(Recipients=recips), ME).title == (
        "Meeting with Alex and Sam"
    )


def test_attendee_names_can_be_disabled():
    ev = map_appointment(appt(Recipients=[recipient("Alex")]), ME, False)
    assert ev.title == "Standup"


def test_dasl_query_is_utc_and_iso_formatted():
    lo = datetime(2026, 10, 1, 10, 0, tzinfo=UTC)
    hi = datetime(2026, 10, 1, 11, 0, tzinfo=UTC)
    q = dasl_window_query(lo, hi)
    assert q == (
        "@SQL=\"urn:schemas:calendar:dtstart\" >= '2026-10-01 10:00' AND "
        "\"urn:schemas:calendar:dtstart\" <= '2026-10-01 11:00'"
    )
