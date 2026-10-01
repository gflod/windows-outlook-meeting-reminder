import pytest

from meeting_reminder.titles import UNTITLED, banner_text, banner_title


@pytest.mark.parametrize(
    ("subject", "names", "expected"),
    [
        ("Standup", [], "Standup"),
        ("  Standup ", [], "Standup"),
        (None, [], UNTITLED),
        ("", [], UNTITLED),
        ("Sync", ["Alex"], "Meeting with Alex"),
        ("Sync", ["Alex", "Sam"], "Meeting with Alex and Sam"),
        ("Sync", ["Alex", "Sam", "Kim"], "Meeting with Alex, Sam +1 more"),
        ("Sync", ["Alex", "Sam", "Kim", "Lee"], "Meeting with Alex, Sam +2 more"),
        ("Sync", ["", ""], "Sync"),
    ],
)
def test_banner_title(subject, names, expected):
    assert banner_title(subject, names) == expected


def test_banner_text():
    assert banner_text("Standup", 5) == "Standup in 5 min"
