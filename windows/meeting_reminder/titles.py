"""Banner text formatting (mirrors ``AppleCalendarService.bannerTitle``)."""

from __future__ import annotations

UNTITLED = "Untitled Meeting"


def banner_title(subject: str | None, attendee_names: list[str]) -> str:
    """Prefer "Meeting with <names>" when there are other attendees.

    Args:
        subject: The event subject; may be empty or None.
        attendee_names: Display names of attendees other than the user.

    Returns:
        The title to show on the banner.
    """
    names = [n for n in attendee_names if n]
    if not names:
        return (subject or "").strip() or UNTITLED
    if len(names) == 1:
        return f"Meeting with {names[0]}"
    if len(names) == 2:
        return f"Meeting with {names[0]} and {names[1]}"
    return f"Meeting with {names[0]}, {names[1]} +{len(names) - 2} more"


def banner_text(title: str, minutes_until: int) -> str:
    """Full banner string, e.g. "Standup in 5 min"."""
    return f"{title} in {minutes_until} min"
