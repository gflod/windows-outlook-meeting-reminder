"""Artwork paths, shared with the macOS project's asset catalogue."""

from __future__ import annotations

from pathlib import Path

ASSETS_DIR = Path(__file__).resolve().parents[2] / "MeetingReminder" / "Assets.xcassets"

AIRPLANE = ASSETS_DIR / "airplane.imageset" / "airplane.png"
BANNER = ASSETS_DIR / "banner.imageset" / "banner.png"
#: The colour app icon reads on both light and dark taskbars; the macOS
#: menu-bar template image is black-only.
TRAY_ICON = ASSETS_DIR / "AppIcon.appiconset" / "icon_64.png"


def check_assets() -> None:
    """Fail fast with a clear message if the artwork is missing.

    Raises:
        FileNotFoundError: Listing every missing file.
    """
    missing = [p for p in (AIRPLANE, BANNER, TRAY_ICON) if not p.is_file()]
    if missing:
        raise FileNotFoundError(
            "MeetingReminder artwork not found: " + ", ".join(map(str, missing))
        )
