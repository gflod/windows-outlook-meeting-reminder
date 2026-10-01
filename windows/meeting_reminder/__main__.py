"""Entry point: ``python -m meeting_reminder [--list] [--debug]``."""

from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import UTC, datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

log = logging.getLogger("meeting_reminder")

MUTEX_NAME = "Local\\MeetingReminder.SingleInstance"


def _log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    path = Path(base) / "MeetingReminder"
    path.mkdir(parents=True, exist_ok=True)
    return path


def configure_logging(debug: bool) -> None:
    """Rotating file log, plus stderr when there is a console."""
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    handlers: list[logging.Handler] = [
        RotatingFileHandler(
            _log_dir() / "meeting_reminder.log",
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
    ]
    if sys.stderr is not None:  # None under pythonw.exe
        handlers.append(logging.StreamHandler())
    for handler in handlers:
        handler.setFormatter(fmt)
        log.addHandler(handler)
    log.setLevel(logging.DEBUG if debug else logging.INFO)


def list_events() -> int:
    """Print the next hour's events from Outlook (read-only) and exit."""
    from meeting_reminder.outlook_source import OutlookCalendarSource, OutlookNotRunning

    try:
        events = OutlookCalendarSource().fetch_upcoming(datetime.now(UTC))
    except OutlookNotRunning:
        sys.stderr.write("Classic Outlook is not running.\n")
        return 1
    for event in events:
        local = event.start.astimezone().strftime("%d %b %Y %H:%M")
        sys.stdout.write(f"{local}  {event.title}\n")
    sys.stdout.write(f"{len(events)} event(s) in the next hour\n")
    return 0


def run_app() -> int:
    """Start the tray app (single instance)."""
    import win32api
    import win32event
    import winerror
    from PySide6.QtWidgets import QApplication, QSystemTrayIcon

    from meeting_reminder.app import AppController
    from meeting_reminder.assets import check_assets

    mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("Another MeetingReminder is already running; exiting")
        return 0

    check_assets()
    app = QApplication(sys.argv)
    app.setApplicationName("MeetingReminder")
    app.setQuitOnLastWindowClosed(False)
    if not QSystemTrayIcon.isSystemTrayAvailable():
        log.error("No system tray available")
        return 1

    controller = AppController(app)
    controller.start()
    log.info("MeetingReminder started")
    code = app.exec()
    del mutex
    return code


def main(argv: list[str] | None = None) -> int:
    """Parse arguments and dispatch."""
    parser = argparse.ArgumentParser(prog="meeting_reminder")
    parser.add_argument(
        "--list", action="store_true", help="print the next hour's events and exit"
    )
    parser.add_argument("--debug", action="store_true", help="verbose logging")
    args = parser.parse_args(argv)
    configure_logging(args.debug)
    return list_events() if args.list else run_app()


if __name__ == "__main__":
    sys.exit(main())
