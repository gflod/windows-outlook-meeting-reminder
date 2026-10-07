"""Controller: wires the Outlook poller, overlay and tray together."""

from __future__ import annotations

import logging

from PySide6.QtCore import QMetaObject, QObject, Qt, QThread, Slot
from PySide6.QtWidgets import QApplication

from meeting_reminder import autostart, settings
from meeting_reminder.models import CalendarEvent, Speed
from meeting_reminder.outlook_source import OutlookCalendarSource
from meeting_reminder.overlay import AirplaneOverlay
from meeting_reminder.poller import PollWorker
from meeting_reminder.titles import banner_text
from meeting_reminder.tray import TrayMenu

log = logging.getLogger(__name__)

TEST_TITLE = "Test Meeting"
TEST_MINUTES = 5


class AppController(QObject):
    """Owns every long-lived part of the running app."""

    def __init__(self, app: QApplication) -> None:
        super().__init__()
        self._app = app
        self._speed = settings.load_speed()
        self._overlay = AirplaneOverlay()

        self._tray = TrayMenu(self._speed, autostart.is_enabled())
        self._tray.speed_changed.connect(self._on_speed_changed)
        self._tray.autostart_toggled.connect(self._on_autostart_toggled)
        self._tray.test_requested.connect(self._on_test)
        self._tray.quit_requested.connect(self.quit)

        self._thread = QThread()
        self._thread.setObjectName("outlook-poller")
        self._worker = PollWorker(OutlookCalendarSource)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.start)
        self._worker.due.connect(self._on_due)
        self._worker.status.connect(self._tray.set_status)

    def start(self) -> None:
        """Show the tray icon and begin polling."""
        self._tray.show()
        self._thread.start()

    @Slot()
    def quit(self) -> None:
        """Stop the poller cleanly and exit."""
        QMetaObject.invokeMethod(
            self._worker, "stop", Qt.ConnectionType.BlockingQueuedConnection
        )
        self._thread.quit()
        self._thread.wait()
        self._tray.hide()
        self._app.quit()

    def _on_due(self, event: CalendarEvent, minutes: int) -> None:
        self._overlay.show(
            banner_text(event.title, minutes), self._speed.flight_seconds
        )

    def _on_test(self) -> None:
        self._overlay.show(
            banner_text(TEST_TITLE, TEST_MINUTES), self._speed.flight_seconds
        )

    def _on_speed_changed(self, speed: Speed) -> None:
        self._speed = speed
        settings.save_speed(speed)

    def _on_autostart_toggled(self, enabled: bool) -> None:
        try:
            autostart.set_enabled(enabled)
        except OSError:
            log.exception("Could not update start-at-login")
        # Re-read the real state; the write may have failed.
        self._tray.set_autostart(autostart.is_enabled())
