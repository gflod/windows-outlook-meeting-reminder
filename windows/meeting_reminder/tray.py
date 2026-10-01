"""System tray icon and menu. Emits user intents; holds no app logic."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import QMenu, QSystemTrayIcon

from meeting_reminder import assets
from meeting_reminder.models import Speed


class TrayMenu(QObject):
    """Tray icon with status, speed, start-at-login, test and quit."""

    speed_changed = Signal(object)
    autostart_toggled = Signal(bool)
    test_requested = Signal()
    quit_requested = Signal()

    def __init__(self, speed: Speed, autostart: bool) -> None:
        super().__init__()
        self._menu = QMenu()

        self._status = self._menu.addAction("Starting…")
        self._status.setEnabled(False)
        self._menu.addSeparator()

        speed_menu = self._menu.addMenu("Plane speed")
        group = QActionGroup(speed_menu)
        group.setExclusive(True)
        self._speed_actions: dict[Speed, QAction] = {}
        for option in Speed:
            action = speed_menu.addAction(option.label)
            action.setCheckable(True)
            action.setChecked(option is speed)
            action.triggered.connect(
                lambda _=False, s=option: self.speed_changed.emit(s)
            )
            group.addAction(action)
            self._speed_actions[option] = action

        self._autostart = self._menu.addAction("Start at login")
        self._autostart.setCheckable(True)
        self._autostart.setChecked(autostart)
        self._autostart.triggered.connect(self.autostart_toggled.emit)
        self._menu.addSeparator()

        self._menu.addAction("Test airplane").triggered.connect(self.test_requested)
        self._menu.addSeparator()
        self._menu.addAction("Quit MeetingReminder").triggered.connect(
            self.quit_requested
        )

        self._icon = QSystemTrayIcon(QIcon(str(assets.TRAY_ICON)))
        self._icon.setToolTip("MeetingReminder")
        self._icon.setContextMenu(self._menu)
        self._icon.activated.connect(self._on_activated)

    def show(self) -> None:
        """Put the icon in the notification area."""
        self._icon.show()

    def hide(self) -> None:
        """Remove the icon."""
        self._icon.hide()

    def set_status(self, text: str) -> None:
        """Update the status line and tooltip."""
        self._status.setText(text)
        self._icon.setToolTip(f"MeetingReminder — {text}")

    def set_autostart(self, enabled: bool) -> None:
        """Reflect the real start-at-login state."""
        self._autostart.setChecked(enabled)

    def _on_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        # Left-click opens the menu too, like the macOS menu-bar item.
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._menu.popup(self._icon.geometry().center())
