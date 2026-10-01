"""The flying airplane: one transparent, click-through window per screen."""

from __future__ import annotations

import logging

import win32con
import win32gui
from PySide6.QtCore import (
    QEasingCurve,
    QPoint,
    QPropertyAnimation,
    QRect,
    QRectF,
    Qt,
    QTimer,
)
from PySide6.QtGui import (
    QColor,
    QFont,
    QFontMetrics,
    QGuiApplication,
    QPainter,
    QPixmap,
    QScreen,
)
from PySide6.QtWidgets import QLabel, QWidget

from meeting_reminder import assets

log = logging.getLogger(__name__)

FONT_FAMILY = "Quicksand"
FONT_WEIGHT = QFont.Weight.Light
FONT_PIXEL_SIZE = 28
TEXT_PAD_X = 50
TEXT_PAD_Y = 22
PLANE_SIZE = 220
#: Negative spacing: the banner overlaps the plane so the rope tucks under.
PLANE_OVERLAP = 10
#: Matches the macOS 110 pt strip placed 65 % up the screen; the content is
#: centred on that strip's midline.
STRIP_HEIGHT = 110
STRIP_TOP_FRACTION = 0.35
FADE_S = 0.6
CLOSE_AFTER_EXTRA_S = 1.5

_CLICK_THROUGH_EX = (
    win32con.WS_EX_TRANSPARENT
    | win32con.WS_EX_LAYERED
    | win32con.WS_EX_NOACTIVATE
    | win32con.WS_EX_TOOLWINDOW
)


def render_banner(text: str, device_pixel_ratio: float) -> QPixmap:
    """Draw banner + plane once into a transparent pixmap."""
    font = QFont(FONT_FAMILY)
    font.setWeight(FONT_WEIGHT)
    font.setPixelSize(FONT_PIXEL_SIZE)
    metrics = QFontMetrics(font)
    banner_w = metrics.horizontalAdvance(text) + 2 * TEXT_PAD_X
    banner_h = metrics.height() + 2 * TEXT_PAD_Y
    width = banner_w - PLANE_OVERLAP + PLANE_SIZE
    height = max(banner_h, PLANE_SIZE)

    pixmap = QPixmap(
        round(width * device_pixel_ratio), round(height * device_pixel_ratio)
    )
    pixmap.setDevicePixelRatio(device_pixel_ratio)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHints(
        QPainter.RenderHint.Antialiasing
        | QPainter.RenderHint.SmoothPixmapTransform
        | QPainter.RenderHint.TextAntialiasing
    )
    # Plane first so the banner draws over it.
    plane = QPixmap(str(assets.AIRPLANE))
    scaled = plane.size().scaled(
        PLANE_SIZE, PLANE_SIZE, Qt.AspectRatioMode.KeepAspectRatio
    )
    plane_rect = QRectF(
        banner_w - PLANE_OVERLAP + (PLANE_SIZE - scaled.width()) / 2,
        (height - scaled.height()) / 2,
        scaled.width(),
        scaled.height(),
    )
    painter.drawPixmap(plane_rect, plane, QRectF(plane.rect()))

    banner_rect = QRectF(0, (height - banner_h) / 2, banner_w, banner_h)
    banner = QPixmap(str(assets.BANNER))
    painter.drawPixmap(banner_rect, banner, QRectF(banner.rect()))

    painter.setFont(font)
    painter.setPen(QColor("white"))
    painter.drawText(banner_rect, Qt.AlignmentFlag.AlignCenter, text)
    painter.end()
    return pixmap


class AirplaneWindow(QWidget):
    """A screen-wide, see-through strip that one airplane flies across."""

    def __init__(self, screen: QScreen, text: str, duration_s: float) -> None:
        super().__init__(
            None,
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
            | Qt.WindowType.WindowDoesNotAcceptFocus,
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self._duration_s = duration_s

        pixmap = render_banner(text, screen.devicePixelRatio())
        content_w = round(pixmap.width() / pixmap.devicePixelRatio())
        content_h = round(pixmap.height() / pixmap.devicePixelRatio())

        geo = screen.geometry()
        mid_y = geo.y() + round(geo.height() * STRIP_TOP_FRACTION) + STRIP_HEIGHT // 2
        self.setGeometry(QRect(geo.x(), mid_y - content_h // 2, geo.width(), content_h))

        self._label = QLabel(self)
        self._label.setPixmap(pixmap)
        self._label.resize(content_w, content_h)
        self._label.move(-content_w, 0)

        self._fly = QPropertyAnimation(self._label, b"pos", self)
        self._fly.setStartValue(QPoint(-content_w, 0))
        self._fly.setEndValue(QPoint(geo.width() + content_w, 0))
        self._fly.setDuration(round(duration_s * 1000))
        self._fly.setEasingCurve(QEasingCurve.Type.Linear)

        self._fade = QPropertyAnimation(self, b"windowOpacity", self)
        self._fade.setStartValue(1.0)
        self._fade.setEndValue(0.0)
        self._fade.setDuration(round(FADE_S * 1000))
        self._fade.setEasingCurve(QEasingCurve.Type.InQuad)

    def launch(self) -> None:
        """Show the window and start the flight; it closes itself afterwards."""
        self.show()
        self._make_click_through()
        self._fly.start()
        QTimer.singleShot(round((self._duration_s - FADE_S) * 1000), self._fade.start)
        QTimer.singleShot(
            round((self._duration_s + CLOSE_AFTER_EXTRA_S) * 1000), self.close
        )

    def _make_click_through(self) -> None:
        # Backstop: Qt's WindowTransparentForInput is not honoured on every
        # Windows setup, so set the native extended styles too.
        hwnd = int(self.winId())
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style | _CLICK_THROUGH_EX)


class AirplaneOverlay:
    """Spawns and keeps alive one AirplaneWindow per connected screen."""

    def __init__(self) -> None:
        self._windows: list[AirplaneWindow] = []

    def show(self, text: str, duration_s: float) -> None:
        """Fly ``text`` across every screen over ``duration_s`` seconds."""
        for screen in QGuiApplication.screens():
            window = AirplaneWindow(screen, text, duration_s)
            window.destroyed.connect(lambda _=None, w=window: self._forget(w))
            self._windows.append(window)
            window.launch()
        log.debug("Airplane launched on %d screen(s)", len(QGuiApplication.screens()))

    def _forget(self, window: AirplaneWindow) -> None:
        self._windows = [w for w in self._windows if w is not window]
