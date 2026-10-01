"""Start-at-login via the per-user Run registry key."""

from __future__ import annotations

import sys
import winreg
from pathlib import Path

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
VALUE_NAME = "MeetingReminder"


def launch_command() -> str:
    """Command that starts the app without a console window.

    Relies on the package being installed (editable) into the running env, so
    ``-m meeting_reminder`` resolves from any working directory.
    """
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    return f'"{pythonw}" -m meeting_reminder'


def is_enabled() -> bool:
    """True if a Run entry for the app exists."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, VALUE_NAME)
    except FileNotFoundError:
        return False
    return True


def set_enabled(enabled: bool) -> None:
    """Add or remove the Run entry.

    Raises:
        OSError: The registry could not be written.
    """
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE
    ) as key:
        if enabled:
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, launch_command())
        else:
            try:
                winreg.DeleteValue(key, VALUE_NAME)
            except FileNotFoundError:
                pass
