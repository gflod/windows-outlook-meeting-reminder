# ✈️ MeetingReminder for Windows

The Windows version of MeetingReminder. Five minutes before each meeting in your
**classic Outlook** calendar (Teams meetings included), a pink airplane trailing
"<meeting> in 5 min" flies across every monitor.

It runs from the system tray, reads Outlook read-only on your PC, and makes no
network calls.

## Requirements

- Windows 10/11
- **Classic Outlook** (`OUTLOOK.EXE`) installed and running. The new Outlook
  (`olk.exe`) has no COM interface, so it can't be used as the calendar source.
  You can keep the new Outlook open as well; classic just needs to be running.
- Miniconda

## Setup

```powershell
cd windows
& "$env:USERPROFILE\miniconda3\Scripts\conda.exe" env create -f environment.yml
& "$env:USERPROFILE\miniconda3\envs\meetingreminder\python.exe" -m pip install -e . --no-deps
```

The editable install lets `python -m meeting_reminder` run from any folder,
which start-at-login depends on.

## Usage

```powershell
$py = "$env:USERPROFILE\miniconda3\envs\meetingreminder\pythonw.exe"
& $py -m meeting_reminder            # tray app, no console window
```

From a console, `python.exe -m meeting_reminder --debug` shows the log as it
runs. `python.exe -m meeting_reminder --list` prints the next hour's events and
exits, which is a quick way to check it can see your calendar.

The tray menu has:
- **Status**: Outlook connected / not running.
- **Plane speed**: Slow, Normal or Fast.
- **Start at login**: adds or removes an `HKCU\...\Run` entry.
- **Test airplane**: shows the airplane now.
- **Quit**.

The log is written to `%LOCALAPPDATA%\MeetingReminder\meeting_reminder.log`.

## Behaviour

- Polls the default Outlook calendar every 30 seconds. Each meeting fires once,
  when it's 4–6 minutes away.
- Each occurrence of a recurring meeting gets its own alert. All-day events and
  cancelled meetings never fire.
- The banner reads "Meeting with Alex and Sam" when there are other attendees,
  and shows the subject otherwise. Rooms are left out of the names.
- The overlay ignores mouse clicks and stays on top of normal and
  borderless-fullscreen windows. It does not show over exclusive-fullscreen
  games.

## Customisation

| What | Where |
|---|---|
| Alert lead time, poll interval | `meeting_reminder/models.py` constants |
| Speeds | `Speed.flight_seconds` in `models.py` |
| Font, padding, plane size, height on screen | constants at the top of `overlay.py` |
| Subject-only banners (if Outlook shows a security prompt) | `USE_ATTENDEE_NAMES` in `outlook_source.py` |

The artwork is read from `../MeetingReminder/Assets.xcassets/`, the same files
the macOS app uses.

## Development

```powershell
$py = "$env:USERPROFILE\miniconda3\envs\meetingreminder\python.exe"
& $py -m pytest -q
& $py -m ruff check . ; & $py -m black --check . ; & $py -m isort --check .
```
