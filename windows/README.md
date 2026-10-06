# ✈️ MeetingReminder for Windows

Five minutes before each meeting, a pink airplane flies across your screen
towing a banner such as **"Meeting with Alex and Sam in 5 min"**. This version
runs on Windows 10/11, sits in the system tray, and reads your calendar from
**classic Outlook**, so Teams meetings and Exchange invites are included. It
only reads your calendar, on your own PC, and makes no network calls.

---

## Before you start: use classic Outlook

The app reads the calendar from the classic Outlook desktop app (`OUTLOOK.EXE`).
The **new Outlook** (`olk.exe`) offers no way for other programs to read its
calendar, so the app can't use it.

To check which Outlook you have, look at the top-right corner of the Outlook
window. If you see a **New Outlook** toggle switched **on**, you're using the
new Outlook. Then either:

- start **Outlook (classic)** from the Start menu (search "Outlook classic"), or
- switch the **New Outlook** toggle off and let Outlook restart.

You can keep the new Outlook open as well, as long as classic Outlook is also
running. You don't need to keep it in the foreground, but it must be running.

---

## Install (one time)

You need [Miniconda](https://docs.anaconda.com/miniconda/) or Anaconda.

```powershell
cd <path-to-repo>\windows
& "$env:USERPROFILE\miniconda3\Scripts\conda.exe" env create -f environment.yml
& "$env:USERPROFILE\miniconda3\envs\meetingreminder\python.exe" -m pip install -e . --no-deps
```

This creates a self-contained `meetingreminder` environment with PySide6 and
pywin32, and registers the app in that environment. Your other Python setups
are left untouched.

Check it can see your calendar:

```powershell
& "$env:USERPROFILE\miniconda3\envs\meetingreminder\python.exe" -m meeting_reminder --list
```

This prints any meetings in the next hour. If it says *Classic Outlook is not
running*, open classic Outlook and try again.

---

## Run

```powershell
& "$env:USERPROFILE\miniconda3\envs\meetingreminder\pythonw.exe" -m meeting_reminder
```

A pink airplane icon appears in the system tray. It may be hidden under the
**^** arrow next to the clock. Click it for the menu:

| Menu item | What it does |
|---|---|
| Status | Shows *Outlook connected* or *Outlook not running* |
| Plane speed | Slow / Normal / Fast |
| Start at login | Starts the app automatically each time you sign in |
| Test airplane | Shows a test airplane immediately |
| Quit MeetingReminder | Stops the app |

Tick **Start at login** once, and from then on it starts by itself whenever you
sign in. It doesn't matter whether you open Outlook before or after the app:
while Outlook is closed it checks again every 30 seconds, and connects once
Outlook is open.

---

## How it behaves

- Checks your default Outlook calendar every 30 seconds and shows the airplane
  once per meeting, about 5 minutes before it starts.
- Each occurrence of a recurring meeting gets its own alert.
- All-day events and cancelled meetings never trigger it.
- Plays on your main display only (the one set as "main display" in Windows
  Display settings). Mouse clicks pass through the airplane, so it never
  gets in your way.
- Shows on top of normal and borderless-fullscreen windows. It won't appear
  over exclusive-fullscreen games.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Tray says *Outlook not running* | Open **classic** Outlook. The new Outlook won't work. |
| A code or settings change isn't showing | Use **Quit MeetingReminder** in the tray, then start the app again. Only one copy can run at a time, so a second launch quietly exits while the old one is still open. |
| Outlook shows a security prompt ("a program is trying to access…") | Set `USE_ATTENDEE_NAMES = False` in `meeting_reminder/outlook_source.py`. The banner will then show the meeting subject instead of attendee names. |
| Start at login stopped working | The startup entry points at this folder. If you moved the repo, untick **Start at login** and tick it again. |
| Want to see what it's doing | Read the log at `%LOCALAPPDATA%\MeetingReminder\meeting_reminder.log`, or run `python.exe -m meeting_reminder --debug` in a console. |

---

## Customise

| What | Where |
|---|---|
| Alert lead time, check interval | `meeting_reminder/models.py` |
| Plane speeds | `Speed.flight_seconds` in `models.py` |
| Font, weight, size, padding, height on screen | Constants at the top of `meeting_reminder/overlay.py` (currently Quicksand Light) |

The airplane, banner and icon artwork is shared with the macOS app, in
`../MeetingReminder/Assets.xcassets/`. After any change, use **Quit
MeetingReminder** and start the app again.

---

## Development

```powershell
$py = "$env:USERPROFILE\miniconda3\envs\meetingreminder\python.exe"
& $py -m pytest -q
& $py -m ruff check . ; & $py -m black --check . ; & $py -m isort --check .
```
