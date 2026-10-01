# MeetingReminder for Windows (classic Outlook) — Design

Date: 1 October 2026
Status: Draft for review
Branch: `feat/windows-outlook-port` (local only — not pushed to the public upstream)

## 1. Goal

A Windows 11 equivalent of the macOS MeetingReminder app: about five minutes
before each meeting on the user's Outlook calendar (including Teams meetings), a
pink airplane trailing a banner "<title> in N min" flies across every monitor.

### Success criteria

- A real meeting on the classic Outlook default calendar shows the airplane once,
  4–6 minutes before its start, on every connected monitor.
- Recurring meetings fire for each occurrence; all-day and cancelled events never fire.
- The overlay is click-through, has smooth (anti-aliased) edges, and floats above
  normal and borderless-fullscreen windows.
- Tray menu offers: status, plane speed, start at login, test airplane, quit.
- No network calls; calendar data is read-only and never leaves the PC.

### Out of scope

- New Outlook (`olk.exe`), Microsoft Graph, or any sign-in flow. The new Outlook
  has no COM object model; if the user moves off classic Outlook this source stops
  working and a Graph source would be a separate piece of work.
- Non-default calendars, shared calendars, click-to-join Teams links.
- Packaging as a standalone `.exe` (possible later with PyInstaller).
- Changes to the macOS Swift project.

## 2. Constraints and decisions

| Decision | Choice | Reason |
|---|---|---|
| Calendar source | Classic Outlook via COM (`pywin32`) | User choice; no Entra app registration needed |
| Stack | Python 3.13 + PySide6 | Fits the user's conda setup; per-pixel alpha overlay and built-in tray icon |
| Environment | New conda env `meetingreminder` (python 3.13, pyside6, pywin32, pytest) | Keeps PySide6 out of `base` |
| Location | `windows/` folder in this repo | Shares artwork from `MeetingReminder/Assets.xcassets/` |

## 3. Structure

```
windows/
  pyproject.toml           # package metadata, ruff/black/isort config, pytest config
  environment.yml          # conda env "meetingreminder"
  README.md                # setup, usage, limitations
  meeting_reminder/
    __init__.py
    __main__.py            # entry point: python -m meeting_reminder
    models.py              # CalendarEvent dataclass, Speed(str, Enum), constants
    titles.py              # pure banner-title formatting
    outlook_source.py      # OutlookCalendarSource: COM read of default calendar
    poller.py              # pure AlertTracker (fire window + dedupe) + QThread poll worker
    overlay.py             # AirplaneOverlay: one transparent window per screen
    tray.py                # TrayMenu: QSystemTrayIcon + QMenu
    autostart.py           # HKCU Run-key read/write
    settings.py            # QSettings wrapper (speed)
    app.py                 # AppController wiring source, poller, overlay, tray
    assets.py              # resolves artwork paths from MeetingReminder/Assets.xcassets
  tests/
    test_poller.py
    test_titles.py
    test_outlook_mapping.py
```

Each unit has one job:

- **`models.py`**
  - `CalendarEvent(id: str, title: str, start: datetime, end: datetime)` is a frozen dataclass with timezone-aware datetimes.
  - `Speed(str, Enum)` has three values: `SLOW` = 22 s, `NORMAL` = 14 s and `FAST` = 8 s.
  - Module constants: `ALERT_MINUTES_BEFORE = 5`, `POLL_INTERVAL_S = 30`, `LOOKAHEAD_MIN = 60` and `NOTIFIED_CAP = 200`.
- **`titles.py`**
  - `banner_title(subject: str | None, attendee_names: list[str]) -> str`.
  - Returns "Meeting with A", "Meeting with A and B" or "Meeting with A, B +N more", which matches `AppleCalendarService.bannerTitle`.
  - Falls back to the subject, or to "Untitled Meeting" if there's no subject.
- **`outlook_source.py`**
  - `OutlookCalendarSource.fetch_upcoming(now: datetime) -> list[CalendarEvent]`.
  - Raises `OutlookNotRunning` when it can't attach to Outlook.
  - Mapping from a COM item to `CalendarEvent` is a separate pure function, `map_appointment(item_like, current_user: str) -> CalendarEvent | None`. It works on anything with the right attributes, so a stub object can test it.
- **`poller.py`**
  - `AlertTracker.due(events, now) -> list[tuple[CalendarEvent, int]]` is pure and has no Qt or COM code.
  - `PollWorker` is a `QObject` on a `QThread`. It initialises COM, calls the source on a timer and emits two signals: `due(CalendarEvent, int)` and `status(str)`.
- **`overlay.py`**
  - `show_airplane(title: str, minutes: int, duration_s: float) -> None` spawns one `AirplaneWindow` per `QScreen`.
- **`tray.py`**
  - Has no logic; it only emits user intents through signals.
- **`app.py`**
  - The only place where the parts are wired together.

## 4. Outlook data access

- **Attach to Outlook:** `win32com.client.GetActiveObject("Outlook.Application")`.
  - If that fails (Outlook isn't running), raise `OutlookNotRunning`. The tray shows "Outlook not running — open Outlook", and the app retries on the next poll.
  - Never call `Dispatch`, because it can launch a hidden Outlook instance.
- **Default calendar:** `Session.GetDefaultFolder(9)`, where 9 is `olFolderCalendar`.
- **Recurring meetings:**
  - Set `items.Sort("[Start]")`, then `items.IncludeRecurrences = True`, in that order. Outlook only expands recurring meetings into individual occurrences if the list is sorted first.
  - With recurrences on, `Count` is unreliable, so iterate with `GetFirst`/`GetNext` rather than by index.
- **Date filter:**
  - Use a DASL query (Outlook's `@SQL=` query syntax), with times in UTC formatted as `YYYY-MM-DD HH:MM`:
    `@SQL="urn:schemas:calendar:dtstart" >= '<now-1min UTC>' AND "urn:schemas:calendar:dtstart" <= '<now+60min UTC>'`
  - This avoids `Restrict("[Start] >= '...'")`, which parses dates using the Windows regional format and misreads en-AU dd/mm dates.
  - **Verify on the real machine:** that DASL with `IncludeRecurrences` expands recurring occurrences. If it doesn't, the fallback is a locale-formatted `Restrict` built with `time.strftime("%x %X")`, plus a Python-side window check on every returned item.
- **Python-side guards:** always re-check `start` against the window in Python, whatever the query returned.
- **Skip:**
  - `AllDayEvent == True`
  - `MeetingStatus in (5, 7)`, which are the two cancelled states (`olMeetingCanceled` and `olMeetingReceivedAndCanceled`).
- **Times:** use `StartUTC`/`EndUTC` and convert them to aware UTC datetimes, so time zones and daylight saving can't shift them.
- **Event identity:**
  - `f"{GlobalAppointmentID}|{StartUTC.isoformat()}"`, because every occurrence of a recurring meeting shares the same `GlobalAppointmentID`.
  - Fall back to `EntryID` when `GlobalAppointmentID` is empty.
- **Attendees:**
  - Take `Recipients[*].Name` and exclude the current user (`Session.CurrentUser.Name`) and the organiser-as-self.
  - Skip resource and room recipients (`Type == 3`, `olResource`).
- **Threading:** all COM calls happen on the `PollWorker` thread. It calls `pythoncom.CoInitialize()` once at startup and `CoUninitialize()` on exit. COM objects are never passed across threads; only `CalendarEvent` values cross over.
- **Errors:**
  - A COM error during a poll is logged, the tray status is updated and that poll is skipped. The timer keeps running.
  - An Outlook restart drops the cached COM object, so the next poll re-attaches.
- **Security prompt risk:** reading `Recipients` can trigger Outlook's "a program is trying to access…" prompt if antivirus isn't recognised. Defender normally prevents this. If the prompt appears, drop attendee names and use the subject only. That switch is a constant: `USE_ATTENDEE_NAMES`.

## 5. Alert logic (mirrors `CalendarPoller.swift`)

- **Poll schedule:** every `POLL_INTERVAL_S` (30 s), plus one poll immediately at start.
- **When an event fires:** `minutes = (start - now) / 60` and `4.0 <= minutes <= 6.0`, and its id hasn't been alerted yet.
- **Displayed minutes:** `ceil(minutes)`, so 4.1 minutes shows as "5 min".
- **Dedupe:** alerted ids are kept in insertion order and capped at 200, with the oldest evicted first.
- **Persistence:** none, which matches the Mac version. A restart inside the window can re-fire once, and that's acceptable.

## 6. Overlay

- **One window per `QScreen`.** Each window is the screen's width by 110 px tall. Its top edge sits 35% down the screen, which matches the Mac's "65% up" (Mac y=0 is the bottom).
- **Window flags:** `FramelessWindowHint | WindowStaysOnTopHint | Tool | WindowTransparentForInput | WindowDoesNotAcceptFocus`.
- **Window attributes:** `WA_TranslucentBackground`, `WA_ShowWithoutActivating`.
- **Click-through backstop:** also set `WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW` via `win32gui`, so clicks pass through even where Qt's input-transparency flag is ignored on Windows.
- **Content:**
  - A horizontal row of two images: the banner, then the airplane, overlapping by 10 px (`spacing = -10`).
  - The banner image is stretched behind the text "<title> in N min", in Comic Sans MS 28 pt, white, single line. The text is padded 50 px left and right and 22 px top and bottom.
  - The airplane is 220×220 and drawn behind the banner.
- **Animation:**
  - A `QPropertyAnimation` with linear easing moves the content from `x = -width` to `x = screen_width + width` over `duration_s`.
  - An opacity fade to 0 over 0.6 s starts at `duration_s - 0.6`.
  - Each window closes and is released at `duration_s + 1.5`.
- **High-DPI:** rely on Qt 6's per-monitor DPI scaling, and position by each screen's logical geometry.
- **Known limitation:** the overlay does not show over exclusive-fullscreen games, and may not show over some PowerPoint slideshow modes.

## 7. Tray and settings

- **Icon:** the `menubar` PNG. Tooltip: "MeetingReminder".
- **Menu items:**
  - Status line (disabled): "Outlook connected", "Outlook not running — open Outlook" or "Error: <short>".
  - Plane speed submenu: Slow, Normal, Fast (exclusive radio items).
  - Start at login (checkable).
  - Test airplane: fires "Test Meeting in 5 min" through the real overlay path.
  - Quit MeetingReminder.
- **Settings:** speed is stored with `QSettings("MeetingReminder", "MeetingReminder")`.
- **Start at login:**
  - Writes or removes `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\MeetingReminder`, with the value `"<env>\pythonw.exe" -m meeting_reminder` and the working directory set to `windows/`.
  - The checkbox state is read back from the registry rather than cached.
- **Single instance:** a named mutex (`win32event.CreateMutex`) stops a second copy from starting.

## 8. Logging

- Uses the standard `logging` module, never `print`. Output goes to a rotating file at `%LOCALAPPDATA%\MeetingReminder\meeting_reminder.log`, plus stderr when run from a console.
- Logs event subjects and times at DEBUG level only. The default level is INFO.

## 9. Testing

**Unit tests (pytest, no Outlook or Qt needed):**
- `test_poller.py`:
  - The window edges: 3.99, 4.0, 6.0 and 6.01 minutes.
  - The rounding (`ceil`).
  - Dedupe across polls.
  - Eviction at the cap.
  - Two occurrences of one recurring meeting fire separately.
- `test_titles.py`: 0, 1, 2 and 3 or more attendees; empty or `None` subject.
- `test_outlook_mapping.py`: `map_appointment` with stub objects covering:
  - all-day events and the two cancelled states, which are skipped
  - id composition, including the `EntryID` fallback
  - UTC conversion
  - excluding self and room recipients from attendee names

**Read-only smoke test:** `python -m meeting_reminder --list` prints the next hour's mapped events from the real Outlook and exits.

**Manual checks:**
- Test airplane on every monitor.
- Click-through over a browser.
- Over a borderless-fullscreen video.
- The start-at-login toggle survives a sign-out.
- One real meeting fires.

**Lint:** ruff, black and isort pass clean (line length 88).

## 10. Risks

| Risk | Mitigation |
|---|---|
| User switches to new Outlook | Documented limitation; Graph source would slot in behind the same `fetch_upcoming` interface |
| DASL + IncludeRecurrences doesn't expand occurrences | Verify first in implementation; fall back to locale-formatted Restrict + Python filter |
| Object Model Guard prompt on Recipients | `USE_ATTENDEE_NAMES` constant; subject-only fallback |
| Overlay not click-through on some setups | Win32 extended-style backstop |
