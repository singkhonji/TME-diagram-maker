"""Low-level I/O against the TME System Diagram dialog.

TME's Schematic Table is a custom-drawn canvas: UI Automation reports
``ControlType: Custom(50025)``, no Name, not keyboard focusable, so nothing can
be read or written through the accessibility tree.  The only channel that works
is synthetic input plus the clipboard.

The one rule everything here follows: **never sleep and hope**.  Every action is
followed by evidence that it landed.  A cell read polls the clipboard until TME
answers, which doubles as the pacing mechanism -- if the app stutters for 800 ms
we wait exactly 800 ms, instead of guessing 100 ms and corrupting the row or
guessing 2000 ms and wasting an hour over a full project.

Two keys are deliberately never sent:

``Ctrl+A``
    Inside a cell editor it selects the text; outside one it selects *every row*
    in the grid, and the paste that follows would overwrite rows already
    verified.  It is also unnecessary -- F2 opens the editor with the old text
    already selected, so a paste replaces it either way (measured 2026-08-12).

``Esc`` on its own
    It cancels the whole System Diagram dialog and discards every row entered so
    far.  It is only ever sent after a successful cell read, which is positive
    proof that a cell editor is open for it to close.
"""

from __future__ import annotations

import ctypes
import os
import tempfile
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass

import win32api
import win32clipboard
import win32con
import win32gui

import config

# --------------------------------------------------------------------------
# Errors
# --------------------------------------------------------------------------


class TmeIoError(Exception):
    """Base class: any condition that must stop the run."""


class Aborted(TmeIoError):
    """The user hit the abort hotkey."""


class FocusLost(TmeIoError):
    """The System Diagram window is not in front, so keys must not be sent."""


class CellReadTimeout(TmeIoError):
    """TME never answered a Ctrl+C.  Never treat this as 'value unchanged'."""


class EditModeFailed(TmeIoError):
    """F2 did not open a cell editor, so the following keys would misfire.

    Measured behaviour: without an open editor, Ctrl+C copies the whole *row*
    (tab separated) and Esc cancels the entire dialog, discarding every row
    entered so far.  Detecting this and stopping is the difference between a
    failed cell and a lost afternoon.
    """


class StrayWindow(TmeIoError):
    """Something else took the foreground -- most often TME's own save prompt.

    TME pops "Save current project?" on its own schedule.  Its default button is
    Yes, so a stray Enter would silently save; keys must never be sent while it
    is up.
    """


class WindowNotFound(TmeIoError):
    """The System Diagram dialog is not open."""


# --------------------------------------------------------------------------
# SendInput plumbing
# --------------------------------------------------------------------------

ULONG_PTR = wintypes.WPARAM
_user32 = ctypes.WinDLL("user32", use_last_error=True)


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class _HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT), ("ki", _KEYBDINPUT), ("hi", _HARDWAREINPUT)]


class _INPUT(ctypes.Structure):
    # The union must carry all three members so sizeof(INPUT) matches what
    # SendInput expects; a keyboard-only struct is too small and the call fails.
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]


_INPUT_KEYBOARD = 1
_KEYEVENTF_KEYUP = 0x0002
_KEYEVENTF_UNICODE = 0x0004
_KEYEVENTF_EXTENDEDKEY = 0x0001

# Keys this module needs.  Extended keys (nav cluster) must set the extended
# flag or some WPF grids read them as the numpad equivalents.
VK: dict[str, int] = {
    "f2": 0x71,
    "f12": 0x7B,
    "enter": 0x0D,
    "tab": 0x09,
    "esc": 0x1B,
    "home": 0x24,
    "end": 0x23,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "delete": 0x2E,
    "ctrl": 0x11,
    "shift": 0x10,
    "alt": 0x12,
    "a": 0x41,
    "c": 0x43,
    "i": 0x49,
    "r": 0x52,
    "v": 0x56,
}

_EXTENDED = frozenset({"home", "end", "left", "up", "right", "down", "delete"})


def _key_event(name: str, *, up: bool) -> _INPUT:
    vk = VK[name]
    flags = _KEYEVENTF_KEYUP if up else 0
    if name in _EXTENDED:
        flags |= _KEYEVENTF_EXTENDEDKEY
    scan = _user32.MapVirtualKeyW(vk, 0)
    event = _INPUT(type=_INPUT_KEYBOARD)
    event.u.ki = _KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=flags, time=0, dwExtraInfo=0)
    return event


def _char_event(char: str, *, up: bool) -> _INPUT:
    flags = _KEYEVENTF_UNICODE | (_KEYEVENTF_KEYUP if up else 0)
    event = _INPUT(type=_INPUT_KEYBOARD)
    event.u.ki = _KEYBDINPUT(
        wVk=0, wScan=ord(char), dwFlags=flags, time=0, dwExtraInfo=0
    )
    return event


def _send(events: list[_INPUT]) -> None:
    array = (_INPUT * len(events))(*events)
    sent = _user32.SendInput(len(events), array, ctypes.sizeof(_INPUT))
    if sent != len(events):
        raise TmeIoError(
            f"SendInput sent {sent}/{len(events)} events "
            f"(WinError {ctypes.get_last_error()})"
        )


# --------------------------------------------------------------------------
# Clipboard
# --------------------------------------------------------------------------

_MESSAGE_BOX_RULE = "-" * 27
"""Windows message boxes answer Ctrl+C with their own text, ruled off like this.

So a popup appearing mid-read hands us the dialog's caption and buttons instead
of the cell's contents.  Recognising that is the difference between "a dialog
interrupted us, deal with it and read again" and "the cell contains the wrong
value", which would stop the run for no reason.
"""


def _is_message_box_text(value: str) -> bool:
    return value.lstrip().startswith(_MESSAGE_BOX_RULE) and value.count(
        _MESSAGE_BOX_RULE
    ) >= 2


MISSING = object()
"""Returned by :func:`clip_get` when the clipboard holds no text at all.

This is a *distinct* state from an empty string, and the difference matters: if
TME answers a Ctrl+C on an empty cell by emptying the clipboard, we can still
tell "TME responded, the cell is empty" apart from "TME never responded".
"""


def clip_get() -> str | object:
    """Read CF_UNICODETEXT, or :data:`MISSING` if the clipboard has no text."""
    last: Exception | None = None
    for _ in range(config.CLIPBOARD_OPEN_ATTEMPTS):
        try:
            win32clipboard.OpenClipboard()
        except Exception as exc:  # another app holds the clipboard lock
            last = exc
            time.sleep(config.CLIPBOARD_OPEN_BACKOFF_S)
            continue
        try:
            if not win32clipboard.IsClipboardFormatAvailable(
                win32clipboard.CF_UNICODETEXT
            ):
                return MISSING
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        except Exception as exc:
            last = exc
            time.sleep(config.CLIPBOARD_OPEN_BACKOFF_S)
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
    raise TmeIoError(f"could not read the clipboard: {last}")


def clipboard_sequence() -> int:
    """Windows' clipboard change counter.

    It ticks on every write by any process, so it is the one way to tell that a
    write we did not make has landed -- ``clip_get`` only shows the content, and
    an overwrite with the same text would be invisible.
    """
    return _user32.GetClipboardSequenceNumber()


def clip_settle(text: str) -> None:
    """Block until the clipboard holds ``text`` and has stopped changing.

    TME finishes writing the clipboard after it has already answered a Ctrl+C.
    Setting our own value and pasting straight away races that late write, and
    losing the race pastes the previously read cell's text.  Waiting for the
    sequence number to hold still means the paste only happens once nothing else
    is in flight -- a condition, not a guessed delay.
    """
    deadline = time.monotonic() + config.CLIPBOARD_SETTLE_TIMEOUT_S
    last = clipboard_sequence()
    stable = 0
    while time.monotonic() < deadline:
        time.sleep(config.CLIPBOARD_POLL_S)
        current = clipboard_sequence()
        if current == last:
            stable += 1
            if stable >= config.CLIPBOARD_SETTLE_SAMPLES:
                if clip_get() == text:
                    return
                # Something else won; take the clipboard back and start over.
                clip_set(text)
                last = clipboard_sequence()
                stable = 0
        else:
            last = current
            stable = 0

    raise TmeIoError(
        f"the clipboard would not settle on {text!r} within "
        f"{config.CLIPBOARD_SETTLE_TIMEOUT_S:.1f} s; refusing to paste into a "
        f"cell while another write is still in flight"
    )


def clip_set(text: str) -> None:
    """Put ``text`` on the clipboard and confirm it stuck.

    The confirmation is not paranoia: TME writes the clipboard asynchronously
    when it copies, so a write of ours can land while one of its own is still in
    flight and lose.  Pasting whatever happens to be there is how a read-back
    sentinel ends up inside a cell.
    """
    last: Exception | None = None
    for _ in range(config.CLIPBOARD_OPEN_ATTEMPTS):
        try:
            win32clipboard.OpenClipboard()
        except Exception as exc:
            last = exc
            time.sleep(config.CLIPBOARD_OPEN_BACKOFF_S)
            continue
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        except Exception as exc:
            last = exc
            time.sleep(config.CLIPBOARD_OPEN_BACKOFF_S)
            continue
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

        if clip_get() == text:
            return
        last = TmeIoError("clipboard did not hold the value just written")
        time.sleep(config.CLIPBOARD_OPEN_BACKOFF_S)

    raise TmeIoError(f"could not write the clipboard: {last}")


# --------------------------------------------------------------------------
# Adaptive throttle
# --------------------------------------------------------------------------


class Throttle:
    """Key delay and read timeout derived from measured TME latency.

    Fast while the app is responsive, automatically slower when it stutters --
    which is the whole point: a fixed delay is either too slow all run or too
    fast exactly when the app is struggling.
    """

    def __init__(self) -> None:
        self.baseline_s = 0.05
        self.ewma_s = self.baseline_s
        self.key_delay_s = config.KEY_DELAY_START_S
        self.read_timeout_s = config.READ_TIMEOUT_START_S
        self.consecutive_timeouts = 0
        self.penalties = 0
        self.max_latency_s = 0.0

    @staticmethod
    def _clamp(value: float, low: float, high: float) -> float:
        return max(low, min(high, value))

    def observe(self, latency_s: float) -> None:
        """Fold a successful read's latency into the running estimate."""
        self.consecutive_timeouts = 0
        self.max_latency_s = max(self.max_latency_s, latency_s)
        self.ewma_s = 0.75 * self.ewma_s + 0.25 * latency_s
        ratio = self.ewma_s / self.baseline_s
        self.key_delay_s = self._clamp(
            config.KEY_DELAY_START_S * ratio,
            config.KEY_DELAY_MIN_S,
            config.KEY_DELAY_MAX_S,
        )
        self.read_timeout_s = self._clamp(
            self.ewma_s * 6.0,
            config.READ_TIMEOUT_MIN_S,
            config.READ_TIMEOUT_MAX_S,
        )

    def penalize(self) -> None:
        """A read timed out: back off hard before retrying."""
        self.consecutive_timeouts += 1
        self.penalties += 1
        self.key_delay_s = self._clamp(
            self.key_delay_s * 2.0, config.KEY_DELAY_MIN_S, config.KEY_DELAY_MAX_S
        )
        self.read_timeout_s = self._clamp(
            self.read_timeout_s * 2.0,
            config.READ_TIMEOUT_MIN_S,
            config.READ_TIMEOUT_MAX_S,
        )

    def summary(self) -> str:
        return (
            f"latency avg {self.ewma_s * 1000:.0f} ms, peak "
            f"{self.max_latency_s * 1000:.0f} ms, key delay "
            f"{self.key_delay_s * 1000:.0f} ms, {self.penalties} backoff(s)"
        )


# --------------------------------------------------------------------------
# Abort watchdog
# --------------------------------------------------------------------------


class AbortWatch:
    """Polls the abort hotkey in a daemon thread.

    ``GetAsyncKeyState`` sees the key regardless of which window has focus, so
    the hotkey still works while TME is in front and needs no extra dependency.
    """

    def __init__(self) -> None:
        self.event = threading.Event()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        def loop() -> None:
            while not self._stop.is_set():
                if _user32.GetAsyncKeyState(config.ABORT_VK) & 0x8000:
                    self.event.set()
                    return
                time.sleep(0.03)

        self._thread = threading.Thread(target=loop, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()

    def check(self) -> None:
        if self.event.is_set():
            raise Aborted(f"aborted by {config.ABORT_KEY_NAME}")


# --------------------------------------------------------------------------
# Window handling
# --------------------------------------------------------------------------


def find_window(title_match: str = config.WINDOW_TITLE_MATCH) -> int:
    """Handle of the visible window whose title contains ``title_match``."""
    matches: list[int] = []

    def visit(hwnd: int, _: object) -> bool:
        if win32gui.IsWindowVisible(hwnd) and title_match in win32gui.GetWindowText(hwnd):
            matches.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    if not matches:
        raise WindowNotFound(
            f"no visible window titled {title_match!r} -- open Cubicost TME and "
            f"the System Diagram dialog first"
        )
    return matches[0]


def foreground_title() -> str:
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())


def describe_window(hwnd: int) -> str:
    try:
        return (
            f"hwnd={hwnd} title={win32gui.GetWindowText(hwnd)!r} "
            f"class={win32gui.GetClassName(hwnd)} rect={win32gui.GetWindowRect(hwnd)}"
        )
    except Exception:
        return f"hwnd={hwnd} (gone)"


def capture_window(hwnd: int, path: str) -> str | None:
    """Save a picture of ``hwnd``, even when it is behind other windows.

    Used when the run stops on an unexpected popup: the user gets to see what
    interrupted it instead of a bare window title.
    """
    try:
        from PIL import Image
    except ImportError:
        return None

    _PW_RENDERFULLCONTENT = 0x00000002
    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    try:
        rect = wintypes.RECT()
        _user32.GetWindowRect(hwnd, ctypes.byref(rect))
        width, height = rect.right - rect.left, rect.bottom - rect.top
        if width <= 0 or height <= 0:
            return None
        window_dc = _user32.GetWindowDC(hwnd)
        mem_dc = gdi32.CreateCompatibleDC(window_dc)
        bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
        gdi32.SelectObject(mem_dc, bitmap)
        _user32.PrintWindow(hwnd, mem_dc, _PW_RENDERFULLCONTENT)
        buffer = ctypes.create_string_buffer(width * height * 4)
        gdi32.GetBitmapBits(bitmap, width * height * 4, buffer)
        Image.frombuffer(
            "RGBA", (width, height), buffer, "raw", "BGRA", 0, 1
        ).convert("RGB").save(path)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(mem_dc)
        _user32.ReleaseDC(hwnd, window_dc)
        return path
    except Exception:
        return None


def compare_images(
    shot_path: str,
    reference_path: str,
    max_diff: float,
    box: tuple[float, float, float, float] | None = None,
) -> tuple[bool, str]:
    """Mean per-channel difference between two images, optionally cropped."""
    try:
        from PIL import Image, ImageChops, ImageStat
    except ImportError:
        return False, "Pillow is not installed, so the dialog cannot be identified"

    if not os.path.exists(reference_path):
        return False, f"no reference image at {reference_path}"

    try:
        with Image.open(shot_path) as shot, Image.open(reference_path) as reference:
            if shot.size != reference.size:
                return False, f"size {shot.size} != reference {reference.size}"
            left, top, right, bottom = (0, 0, *shot.size)
            if box:
                width, height = shot.size
                left, top = int(width * box[0]), int(height * box[1])
                right, bottom = int(width * box[2]), int(height * box[3])
            crop = (left, top, right, bottom)
            difference = ImageChops.difference(
                shot.convert("RGB").crop(crop), reference.convert("RGB").crop(crop)
            )
            mean = sum(ImageStat.Stat(difference).mean) / 3.0
    except Exception as exc:
        return False, f"comparison failed: {exc}"

    verdict = "matches" if mean <= max_diff else "does not match"
    return mean <= max_diff, f"{verdict} the reference (mean difference {mean:.2f})"


def looks_like(
    hwnd: int,
    reference_path: str,
    max_diff: float,
    box: tuple[float, float, float, float] | None = None,
) -> tuple[bool, str]:
    """Is ``hwnd`` the dialog stored at ``reference_path``?

    Used to decide whether a popup may be answered automatically.  Matching the
    picture rather than the window title matters: TME titles several different
    Yes/No questions "OK", and clicking through an unrecognised one would be a
    decision this script has no business making.
    """
    shot_path = os.path.join(tempfile.gettempdir(), "tme_dialog_check.png")
    if capture_window(hwnd, shot_path) is None:
        return False, "could not capture the dialog"
    return compare_images(shot_path, reference_path, max_diff, box)


def click(x: int, y: int) -> None:
    """Click a screen point.

    The only mouse use in this module, and only ever aimed at a dialog button
    that has already been identified by picture.  Filling the table itself never
    touches the mouse, which is what keeps it immune to zoom and scrolling.
    """
    win32api.SetCursorPos((x, y))
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def click_in_window(hwnd: int, ratio: tuple[float, float]) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    click(left + int((right - left) * ratio[0]), top + int((bottom - top) * ratio[1]))


def focus_window(hwnd: int, *, maximize: bool = True, timeout_s: float = 5.0) -> None:
    """Bring ``hwnd`` to the front (and maximise it) before typing into it.

    Maximising is deliberate rather than left to the user: it makes the dialog's
    geometry the same on every run, so scrolling behaves identically.
    """
    if maximize:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if win32gui.GetForegroundWindow() == hwnd:
            return
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            # Windows blocks foreground changes from a background process unless
            # an input event has just occurred; a bare Alt tap satisfies that.
            _send([_key_event("alt", up=False), _key_event("alt", up=True)])
        time.sleep(0.05)

    raise FocusLost(
        f"could not bring window {hwnd} to the front "
        f"(foreground is {foreground_title()!r})"
    )


# --------------------------------------------------------------------------
# The session object: everything above, wired together
# --------------------------------------------------------------------------


@dataclass
class CellRead:
    """One clipboard read-back."""

    value: str
    latency_s: float
    cleared: bool = False
    """True when TME answered by emptying the clipboard (an empty cell)."""


class TmeSession:
    """A guarded input channel to one System Diagram window."""

    def __init__(
        self,
        *,
        dry_run: bool = False,
        verbose: bool = False,
        auto_dismiss_save: bool = False,
    ) -> None:
        self.dry_run = dry_run
        self.verbose = verbose
        self.auto_dismiss_save = auto_dismiss_save
        self.throttle = Throttle()
        self.abort = AbortWatch()
        self.hwnd: int | None = None
        self._sentinel_counter = 0
        self.reads = 0
        self.writes = 0
        self.saves_dismissed = 0

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "TmeSession":
        self.hwnd = find_window()
        self.abort.start()
        if not self.dry_run:
            focus_window(self.hwnd)
        return self

    def __exit__(self, *exc: object) -> None:
        self.abort.stop()

    # -- guards ------------------------------------------------------------

    def guard(self) -> None:
        """Refuse to send input unless abort is clear and TME is in front.

        Typing into the wrong window is the worst failure this script can cause,
        so this runs before every single key batch.
        """
        self.abort.check()
        if self.dry_run:
            return
        assert self.hwnd is not None
        if win32gui.GetForegroundWindow() == self.hwnd:
            return

        # Give a brief alt-tab a chance to resolve itself before giving up.
        deadline = time.monotonic() + config.FOCUS_REACQUIRE_TIMEOUT_S
        while time.monotonic() < deadline:
            self.abort.check()
            time.sleep(0.05)
            if win32gui.GetForegroundWindow() == self.hwnd:
                return

        intruder = win32gui.GetForegroundWindow()

        if self.auto_dismiss_save and self._try_dismiss_save_prompt(intruder):
            return

        detail = describe_window(intruder)
        shot = capture_window(intruder, config.INTRUDER_SHOT_FILE)

        # A Qt window owned by TME is one of its own modals -- the save prompt,
        # the duplicate-board-name question, or a validation complaint.  Anything
        # else belongs to a different application.
        try:
            owned_by_tme = win32gui.GetClassName(intruder).startswith(
                "Qt5"
            ) and self._owned_by_tme(intruder)
        except Exception:
            owned_by_tme = False

        message = (
            f"{config.WINDOW_TITLE_MATCH!r} no longer has focus -- stopping "
            f"rather than typing blind.\n  intruder: {detail}"
        )
        if shot:
            message += f"\n  picture of it: {shot}"
        if owned_by_tme:
            message += (
                "\n  This looks like one of TME's own dialogs (its save prompt "
                "appears on its own schedule).  Dismiss it, then resume."
            )
        raise StrayWindow(message)

    def _owned_by_tme(self, hwnd: int) -> bool:
        """Is ``hwnd`` a modal belonging to TME rather than another application?

        Which window owns the modal depends on which one raised it: the save
        prompt is owned by TME's main window, while the questions the schematic
        table asks are owned by the System Diagram dialog itself.  Both count.
        """
        try:
            owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
        except Exception:
            return False
        if not owner:
            return False
        if owner == self.hwnd:
            return True
        title = win32gui.GetWindowText(owner)
        return "Cubicost" in title or config.WINDOW_TITLE_MATCH in title

    def _try_dismiss_save_prompt(self, hwnd: int) -> bool:
        """Answer No to TME's save prompt, if that is genuinely what this is.

        Every condition has to hold -- right title, right owner, and a picture
        that matches the stored reference -- because the alternative is clicking
        a button in a dialog nobody has read.
        """
        if win32gui.GetWindowText(hwnd) != config.SAVE_PROMPT_TITLE:
            return False
        if not self._owned_by_tme(hwnd):
            return False

        matched, why = looks_like(
            hwnd,
            config.SAVE_PROMPT_REFERENCE,
            config.SAVE_PROMPT_MAX_DIFF,
            config.SAVE_PROMPT_COMPARE_BOX,
        )
        if not matched:
            print(f"    a dialog titled {config.SAVE_PROMPT_TITLE!r} appeared but "
                  f"{why}; leaving it alone")
            return False

        print("    TME asked to save the project; answering No and carrying on")
        click_in_window(hwnd, config.SAVE_PROMPT_NO_RATIO)

        # Wait for the prompt itself to close.  Where focus goes afterwards is a
        # separate question -- it often lands on TME's main window rather than
        # back on the dialog -- so bring the dialog forward as its own step.
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                break
            time.sleep(0.1)
        else:
            capture_window(hwnd, config.INTRUDER_SHOT_FILE)
            print(f"    the save prompt did not close (see "
                  f"{config.INTRUDER_SHOT_FILE}); stopping instead")
            return False

        assert self.hwnd is not None
        try:
            focus_window(self.hwnd, timeout_s=5.0)
        except FocusLost:
            print("    could not get back to the table after the save prompt")
            return False

        self.saves_dismissed += 1
        return True

    # -- key sending -------------------------------------------------------

    def tap(self, *names: str) -> None:
        """Press and release each key in turn."""
        self.guard()
        for name in names:
            self._trace(f"tap {name}")
            if self.dry_run:
                continue
            _send([_key_event(name, up=False)])
            time.sleep(self.throttle.key_delay_s)
            _send([_key_event(name, up=True)])
            time.sleep(self.throttle.key_delay_s)

    def chord(self, modifier: str, key: str) -> None:
        """Hold ``modifier``, tap ``key``, release -- e.g. ``chord("ctrl", "v")``."""
        self.guard()
        self._trace(f"chord {modifier}+{key}")
        if self.dry_run:
            return
        _send([_key_event(modifier, up=False)])
        time.sleep(self.throttle.key_delay_s)
        _send([_key_event(key, up=False)])
        time.sleep(self.throttle.key_delay_s)
        _send([_key_event(key, up=True)])
        time.sleep(self.throttle.key_delay_s)
        _send([_key_event(modifier, up=True)])
        time.sleep(self.throttle.key_delay_s)

    def type_text(self, text: str) -> None:
        """Type ``text`` as Unicode characters (for dropdowns that reject paste)."""
        self.guard()
        if self.dry_run:
            self._trace(f"type {text!r}")
            return
        for char in text:
            _send([_char_event(char, up=False), _char_event(char, up=True)])
            time.sleep(self.throttle.key_delay_s)

    # -- the state oracle --------------------------------------------------

    def _next_sentinel(self) -> str:
        # Unique per call, so a stale clipboard from an earlier read can never be
        # mistaken for a fresh answer.
        self._sentinel_counter += 1
        return f"__TMEIO_SENTINEL_{self._sentinel_counter:06d}__"

    def _await_clipboard(
        self, sentinel: str, timeout_s: float | None = None
    ) -> tuple[str | object, float]:
        """Poll until the clipboard differs from ``sentinel``.

        This poll *is* the pacing mechanism.  A stuttering TME is waited out for
        exactly as long as it needs, rather than against a guessed delay that is
        either too short (corrupt data) or too long (a wasted hour per project).
        """
        timeout = self.throttle.read_timeout_s if timeout_s is None else timeout_s
        started = time.monotonic()
        deadline = started + timeout
        got: str | object = sentinel
        while time.monotonic() < deadline:
            self.abort.check()
            current = clip_get()
            # An empty clipboard is a transient state, not an answer: a copy is
            # EmptyClipboard followed by SetClipboardData, and sampling between
            # the two used to be read as "TME replied", which desynchronised
            # every step that followed.
            if current is not MISSING and current != sentinel:
                got = current
                break
            time.sleep(config.CLIPBOARD_POLL_S)
        return got, time.monotonic() - started

    # A bare Ctrl+C -- the "Copy Row" shortcut in TME's own menu -- was measured
    # to return nothing at all when a cell merely has focus, so whole-row reads
    # are not available and verification goes cell by cell.  The tab check inside
    # read_cell still guards the case where it behaves otherwise.

    def read_cell(
        self,
        *,
        retries: int = 1,
        timeout_s: float | None = None,
        expect_empty: bool = False,
    ) -> CellRead:
        """Copy the focused cell and return what it holds.

        Doubles as proof that the cell really has focus: if TME does not answer
        the Ctrl+C, that is a failure, never "the value was unchanged".

        ``expect_empty`` is for probing a cell that should be blank.  An empty
        cell answers with silence, so such a read uses a short timeout and does
        not drag the latency estimate or the backoff counter with it.  Silence is
        only ever weak evidence of emptiness -- the authoritative checks are the
        post-write row sweep and the previous-row-intact check.
        """
        last_error: TmeIoError | None = None
        popup_reads = 0
        if expect_empty:
            timeout_s = timeout_s or config.EXPECT_EMPTY_TIMEOUT_S
            retries = 0
        attempt = 0
        while True:
            sentinel = self._next_sentinel()
            clip_set(sentinel)

            self.tap("f2")
            self.chord("ctrl", "c")
            got, latency = self._await_clipboard(sentinel, timeout_s)

            if got is MISSING or got == sentinel:
                # No answer: we cannot tell an empty cell from a dropped Ctrl+C,
                # and we have no evidence an editor is open -- so do NOT send Esc
                # to close one.  A bare Esc cancels the whole dialog.
                if not expect_empty:
                    self.throttle.penalize()
                attempt += 1
                self._trace(
                    f"read -> NO ANSWER (attempt {attempt}; the editor, if "
                    f"F2 opened one, is left open on purpose)"
                )
                last_error = CellReadTimeout(
                    f"no clipboard answer within "
                    f"{(timeout_s or self.throttle.read_timeout_s) * 1000:.0f} ms "
                    f"(attempt {attempt}/{retries + 1}); the cell may be empty or "
                    f"the cell may not be focused -- both are failures here"
                )
                if attempt > retries:
                    raise last_error
                continue

            assert isinstance(got, str)

            if _is_message_box_text(got):
                # A Windows message box answers Ctrl+C with its own text, so this
                # is the popup talking, not the cell.  Deal with the popup (the
                # guard dismisses TME's save prompt, or stops for anything else)
                # and read again rather than reporting a wrong value.
                popup_reads += 1
                if popup_reads > 3:
                    raise CellReadTimeout(
                        "a dialog kept answering the clipboard instead of the "
                        "cell; stopping"
                    )
                self.guard()
                continue

            if "\t" in got:
                # Tabs mean this was a row copy, i.e. F2 never opened an editor.
                # Esc now would cancel the dialog, so bail out without pressing it.
                raise EditModeFailed(
                    f"F2 did not open a cell editor: Ctrl+C returned a whole row "
                    f"({len(got.split(chr(9)))} tab-separated fields). Refusing "
                    f"to send Esc, which would cancel the System Diagram dialog."
                )

            self._trace(f"read -> {got!r} ({latency * 1000:.0f} ms)")
            self.tap("esc")  # safe: proven to be inside a cell editor
            self.throttle.observe(latency)
            self.reads += 1
            return CellRead(value=got, latency_s=latency)

    def write_cell(self, value: str, *, commit: str = "enter") -> None:
        """Replace the focused cell's contents with ``value``.

        F2 opens the editor with the old text selected, so the paste replaces it
        without any Ctrl+A -- see the module docstring for why that key is never
        sent.  Enter commits and leaves focus on the same cell.

        Does not verify: callers pair this with a read-back of the cell and a
        sweep of the whole row, because a per-cell read-back proves the value
        arrived but not that it arrived in the right *column*.
        """
        if self.dry_run:
            # Leave the real clipboard alone while only printing the plan; the
            # caller has already logged the column and value.
            self.writes += 1
            return

        self._trace(f"write <- {value!r}")
        clip_set(value)
        self.tap("f2")

        # Confirm again immediately before pasting.  Between the two points TME
        # has processed an F2, and pasting whatever is on the clipboard rather
        # than what we meant to write is a silent corruption.
        #
        # Comparing the content once is not enough: TME's own clipboard write
        # from the read-back before this one can still be in flight and land
        # between the check and the paste.  clip_settle waits for the clipboard
        # sequence number to stop moving first, so nothing is on its way in.
        if clip_get() != value:
            clip_set(value)
        clip_settle(value)

        self.chord("ctrl", "v")
        if commit:
            self.tap(commit)
        self.writes += 1

    def require_cell_focus(self, what: str = "a cell") -> str:
        """Confirm a cell really has keyboard focus, and return its contents.

        Reopening the dialog leaves focus outside the grid: Ctrl+I still adds a
        row, but F2 opens no editor and every later key lands somewhere
        unpredictable.  Establishing focus up front turns that into a clear
        instruction instead of a scrambled table.
        """
        try:
            return self.read_cell().value
        except TmeIoError as exc:
            raise FocusLost(
                f"could not read {what}, so the Schematic Table does not have "
                f"keyboard focus. Click once inside the table, then run again.\n"
                f"  underlying: {exc}"
            ) from exc

    # -- misc --------------------------------------------------------------

    def _trace(self, message: str) -> None:
        if self.verbose or self.dry_run:
            print(f"    [{'dry' if self.dry_run else 'key'}] {message}")

    def stats(self) -> str:
        line = f"{self.reads} read(s), {self.writes} write(s); {self.throttle.summary()}"
        if self.saves_dismissed:
            line += f"; answered No to {self.saves_dismissed} save prompt(s)"
        return line
