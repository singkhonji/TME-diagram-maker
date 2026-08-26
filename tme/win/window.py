"""Finding, focusing, capturing and comparing the TME windows.

The capture and compare pair exists so a dialog can be identified by its pixels
rather than its title -- TME reuses the title "OK" for more than one question,
and answering an unrecognised one would be a decision this code has no business
making.

The PIL imports are function-local inside capture_window and compare_images, as
they were in tmeio.py.  Left there deliberately: hoisting them would be an edit,
and this split is a move.
"""

from __future__ import annotations

import ctypes
import os
import tempfile
import time
from ctypes import wintypes

import win32api
import win32con
import win32gui

from tme import config
from tme.win.errors import FocusLost, WindowNotFound
from tme.win.input import _key_event, _send
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
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


