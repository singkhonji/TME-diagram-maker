"""Synthetic keyboard input through win32 SendInput.

The Schematic Table is a custom-drawn canvas with no accessibility tree, so
these structs are the only way to reach it.  Nothing here has changed since it
was written, and nothing here should.
"""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes

from tme.win.errors import TmeIoError
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
# --------------------------------------------------------------------------
# SendInput plumbing
# --------------------------------------------------------------------------

ULONG_PTR = wintypes.WPARAM


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


