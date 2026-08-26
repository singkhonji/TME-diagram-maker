"""The user32 handle, shared by every layer that calls into it.

Its own module because four of the six modules below call it -- input for
SendInput and MapVirtualKeyW, clipboard for GetClipboardSequenceNumber, timing
for GetAsyncKeyState, window for PrintWindow.  It was line 94 of tmeio.py,
inside what is now input.py, and it is the only line this split moves across a
file boundary.
"""

from __future__ import annotations

import ctypes

_user32 = ctypes.WinDLL("user32", use_last_error=True)
