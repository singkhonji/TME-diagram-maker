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

The names re-exported below are the ones this package's callers actually use.
They are re-exported so ``from tme import win as tmeio`` keeps every existing
call site working unchanged -- which is what makes the split provably a move.
"""

from __future__ import annotations

from tme.win.errors import (
    Aborted,
    CellReadTimeout,
    EditModeFailed,
    FocusLost,
    StrayWindow,
    TmeIoError,
    WindowNotFound,
)
from tme.win.session import CellRead, TmeSession
from tme.win.window import capture_window, find_window, focus_window

__all__ = [
    "Aborted",
    "CellRead",
    "CellReadTimeout",
    "EditModeFailed",
    "FocusLost",
    "StrayWindow",
    "TmeIoError",
    "TmeSession",
    "WindowNotFound",
    "capture_window",
    "find_window",
    "focus_window",
]
