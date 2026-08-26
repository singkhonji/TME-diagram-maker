"""What the TME I/O layer raises.

Its own module because every layer below raises from it -- input, clipboard and
window all do -- and session imports all three.  Anywhere else is a cycle.
"""

from __future__ import annotations

# --- BEGIN MOVED FROM tmeio.py ---
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


