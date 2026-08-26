"""Reading and writing the clipboard, which is how cells are read and written.

The settle protocol here guards a real race: TME can still be finishing its own
clipboard write when we set ours, and the paste that follows would carry the
previously read cell's text.  See the comments on CLIPBOARD_SETTLE_SAMPLES.
"""

from __future__ import annotations

import time

import win32clipboard

from tme import config
from tme.win.errors import TmeIoError
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
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


