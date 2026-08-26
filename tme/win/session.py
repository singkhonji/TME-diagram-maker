"""TmeSession: the one object that knows how TME responds.

This is the file that changes when TME changes.  Everything it stands on --
SendInput, the clipboard protocol, window handling -- does not.
"""

from __future__ import annotations

import time
from dataclasses import dataclass

import win32con
import win32gui

from tme import config
from tme.win.clipboard import (
    MISSING,
    _is_message_box_text,
    clip_get,
    clip_set,
    clip_settle,
)
from tme.win.errors import (
    Aborted,
    CellReadTimeout,
    EditModeFailed,
    FocusLost,
    StrayWindow,
    TmeIoError,
    WindowNotFound,
)
from tme.win.input import _char_event, _key_event, _send
from tme.win.timing import AbortWatch, Throttle
from tme.win.window import (
    capture_window,
    click_in_window,
    describe_window,
    find_window,
    focus_window,
    foreground_title,
    looks_like,
)

# --- BEGIN MOVED FROM tmeio.py ---
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
