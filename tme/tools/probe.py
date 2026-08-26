"""Phase 0: measure how the TME Schematic Table actually responds to keys.

Nothing about a custom-drawn grid can be looked up -- Tab may or may not move
right, F2 may or may not select the old text, the System column may or may not
accept a paste.  Guessing any of it produces a filler that silently writes into
the wrong column, so each behaviour is established here on a scratch project and
recorded in ``probe_results.json`` for the real filler to read.

Run it with the System Diagram dialog open on a **throwaway** project.  It types
junk rows; close the dialog with Cancel afterwards so nothing is committed.

Established so far (2026-08-12), kept below as regression checks:
  * Ctrl+I adds a board row and focus lands on its Name cell (EPower-N)
  * a cell read needs F2 *and* Ctrl+A -- F2 opens the editor without selecting,
    so Ctrl+C alone copies nothing
  * Enter commits and leaves focus on the same cell
  * Tab moves exactly one column right; Down enters the circuit row;
    Ctrl+Left returns to the first column
  * all five target columns accept a paste, System's dropdown included
  * the circuit Name is generated as "<board name>-<circuit number>"
  * an empty cell answers a Ctrl+C with silence, indistinguishable from a
    dropped copy, so emptiness can never be positively verified
  * a tab-separated paste does NOT spread across cells -- tabs become spaces
  * a bare Ctrl+C is not a row copy; it returns nothing
  * a bare Esc cancels the whole dialog, discarding every row
  * the grid scrolls itself: writing and reading worked 28 rows past the viewport
"""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path
from typing import Callable

from tme import win as tmeio
from tme import config

SCREENSHOT_DIR = Path("probe_shots")

ANCHOR_ELEMENT = config.CIRCUIT_DEFAULTS["corresponding_element"]
ANCHOR_ELEVATION = config.CIRCUIT_DEFAULTS["elevation"]
ANCHOR_SYSTEM = config.CIRCUIT_DEFAULTS["system"]

IDX = config.COLUMN_INDEX


class Probe:
    """Runs the steps, collects findings, and never lets one failure end the run."""

    def __init__(self, session: tmeio.TmeSession, shots: bool = True) -> None:
        self.s = session
        self.shots = shots
        self.findings: dict[str, object] = {}
        self.failures: list[str] = []
        # A scratch project accumulates boards across runs, and TME stops to ask
        # about a repeated board name -- so each run picks its own.
        self.board_name = f"PROBE_{time.strftime('%H%M%S')}"
        if shots:
            SCREENSHOT_DIR.mkdir(exist_ok=True)

    # -- helpers -----------------------------------------------------------

    def note(self, key: str, value: object) -> None:
        self.findings[key] = value
        print(f"      -> {key} = {value!r}")

    def shot(self, name: str) -> None:
        if not self.shots:
            return
        try:
            from PIL import ImageGrab

            ImageGrab.grab().save(SCREENSHOT_DIR / f"{name}.png")
        except Exception as exc:
            print(f"      (screenshot {name} failed: {exc})")

    def cell(self, label: str, *, expect_empty: bool = False) -> str | None:
        """Read the focused cell.  None means TME stayed silent."""
        try:
            got = self.s.read_cell(expect_empty=expect_empty)
        except tmeio.EditModeFailed as exc:
            print(f"      {label}: EDIT MODE FAILED ({exc})")
            return None
        except tmeio.CellReadTimeout:
            note = "as expected for an empty cell" if expect_empty else "NO ANSWER"
            print(f"      {label}: silent ({note})")
            return None
        print(f"      {label}: {got.value!r}  [{got.latency_s * 1000:.0f} ms]")
        return got.value

    def sweep(self, label: str) -> list[str | None]:
        """Read the verifiable columns of the focused row, left to right.

        Anchors at column one first, so the sweep always describes the row it
        was asked about rather than wherever focus happened to be.  Stops after
        the ninth column: Tab on Remarks wraps to the row below.  Columns that
        must never see an F2 are tabbed past and reported as None.
        """
        print(f"      {label}:")
        self.home()
        cells: list[str | None] = []
        for i, key in enumerate(config.COLUMN_KEYS):
            if key in config.NEVER_F2_KEYS:
                cells.append(None)
            else:
                label_text = config.COLUMN_LABELS[key]
                cells.append(self.cell(f"          {i + 1:>2} {label_text:<22}"))
            if i < config.N_COLUMNS - 1:
                self.s.tap("tab")
        self.home()
        return cells

    def home(self) -> None:
        self.s.chord("ctrl", "left")

    def step(self, name: str, fn: Callable[[], None]) -> None:
        print(f"\n--- {name} ---")
        try:
            fn()
        except (tmeio.Aborted, tmeio.StrayWindow, tmeio.FocusLost):
            raise
        except Exception as exc:
            self.failures.append(f"{name}: {type(exc).__name__}: {exc}")
            print(f"      STEP FAILED: {type(exc).__name__}: {exc}")

    # -- steps -------------------------------------------------------------

    def v0_setup(self) -> None:
        """Regression: Ctrl+I, focus on the board Name, and overwrite it."""
        self.shot("00_before")
        self.s.chord("ctrl", "i")
        time.sleep(0.3)  # only here: the grid may be empty, nothing to poll on yet

        # Ctrl+I puts focus on the new board's row but keeps whatever *column*
        # was current, so the Name cell has to be navigated to explicitly.
        self.home()

        # Then establish focus before anything else: if the grid never had
        # keyboard focus, Ctrl+I still adds a row but F2 opens no editor and
        # every later key lands somewhere unpredictable.
        name = self.s.require_cell_focus("the new board's Name cell")
        print(f"      board Name after Ctrl+I: {name!r}")
        self.note("ctrl_i_focus_value", name)
        self.note(
            "ctrl_i_focus_is_board_name",
            name.startswith(config.DB_DEFAULT_NAME_PREFIX),
        )
        # Unique per run: TME stops and asks whether to merge if a board name
        # already exists, which would end the probe on the second run.
        self.s.write_cell(self.board_name, commit="enter")
        after = self.cell("board Name after write")
        self.note("board_name_writable", after == self.board_name)
        self.shot("01_board_named")

    def v1_paste_modes(self) -> None:
        """V1: F2 alone must select the old text, so no Ctrl+A is ever needed.

        This is the linchpin of the safety design.  A Ctrl+A that landed outside
        an editor would select every row, and the paste after it would overwrite
        rows already verified -- exactly what produced TME's "Please select only
        1 distribution box" prompt during an earlier run.  If F2 selects for us,
        that key never has to be sent at all.
        """
        self.s.tap("down")
        self.note("circuit_row_reached", True)
        self.s.tap("tab")  # Circuit Number
        blank = self.cell("fresh Circuit Number", expect_empty=True)
        self.note("fresh_cell_is_silent", blank is None)

        self.s.write_cell("FIRSTVAL", commit="enter")
        first = self.cell("after pasting into the empty cell")
        self.note("paste_into_empty_cell", first)
        self.note("empty_cell_paste_works", first == "FIRSTVAL")

        self.s.write_cell("SECONDVAL", commit="enter")
        second = self.cell("after pasting over it")
        self.note("overwrite_result", second)
        self.note("f2_selects_all_so_paste_replaces", second == "SECONDVAL")
        self.note("paste_appended_instead", bool(second and second != "SECONDVAL"))
        self.shot("02_paste_modes")

    def v2_tab_wraps_at_remarks(self) -> None:
        """V2: Tab on Remarks moves to the row below -- does it *create* that row?

        If it creates one, a row can be filled and advanced in a single pass with
        no Ctrl+R at all.  If it only moves onto an existing row, then Tab past
        the last row must be avoided.
        """
        self.home()
        # Fill a full row so every column is identifiable afterwards.
        values = {
            "circuit_number": "TABW1",
            "cable_spec": "TAB WRAP SPEC",
            "terminal_load": "20A",
            "system": "Power System",
            "remarks": "TAB WRAP REMARK",
        }
        for i, key in enumerate(config.COLUMN_KEYS):
            if key in values:
                self.s.write_cell(values[key], commit="enter")
            if i < config.N_COLUMNS - 1:
                self.s.tap("tab")

        # Focus is on Remarks (column 9).  One more Tab is the wrap.
        self.note("row_before_wrap", self.sweep("row before the wrap"))

        # Walk back out to Remarks, then Tab once.
        for _ in range(config.N_COLUMNS - 1):
            self.s.tap("tab")
        self.s.tap("tab")
        landed = self.cell("cell focused after Tab past Remarks", expect_empty=True)
        self.note("tab_past_remarks_landed_on", landed)

        self.home()
        wrapped = self.sweep("row after the wrap")
        self.note("row_after_wrap", wrapped)
        self.note(
            "tab_wrap_row_is_fresh_circuit",
            bool(wrapped)
            and wrapped[IDX["system"]] == ANCHOR_SYSTEM
            and wrapped[IDX["circuit_number"]] is None
            and wrapped[IDX["name"]] is None,
        )
        self.note(
            "tab_wrap_did_not_land_back_on_same_row",
            bool(wrapped) and wrapped[IDX["circuit_number"]] != "TABW1",
        )
        self.shot("03_tab_wrap")

    def v3_add_circuit(self) -> None:
        """V3: Ctrl+R adds a circuit row -- where does focus land, is the old row safe?"""
        self.home()
        self.s.tap("tab")
        self.s.write_cell("PREV1", commit="enter")
        before = self.cell("previous row Circuit Number")

        self.s.chord("ctrl", "r")
        landed = self.cell("cell focused after Ctrl+R", expect_empty=True)
        self.note("ctrl_r_focus_value", landed)

        self.home()
        fresh = self.sweep("row after Ctrl+R")
        self.note("ctrl_r_new_row", fresh)
        self.note(
            "ctrl_r_new_row_has_defaults",
            bool(fresh) and fresh[IDX["system"]] == ANCHOR_SYSTEM,
        )
        self.note(
            "ctrl_r_new_row_is_empty",
            bool(fresh) and fresh[IDX["circuit_number"]] is None,
        )

        # The check that catches a swallowed Ctrl+R during the real run: without
        # it, row k's values would overwrite row k-1 and still read back clean.
        self.s.tap("up")
        self.home()
        self.s.tap("tab")
        above = self.cell("previous row Circuit Number after Up")
        self.note("previous_row_intact_after_ctrl_r", above == before)
        self.note("up_moves_one_row", above == before)
        self.shot("04_after_ctrl_r")

    def v4_second_board(self) -> None:
        """V4: a second Ctrl+I makes a new top-level board, not a nested one."""
        self.s.chord("ctrl", "i")
        time.sleep(0.3)
        self.home()
        name = self.cell("Name after the second Ctrl+I")
        self.note("second_ctrl_i_focus_value", name)
        self.note(
            "second_ctrl_i_makes_new_board",
            bool(name and name.startswith(config.DB_DEFAULT_NAME_PREFIX)),
        )
        self.shot("05_second_board")

    # -- driver ------------------------------------------------------------

    def run(self, only: set[str] | None = None) -> None:
        steps: list[tuple[str, Callable[[], None]]] = [
            ("V0 board setup", self.v0_setup),
            ("V1 paste modes (no Ctrl+A anywhere)", self.v1_paste_modes),
            ("V2 Tab wrapping at Remarks", self.v2_tab_wraps_at_remarks),
            ("V3 Ctrl+R adds a circuit row", self.v3_add_circuit),
            ("V4 second board", self.v4_second_board),
        ]
        for name, fn in steps:
            code = name.split()[0]
            if only and code not in only:
                continue
            self.step(name, fn)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--only", help="comma-separated step codes, e.g. V1,V2")
    parser.add_argument("--no-shots", action="store_true")
    parser.add_argument("--out", default=config.PROBE_RESULTS_FILE)
    args = parser.parse_args(argv)

    only = set(args.only.upper().split(",")) if args.only else None

    print("Probe: measuring TME Schematic Table key behaviour")
    print(f"Abort at any time with {config.ABORT_KEY_NAME}.")
    print("This types junk rows -- close the dialog with Cancel when done.\n")

    with tmeio.TmeSession() as session:
        probe = Probe(session, shots=not args.no_shots)
        try:
            probe.run(only)
        except tmeio.Aborted as exc:
            print(f"\nABORTED: {exc}")
        except (tmeio.StrayWindow, tmeio.FocusLost) as exc:
            print(f"\nSTOPPED: {exc}")
        print(f"\n{session.stats()}")

    payload = {
        "findings": probe.findings,
        "failures": probe.failures,
        "throttle": {
            "ewma_ms": round(session.throttle.ewma_s * 1000),
            "peak_ms": round(session.throttle.max_latency_s * 1000),
            "backoffs": session.throttle.penalties,
        },
    }
    Path(args.out).write_text(json.dumps(payload, indent=2, ensure_ascii=False), "utf-8")
    print(f"\nwrote {args.out}")

    if probe.failures:
        print(f"\n{len(probe.failures)} step(s) failed:")
        for failure in probe.failures:
            print(f"  ! {failure}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
