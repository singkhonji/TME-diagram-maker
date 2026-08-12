"""Row-level operations on the Schematic Table.

Everything here is expressed in terms of the measured behaviour recorded in
probe.py's docstring.  The three rules that shape this module:

* Navigation uses the right arrow, never Tab.  Both move exactly one column, but
  Tab on Remarks wraps onto the row below, and a wrap in the middle of a write
  would put the next value on the wrong row entirely.
* F2 is never pressed on Corresponding Element, Conduit Size or Elevation.  Those
  columns have their own editors and validators -- Conduit Size answers an empty
  commit with "Please enter an integer between [10,5000]" and Elevation with
  "Elevation input error", either of which stops the run.
* Writing a value and reading it back proves the value arrived, not that it
  arrived in the right column.  Column alignment is proven separately, by the
  Name cell: TME generates it as "<board name>-<circuit number>", so a value one
  column out of place leaves Circuit Number empty and shows up in Name.
"""

from __future__ import annotations

from dataclasses import dataclass

import config
import tmeio
from excel_reader import Board, Circuit


@dataclass
class Mismatch:
    """One cell that did not read back as intended."""

    column: str
    wanted: str
    got: str | None

    def __str__(self) -> str:
        got = "<no answer>" if self.got is None else repr(self.got)
        return f"{config.COLUMN_LABELS[self.column]}: wanted {self.wanted!r}, got {got}"


class RowMismatch(tmeio.TmeIoError):
    """A row did not read back as intended, after a retry."""


class SchematicGrid:
    """Drives one open System Diagram dialog."""

    def __init__(self, session: tmeio.TmeSession) -> None:
        self.s = session
        self.board_name = ""

    # -- navigation --------------------------------------------------------

    def home(self) -> None:
        """Move to column one of the current row."""
        self.s.chord("ctrl", "left")

    def goto(self, key: str) -> None:
        """Put focus on column ``key`` of the current row, starting from home."""
        self.home()
        for _ in range(config.COLUMN_INDEX[key]):
            self.s.tap("right")

    # -- boards ------------------------------------------------------------

    def add_board(self, name: str) -> None:
        """Add a distribution board and rename it from TME's EPower-N default.

        Ctrl+I leaves focus on the new board's row but keeps whatever column was
        current, so column one has to be navigated to before anything is read.
        """
        self.s.chord("ctrl", "i")
        self.home()

        if self.s.dry_run:
            print(f"    [dry] would name this board {name!r}")
            self.board_name = name
            return

        default = self.s.require_cell_focus("the new board's Name cell")
        if not default.startswith(config.DB_DEFAULT_NAME_PREFIX):
            raise RowMismatch(
                f"expected a new board named {config.DB_DEFAULT_NAME_PREFIX}N "
                f"after Ctrl+I, but the Name cell reads {default!r}. The grid may "
                f"not be where we think it is; stopping."
            )

        self.s.write_cell(name, commit="enter")
        got = self.s.read_cell().value
        if got != name:
            raise RowMismatch(
                f"board name did not take: wanted {name!r}, got {got!r}"
            )
        self.board_name = name

    def add_circuit(self) -> None:
        """Add a circuit row under the current board and go to its column one."""
        self.s.chord("ctrl", "r")
        self.home()

    # -- rows --------------------------------------------------------------

    def expected_name(self, circuit: Circuit) -> str:
        """What TME generates in the Name column for this circuit."""
        return f"{self.board_name}-{circuit.circuit_number}"

    def write_row(self, circuit: Circuit) -> None:
        """Fill the current row, walking left to right with the right arrow.

        Every written cell is read back before moving on.  That is not just an
        early check: System is a dropdown, and reading closes the editor that its
        commit leaves open.  Without it the next right arrow goes to the dropdown
        instead of the grid, focus never leaves the column, and the following
        value lands on top of the one just written.
        """
        self.home()
        for index, key in enumerate(config.COLUMN_KEYS):
            value = circuit.write.get(key)
            if value:
                self._write_and_confirm(key, value)
            if index < config.N_COLUMNS - 1:
                self.s.tap("right")

    def _write_and_confirm(self, key: str, value: str) -> None:
        """Write one cell and confirm it took, retrying the cell once."""
        if self.s.dry_run:
            print(f"    [dry] {config.COLUMN_LABELS[key]} <- {value!r}")
            self.s.write_cell(value, commit="enter")
            return
        for attempt in (1, 2):
            self.s.write_cell(value, commit="enter")
            got = self._read()
            if got == value:
                return
            if attempt == 2:
                raise RowMismatch(
                    f"{config.COLUMN_LABELS[key]} would not take {value!r} "
                    f"(reads {got!r} after a retry)"
                )

    def verify_row(self, circuit: Circuit) -> list[Mismatch]:
        """Read the row back and report every cell that disagrees with Excel."""
        if self.s.dry_run:
            return []

        wanted = dict(circuit.values)
        wanted["name"] = self.expected_name(circuit)

        mismatches: list[Mismatch] = []
        self.home()
        for index, key in enumerate(config.COLUMN_KEYS):
            if key in config.VERIFY_KEYS:
                expected = wanted[key]
                got = self._read(expect_empty=not expected)
                if expected:
                    if got != expected:
                        mismatches.append(Mismatch(key, expected, got))
                elif got is not None:
                    # Meant to be blank but something is in it.  The reverse --
                    # silence where a value is expected -- is caught above,
                    # because an empty cell and a dropped copy look identical.
                    mismatches.append(Mismatch(key, "", got))
            if index < config.N_COLUMNS - 1:
                self.s.tap("right")
        return mismatches

    def _read(self, *, expect_empty: bool = False) -> str | None:
        try:
            return self.s.read_cell(expect_empty=expect_empty).value
        except tmeio.CellReadTimeout:
            return None

    def fill_row(self, circuit: Circuit, *, previous: Circuit | None = None) -> None:
        """Write one circuit row, verify it, and retry once before giving up."""
        for attempt in (1, 2):
            self.write_row(circuit)
            mismatches = self.verify_row(circuit)
            if not mismatches:
                break
            if attempt == 2:
                detail = "\n    ".join(str(m) for m in mismatches)
                raise RowMismatch(
                    f"row {circuit.circuit_number} still wrong after a retry:\n"
                    f"    {detail}"
                )
            print(
                f"      retrying {circuit.circuit_number}: "
                f"{len(mismatches)} cell(s) off ({mismatches[0]})"
            )

        if previous is not None:
            self.assert_previous_intact(previous)

    def assert_previous_intact(self, previous: Circuit) -> None:
        """Confirm the row above still holds its own Circuit Number.

        This catches a swallowed Ctrl+R, where the current row's values would
        have been written over the previous row instead.  Nothing else notices:
        the row reads back exactly as just written, and only the row that
        vanished is missing.
        """
        if self.s.dry_run:
            return

        self.s.tap("up")
        self.goto("circuit_number")
        got = self._read()
        if got != previous.circuit_number:
            raise RowMismatch(
                f"the row above should still be {previous.circuit_number!r} but "
                f"reads {got!r} -- a new-row keystroke was probably missed, and "
                f"this row was written over the previous one"
            )
        self.s.tap("down")
        self.home()

    # -- boards end to end -------------------------------------------------

    def fill_board(self, board: Board, *, on_row: object = None) -> None:
        """Add a board and fill every one of its circuits."""
        self.add_board(board.name)

        previous: Circuit | None = None
        for index, circuit in enumerate(board.circuits):
            if index > 0:
                # Ctrl+I already supplies the board's first circuit row.
                self.add_circuit()
            else:
                self.s.tap("down")
                self.home()

            self.fill_row(circuit, previous=previous)
            previous = circuit
            if callable(on_row):
                on_row(board, circuit, index)
