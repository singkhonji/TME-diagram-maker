"""Read the approved Excel schedule into distribution boards + circuit rows.

Pure offline: never touches TME.  Run ``python excel_reader.py --report`` to see
what the filler would do before any key is sent.

Workbook shape (see Test Schedule.xlsx):
    row 1                       header, matching TME's nine columns A..I
    column A filled, B empty    a distribution board row -> starts a new block
    column A == AUTO_NAME_TOKEN a circuit row belonging to the current block
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field

import openpyxl

from tme import config, conventions


class ExcelStructureError(Exception):
    """The workbook does not match the expected layout, so refuse to run."""


@dataclass
class Circuit:
    """One circuit row: what to write, and what the whole row must read back as."""

    excel_row: int
    values: dict[str, str]          # every column key -> intended final value
    write: dict[str, str]           # subset the script actually types
    is_spare: bool

    def expected_vector(self) -> list[str]:
        """The nine cells, in TME's left-to-right order, for the row sweep."""
        return [self.values[k] for k in config.COLUMN_KEYS]

    @property
    def circuit_number(self) -> str:
        return self.values["circuit_number"]


@dataclass
class Board:
    """A distribution board and its circuits."""

    name: str
    excel_row: int
    circuits: list[Circuit] = field(default_factory=list)

    @property
    def spare_count(self) -> int:
        return sum(1 for c in self.circuits if c.is_spare)


def _cell_text(value: object) -> str:
    """Normalise a cell to a plain string; blank for None."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _resolve_circuit(
    row_idx: int, raw: dict[str, str], spare_system: str
) -> Circuit:
    """Turn one raw Excel row into intended values + the subset to type.

    ``spare_system`` is the client rule for what System a spare way takes;
    it is passed in rather than read from a global so that one workbook
    cannot be resolved two different ways depending on who called.
    """
    is_spare = raw["cable_spec"].upper() == conventions.SPARE_MARKER

    values: dict[str, str] = {}
    for key in config.COLUMN_KEYS:
        cell = raw[key]

        if key == "name":
            # TME generates this from board name + circuit number.
            values[key] = ""
            continue

        if cell == conventions.SKIP_TOKEN:
            # Explicitly "leave TME's default alone" -- still verified.
            values[key] = config.CIRCUIT_DEFAULTS[key]
            continue

        if key == "system" and not cell:
            # Spare ways have System blank in Excel; user's rule is to match the
            # rest of the board rather than leave TME's Lighting System default.
            values[key] = spare_system
            continue

        values[key] = cell

    write = {
        key: values[key]
        for key in config.WRITE_KEYS
        if values[key] and values[key] != config.CIRCUIT_DEFAULTS.get(key)
    }
    return Circuit(excel_row=row_idx, values=values, write=write, is_spare=is_spare)


def load_boards(
    path: str, *, spare_system: str | None = None
) -> tuple[list[Board], list[str]]:
    """Parse the workbook.  Returns (boards, warnings); raises on bad structure.

    ``spare_system`` defaults to the built-in convention.  Callers that know
    which client a workbook belongs to pass that client's rule instead --
    see tme.project.
    """
    if spare_system is None:
        spare_system = conventions.SPARE_SYSTEM_FALLBACK
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.worksheets[0]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()

    if not rows:
        raise ExcelStructureError(f"{path}: worksheet is empty")

    # Row 1 must line up with TME's own header, or every column mapping is wrong.
    header = [_cell_text(v) for v in rows[0][: config.N_COLUMNS]]
    expected_header = [label for _, label, _ in config.COLUMNS]
    if header != expected_header:
        raise ExcelStructureError(
            f"{path}: header row does not match TME columns.\n"
            f"  expected: {expected_header}\n"
            f"  found:    {header}"
        )

    boards: list[Board] = []
    warnings: list[str] = []

    for offset, row in enumerate(rows[1:], start=2):
        cells = list(row) + [None] * (config.N_COLUMNS - len(row))
        raw = {key: _cell_text(cells[i]) for i, key in enumerate(config.COLUMN_KEYS)}

        if not any(raw.values()):
            continue  # blank spacer row

        is_board_row = bool(raw["name"]) and raw["name"] != conventions.AUTO_NAME_TOKEN \
            and not raw["circuit_number"]

        if is_board_row:
            boards.append(Board(name=raw["name"], excel_row=offset))
            continue

        if not boards:
            raise ExcelStructureError(
                f"{path}: row {offset} is a circuit but no distribution board "
                f"row came before it"
            )

        if not raw["circuit_number"]:
            warnings.append(f"row {offset}: circuit has no Circuit Number -- skipped")
            continue

        circuit = _resolve_circuit(offset, raw, spare_system)

        # A tab or newline inside a value would silently break clipboard paste
        # (TME would treat it as a cell/row separator).
        for key, value in circuit.write.items():
            if "\t" in value or "\n" in value or "\r" in value:
                warnings.append(
                    f"row {offset} {config.COLUMN_LABELS[key]}: value contains a "
                    f"tab/newline, which would break paste -- fix the Excel cell"
                )

        boards[-1].circuits.append(circuit)

    # TME rejects a duplicate board name with a modal asking to Merge or create
    # New, which stops the run dead.  Catch it here, before any key is sent.
    seen_boards: dict[str, int] = {}
    for board in boards:
        if board.name in seen_boards:
            warnings.append(
                f"board name {board.name!r} appears twice (rows "
                f"{seen_boards[board.name]} and {board.excel_row}); TME will "
                f"stop and ask whether to merge them"
            )
        seen_boards[board.name] = board.excel_row

    for board in boards:
        if not board.circuits:
            warnings.append(
                f"board {board.name!r} (row {board.excel_row}) has no circuits"
            )
        seen: dict[str, int] = {}
        for circuit in board.circuits:
            number = circuit.circuit_number
            if number in seen:
                warnings.append(
                    f"board {board.name!r}: Circuit Number {number!r} appears "
                    f"twice (rows {seen[number]} and {circuit.excel_row})"
                )
            seen[number] = circuit.excel_row

    if not boards:
        raise ExcelStructureError(f"{path}: no distribution board rows found")

    return boards, warnings


def print_report(boards: list[Board], warnings: list[str]) -> None:
    total = sum(len(b.circuits) for b in boards)
    print(f"{len(boards)} distribution board(s), {total} circuit(s) total\n")

    for board in boards:
        normal = len(board.circuits) - board.spare_count
        print(
            f"  {board.name}  (Excel row {board.excel_row})"
            f"  -> {len(board.circuits)} circuits"
            f"  [{normal} normal + {board.spare_count} SPARE]"
        )
        if board.circuits:
            first, last = board.circuits[0], board.circuits[-1]
            print(f"      {first.circuit_number} .. {last.circuit_number}")
            written = ", ".join(config.COLUMN_LABELS[k] for k in first.write)
            print(f"      writes: {written}")

    print()
    if warnings:
        print(f"{len(warnings)} warning(s):")
        for warning in warnings:
            print(f"  ! {warning}")
    else:
        print("no warnings")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="a built schedule .xlsx")
    parser.add_argument("--report", action="store_true", help="print the summary")
    parser.add_argument(
        "--dump", action="store_true", help="print every row's expected nine cells"
    )
    args = parser.parse_args(argv)

    try:
        boards, warnings = load_boards(args.workbook)
    except (ExcelStructureError, FileNotFoundError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if args.dump:
        for board in boards:
            print(f"\n=== {board.name} ===")
            for circuit in board.circuits:
                print(f"  row {circuit.excel_row}: {circuit.expected_vector()}")
        print()

    print_report(boards, warnings)
    return 1 if warnings else 0


if __name__ == "__main__":
    raise SystemExit(main())
