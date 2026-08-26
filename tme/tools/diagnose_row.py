"""Replay one circuit row's write/verify sequence with every key logged.

Written to work out why DB-Y1-GF-WH2-1 L-R6 comes out with its Cable
Specification holding the circuit number and its Terminal Load holding the cable
-- the same way on every run, with no dropped keys (0 backoffs, 0 ms latency),
so it is not a timing race.

It replays onto whatever row currently has focus, which is meant to be a row that
is already wrong and due to be deleted: nothing here is safe to point at a row
you want to keep.

    python diagnose_row.py "drawings/BLD_Y1-schedule.xlsx" DB-Y1-GF-WH2-1 L-R6
"""

from __future__ import annotations

import sys

from tme import grid
from tme import win as tmeio
from tme import config
from tme.schedule.reader import load_boards


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 3:
        print(__doc__, file=sys.stderr)
        return 2
    workbook, board_name, circuit_number = argv

    boards, _ = load_boards(workbook)
    board = next((b for b in boards if b.name == board_name), None)
    if board is None:
        print(f"no board {board_name!r}", file=sys.stderr)
        return 2
    circuit = next(
        (c for c in board.circuits if c.circuit_number == circuit_number), None
    )
    if circuit is None:
        print(f"no circuit {circuit_number!r} in {board_name!r}", file=sys.stderr)
        return 2

    print(f"replaying {board_name} {circuit_number}")
    print(f"  writes: {circuit.write}")
    print(f"  expects: {dict(zip(config.COLUMN_KEYS, circuit.expected_vector()))}\n")

    with tmeio.TmeSession(verbose=True) as session:
        table = grid.SchematicGrid(session)
        table.board_name = board_name

        print("-- where are we now? reading the focused cell --")
        try:
            print(f"   focused cell holds {session.read_cell().value!r}\n")
        except tmeio.TmeIoError as exc:
            print(f"   could not read: {exc}\n")
            return 1

        print("-- sweeping the row as it stands, column by column --")
        seen: dict[str, str | None] = {}
        table.home()
        for index, key in enumerate(config.COLUMN_KEYS):
            if key in config.NEVER_F2_KEYS:
                # F2 here raises a TME validation dialog -- "Please enter an
                # integer between [10,5000]" on Conduit Size, "Elevation input
                # error" on Elevation -- which then blocks everything after it.
                print(f"   col {index} {config.COLUMN_LABELS[key]:<22} = (not read: F2 is unsafe here)")
                if index < config.N_COLUMNS - 1:
                    session.tap("right")
                continue
            try:
                value = session.read_cell().value
            except tmeio.CellReadTimeout:
                value = None
            seen[key] = value
            print(f"   col {index} {config.COLUMN_LABELS[key]:<22} = {value!r}")
            if index < config.N_COLUMNS - 1:
                session.tap("right")

        # Refuse to write unless this really is the broken row.  Replaying onto a
        # row that is already correct would damage verified work.
        if seen.get("circuit_number") != circuit_number:
            print(
                f"\nrefusing to write: the focused row's Circuit Number reads "
                f"{seen.get('circuit_number')!r}, not {circuit_number!r}"
            )
            return 1

        print("\n-- replaying write_row --")
        table.write_row(circuit)

        print("\n-- replaying verify_row --")
        for mismatch in table.verify_row(circuit):
            print(f"   MISMATCH {mismatch}")

        print(f"\n{session.stats()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
