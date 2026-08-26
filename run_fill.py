"""Fill the TME Schematic Table from the approved Excel schedule.

Preconditions, all checked before a key is sent:
  * Cubicost TME is open with the System Diagram dialog showing
  * the Schematic Table has keyboard focus -- click once inside it
  * the boards named in the Excel do not already exist in this project

The run stops at the first thing it cannot account for: a cell that does not read
back, a row that landed in the wrong column, a previous row overwritten, a popup
stealing focus, or F12.  Stopping loses nothing that was already verified, but
note that TME commits rows to the project as they are created and Cancel does
*not* undo them -- a partly filled board has to be deleted by hand before that
board is retried, or TME will stop to ask about the repeated name.

The OK button is never pressed.  Review the table, then commit it yourself.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import grid
from tme import win as tmeio
from tme import config
from tme.schedule.reader import Board, ExcelStructureError, load_boards


def preflight(boards: list[Board], warnings: list[str], force: bool) -> bool:
    """Report what is about to happen, and refuse on unresolved warnings."""
    total = sum(len(b.circuits) for b in boards)
    print(f"about to fill {len(boards)} board(s), {total} circuit(s):")
    for board in boards:
        print(f"  {board.name}: {len(board.circuits)} circuits")
    if warnings:
        print(f"\n{len(warnings)} warning(s) from the workbook:")
        for warning in warnings:
            print(f"  ! {warning}")
        if not force:
            print("\nrefusing to start; fix the workbook or pass --force")
            return False
    return True


def save_state(path: Path, payload: dict[str, object]) -> None:
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), "utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", default=config.DEFAULT_WORKBOOK)
    parser.add_argument("--dry-run", action="store_true",
                        help="print the key sequence without sending anything")
    parser.add_argument("--start-at", default=None, metavar="BOARD",
                        help="skip every board before this one, for resuming "
                             "after a stop. The skipped boards must already be "
                             "in the TME project, or a feeder pointing at one "
                             "will have nothing to link to.")
    parser.add_argument("--limit-boards", type=int, default=None)
    parser.add_argument("--limit-rows", type=int, default=None,
                        help="stop after this many circuits per board")
    parser.add_argument("--verbose", action="store_true",
                        help="log every key sent and every value read back, for "
                             "working out where a row went wrong")
    parser.add_argument("--force", action="store_true",
                        help="start even though the workbook has warnings")
    parser.add_argument("--auto-dismiss-save", action="store_true",
                        help="answer No to TME's 'Save current project?' prompt "
                             "and carry on, instead of stopping for it. Only that "
                             "one dialog, identified by picture; anything else "
                             "still stops the run.")
    args = parser.parse_args(argv)

    try:
        boards, warnings = load_boards(args.workbook)
    except (ExcelStructureError, FileNotFoundError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if args.start_at is not None:
        names = [b.name for b in boards]
        if args.start_at not in names:
            print(f"ERROR: no board named {args.start_at!r} in the workbook; "
                  f"it has {', '.join(names)}", file=sys.stderr)
            return 2
        skipped = names.index(args.start_at)
        boards = boards[skipped:]
        print(f"skipping the first {skipped} board(s); they must already be in "
              f"the project\n")

    if args.limit_boards is not None:
        boards = boards[: args.limit_boards]
    if args.limit_rows is not None:
        for board in boards:
            board.circuits = board.circuits[: args.limit_rows]

    if not preflight(boards, warnings, args.force):
        return 2

    print(f"\nabort at any time with {config.ABORT_KEY_NAME}\n")

    state_path = Path(config.STATE_FILE)
    done: list[str] = []
    started = time.monotonic()
    rows_done = 0

    def on_row(board: Board, circuit: object, index: int) -> None:
        nonlocal rows_done
        rows_done += 1
        number = getattr(circuit, "circuit_number", "?")
        print(f"    {board.name} [{index + 1}/{len(board.circuits)}] {number} ok")
        save_state(state_path, {
            "workbook": args.workbook,
            "boards_done": done,
            "current_board": board.name,
            "rows_done_in_board": index + 1,
            "rows_total_in_board": len(board.circuits),
        })

    exit_code = 0
    if args.auto_dismiss_save:
        print("save prompts will be answered No automatically; every other "
              "dialog still stops the run\n")

    with tmeio.TmeSession(
        dry_run=args.dry_run,
        verbose=args.verbose,
        auto_dismiss_save=args.auto_dismiss_save,
    ) as session:
        table = grid.SchematicGrid(session)
        try:
            for board in boards:
                print(f"  {board.name} ({len(board.circuits)} circuits)")
                table.fill_board(board, on_row=on_row)
                done.append(board.name)
        except tmeio.Aborted as exc:
            print(f"\nABORTED: {exc}")
            exit_code = 1
        except tmeio.StrayWindow as exc:
            print(f"\nSTOPPED -- something interrupted the run:\n{exc}")
            exit_code = 1
        except (tmeio.FocusLost, tmeio.EditModeFailed) as exc:
            print(f"\nSTOPPED: {exc}")
            exit_code = 1
        except grid.RowMismatch as exc:
            print(f"\nSTOPPED -- what was written does not match the Excel:\n{exc}")
            exit_code = 1
        except tmeio.TmeIoError as exc:
            print(f"\nSTOPPED: {type(exc).__name__}: {exc}")
            exit_code = 1

        elapsed = time.monotonic() - started
        print(f"\n{rows_done} row(s) verified in {elapsed:.0f}s")
        if rows_done:
            print(f"{elapsed / rows_done:.2f}s per row")
        print(session.stats())

    if exit_code:
        print(f"\nstate written to {state_path}")
        print("Before retrying a board that was only partly filled, delete it in "
              "TME first -- rows already created stay in the project, and TME "
              "stops to ask about a repeated board name.")
    else:
        print("\nAll rows verified. Review the table, then press OK yourself to "
              "commit it -- this script never does.")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
