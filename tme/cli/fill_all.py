"""Fill every board in a workbook, one subprocess per board.

    python -m tme.cli.fill_all projects/<client>/buildings/<slug> [--auto-dismiss-save]

Replaces the hand-written per-building shell loops.  Two things it fixes:

The board order is read from the workbook rather than kept by hand.  It is
already correct there: build.order_boards() writes boards children-first,
because TME links a circuit to a child board by name and the child has to
exist first.  The shell version repeated those names in a second place, where
a YAML edit could silently fall out of step -- and the failure lands mid-run,
after rows that cannot be removed.

There is no shell pipeline, so the exit-status trap CLAUDE.md warns about
cannot happen: a pipeline reports the *last* command's status, which turned a
failed run into a reported success.  Here each board's exit code is read
directly and the first non-zero one stops everything.

One process per board is deliberate, not incidental: a stop never rewrites a
board that already verified, so the next attempt resumes with --from instead
of starting over.

The OK button is never pressed.  Review the table, then commit it yourself.
"""

from __future__ import annotations

import argparse
import subprocess
import sys

from tme.paths import RunPaths
from tme.schedule.reader import load_boards


def board_order(workbook) -> list[str]:
    """Board names in the order the workbook stores them: children first."""
    boards, _ = load_boards(str(workbook))
    return [b.name for b in boards]


def drop_until(boards: list[str], first: str | None) -> list[str]:
    """Everything from `first` onward, for resuming after a stop."""
    if first is None:
        return boards
    if first not in boards:
        raise ValueError(
            f"no board named {first!r} in the workbook; it has "
            f"{', '.join(boards)}"
        )
    return boards[boards.index(first):]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("target", help="building folder or workbook path")
    parser.add_argument("--from", dest="from_board", default=None,
                        metavar="BOARD",
                        help="start at this board. Everything before it must "
                             "already be in the TME project, or a feeder "
                             "pointing at one has nothing to link to.")
    parser.add_argument("--dry-run", action="store_true",
                        help="print the board order and exit without touching "
                             "TME")
    parser.add_argument("--auto-dismiss-save", action="store_true",
                        help="passed through to each board's run")
    parser.add_argument("--verbose", action="store_true",
                        help="passed through to each board's run")
    args = parser.parse_args(argv)

    try:
        paths = RunPaths.resolve(args.target)
        boards = drop_until(board_order(paths.workbook), args.from_board)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if args.dry_run:
        print(f"{len(boards)} board(s), children first:")
        for i, name in enumerate(boards, start=1):
            print(f"  {i:2d}. {name}")
        return 0

    passthrough = []
    if args.auto_dismiss_save:
        passthrough.append("--auto-dismiss-save")
    if args.verbose:
        passthrough.append("--verbose")

    print(f"{len(boards)} board(s) to fill; log: {paths.log}")
    paths.log.write_text("", "utf-8")

    for name in boards:
        print(f"  {name} ...", flush=True)
        with paths.log.open("a", encoding="utf-8") as log:
            log.write(f"=== BOARD {name} ===\n")
            log.flush()
            completed = subprocess.run(
                [sys.executable, "-m", "tme.cli.fill", str(paths.root),
                 "--start-at", name, "--limit-boards", "1", *passthrough],
                stdout=log, stderr=subprocess.STDOUT,
            )
            if completed.returncode != 0:
                log.write(f"=== STOPPED at {name} (exit "
                          f"{completed.returncode}) ===\n")
                print(f"\nSTOPPED at {name} (exit {completed.returncode})")
                print(f"see {paths.log}")
                print("Before retrying this board, delete it in TME first -- "
                      "rows already created stay in the project, and TME stops "
                      "to ask about a repeated board name.")
                return completed.returncode
            log.write(f"=== OK {name} ===\n")

    with paths.log.open("a", encoding="utf-8") as log:
        log.write(f"=== ALL {len(boards)} BOARDS DONE ===\n")
    print(f"\nAll {len(boards)} boards filled. Review the table, then press OK "
          f"yourself to commit it -- this never does.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
