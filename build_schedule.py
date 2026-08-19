"""Turn a hand-read drawing file (drawings/*.yaml) into a TME schedule workbook.

The YAML is the reviewable artefact: it is written in the same shape the drawing
is marked up in -- a *group* of circuits sharing one load description, one
breaker rating and one cable -- so a QS can check it against the sheet without
reading a spreadsheet of four hundred nearly identical rows.  This script does
the mechanical part:

  * expands each group into one row per circuit reference
  * prefixes the circuit number with its section (``L`` + ``R1`` -> ``L-R1``),
    because R1/Y1/B1 restart in every section of a board while TME names a row
    ``<board>-<circuit number>`` and would otherwise generate duplicates
  * puts a child board's name in Terminal Load, and **orders the boards so every
    child is written before the parent that points at it** -- TME cannot link to
    a board that does not exist yet
  * writes the nine columns in the exact order ``config.COLUMNS`` defines

Nothing is trusted on the way out: the finished workbook is read back through
``excel_reader.load_boards`` -- the same check ``run_fill.py`` runs -- and the
script fails rather than hand over a file that would stop the filler later.

    python build_schedule.py drawings/BLD_Y1.yaml -o drawings/BLD_Y1-schedule.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl
import yaml

import config
from excel_reader import ExcelStructureError, load_boards, print_report


class DrawingError(Exception):
    """The drawing file does not describe something this script can build."""


def circuit_number(prefix: str, ref: str) -> str:
    return f"{prefix}-{ref}" if prefix else ref


def board_rows(board: dict) -> list[list[str]]:
    """Every circuit row for one board, left to right in TME's column order."""
    rows: list[list[str]] = []
    board_system = board.get("system", "Power System")

    for section in board.get("sections", []):
        prefix = section.get("prefix", "") or ""
        section_system = section.get("system", board_system)
        section_cable = section.get("cable", "")

        for group in section.get("groups", []):
            refs = group.get("refs") or []
            if not refs:
                raise DrawingError(f"{board['name']}: a group has no refs")
            spare = bool(group.get("spare"))
            load = group.get("load", "")
            child = group.get("to", "")
            rating = group.get("rating", "")
            note = group.get("note", "")
            system = group.get("system", section_system)

            if spare:
                cable = config.SPARE_MARKER
            elif group.get("no_cable"):
                cable = ""
            else:
                cable = group.get("cable", section_cable)

            # Terminal Load is the child board's name for a feeder, and the
            # breaker rating for anything that ends at a real load.
            terminal = child or rating
            if not terminal:
                raise DrawingError(
                    f"{board['name']} {refs[0]!r}: needs either a rating or a "
                    f"'to' board for the Terminal Load column"
                )

            remarks = note if spare else (load or note)

            for ref in refs:
                rows.append([
                    config.AUTO_NAME_TOKEN,
                    circuit_number(prefix, str(ref)),
                    cable,
                    terminal,
                    config.SKIP_TOKEN,
                    "",
                    config.SKIP_TOKEN,
                    system,
                    remarks,
                ])
    return rows


def order_boards(boards: list[dict]) -> list[dict]:
    """Children first, then the boards that feed them.

    TME links a circuit to a child board by name, and the child has to be in the
    project already -- so this is a hard ordering requirement, not a nicety.
    """
    by_name = {b["name"]: b for b in boards}
    for board in boards:
        for section in board.get("sections", []):
            for group in section.get("groups", []):
                child = group.get("to")
                if child and child not in by_name:
                    raise DrawingError(
                        f"{board['name']} feeds {child!r}, which is not a board "
                        f"in this file"
                    )

    ordered: list[dict] = []
    state: dict[str, str] = {}  # name -> "visiting" | "done"

    def visit(board: dict, trail: tuple[str, ...]) -> None:
        name = board["name"]
        if state.get(name) == "done":
            return
        if state.get(name) == "visiting":
            raise DrawingError(
                f"boards feed each other in a loop: {' -> '.join(trail + (name,))}"
            )
        state[name] = "visiting"
        for section in board.get("sections", []):
            for group in section.get("groups", []):
                child = group.get("to")
                if child:
                    visit(by_name[child], trail + (name,))
        state[name] = "done"
        ordered.append(board)

    for board in boards:
        visit(board, ())
    return ordered


def build(drawing_path: str, out_path: str) -> tuple[int, int]:
    data = yaml.safe_load(Path(drawing_path).read_text("utf-8"))
    boards = data.get("boards") or []
    if not boards:
        raise DrawingError(f"{drawing_path}: no boards")

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append([label for _, label, _ in config.COLUMNS])

    circuits = 0
    for board in order_boards(boards):
        sheet.append([board["name"]] + [""] * (config.N_COLUMNS - 1))
        rows = board_rows(board)
        if not rows:
            raise DrawingError(f"{board['name']}: no circuits")
        for row in rows:
            sheet.append(row)
        circuits += len(rows)

    for index, (_, label, _) in enumerate(config.COLUMNS, start=1):
        sheet.column_dimensions[
            openpyxl.utils.get_column_letter(index)
        ].width = max(14, min(48, len(label) + 6))

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    workbook.save(out_path)
    return len(boards), circuits


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("drawing", help="drawings/*.yaml")
    parser.add_argument("-o", "--out", default=None,
                        help="output workbook (default: alongside the yaml)")
    parser.add_argument("--allow-warnings", action="store_true",
                        help="write the workbook even if the read-back warns")
    args = parser.parse_args(argv)

    out = args.out or str(Path(args.drawing).with_name(
        Path(args.drawing).stem + "-schedule.xlsx"
    ))

    try:
        n_boards, n_circuits = build(args.drawing, out)
    except (DrawingError, yaml.YAMLError, FileNotFoundError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(f"wrote {out}: {n_boards} board(s), {n_circuits} circuit(s)\n")

    # Read it back through the filler's own checker rather than trusting it.
    try:
        parsed, warnings = load_boards(out)
    except ExcelStructureError as exc:
        print(f"ERROR: the workbook just written does not parse: {exc}",
              file=sys.stderr)
        return 2

    print_report(parsed, warnings)
    if warnings and not args.allow_warnings:
        print("\nrefusing to call this done; fix the drawing file",
              file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
