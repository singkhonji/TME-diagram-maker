"""Delete distribution boards from the top of the Schematic Table.

Needed because nothing in TME undoes a fill: rows are committed to the project
the moment they are created, and *neither Esc nor the Cancel button removes
them*.  So a board that was only partly filled has to be deleted before that
board is retried, or TME stops to ask whether to merge the repeated name.

Deleting is also the one operation that needs the mouse.  The Del key does
nothing in this grid -- only the toolbar's "Delete Row" button works, and on a
board row it takes that board's circuits with it.

Deliberately cautious: it only clicks Delete while sitting on a row whose Circuit
Number is empty (which is what distinguishes a board row from a circuit row --
the name cannot, since real board names contain hyphens too), it stops the moment
an unexpected dialog appears, and it never deletes more boards than asked.

    python delete_boards.py          # delete the topmost board
    python delete_boards.py 5        # delete up to five, stopping when empty
"""

from __future__ import annotations

import sys
import time

import win32api
import win32con
import win32gui

from tme import win as tmeio

# Measured against the maximised System Diagram dialog.
DELETE_ROW_RATIO = (0.302, 0.085)
GRID_BODY_RATIO = (0.60, 0.60)


def _click(x: int, y: int) -> None:
    win32api.SetCursorPos((x, y))
    time.sleep(0.15)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def _click_ratio(ratio: tuple[float, float]) -> None:
    hwnd = tmeio.find_window()
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    _click(left + int((right - left) * ratio[0]), top + int((bottom - top) * ratio[1]))


def click_delete_row() -> None:
    _click_ratio(DELETE_ROW_RATIO)


def focus_grid() -> None:
    """Click the table body so the grid has keyboard focus."""
    _click_ratio(GRID_BODY_RATIO)
    time.sleep(0.3)


def popups() -> list[tuple[int, str]]:
    found: list[tuple[int, str]] = []

    def visit(hwnd: int, _: object) -> bool:
        title = win32gui.GetWindowText(hwnd)
        if win32gui.IsWindowVisible(hwnd) and title in ("Prompt", "OK"):
            found.append((hwnd, title))
        return True

    win32gui.EnumWindows(visit, None)
    return found


def to_top(session: tmeio.TmeSession, rows: int = 80) -> None:
    """Walk to the first row.  Up on the top row does nothing, so overshooting is safe."""
    for _ in range(rows):
        session.tap("up")
    session.chord("ctrl", "left")


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    limit = int(argv[0]) if argv else 1

    tmeio.focus_window(tmeio.find_window())
    focus_grid()

    with tmeio.TmeSession() as session:
        for round_number in range(1, limit + 1):
            to_top(session)
            try:
                name = session.read_cell().value
            except tmeio.CellReadTimeout:
                print(f"round {round_number}: the table is already empty")
                return 0

            session.tap("right")
            try:
                circuit_number = session.read_cell(expect_empty=True).value
            except tmeio.CellReadTimeout:
                circuit_number = ""
            session.chord("ctrl", "left")

            if circuit_number:
                print(
                    f"round {round_number}: the top row is {name!r} with Circuit "
                    f"Number {circuit_number!r}, so it is a circuit and not a "
                    f"board. Stopping rather than guessing."
                )
                return 1

            print(f"round {round_number}: deleting board {name!r}")
            click_delete_row()
            time.sleep(0.8)

            interrupted = popups()
            if interrupted:
                for hwnd, title in interrupted:
                    tmeio.capture_window(hwnd, f"delete_popup_{title}.png")
                print(f"  a dialog appeared: {interrupted} (captured); stopping")
                return 1

            focus_grid()
            to_top(session)
            try:
                after = session.read_cell().value
            except tmeio.CellReadTimeout:
                print("  deleted; the table is now empty")
                return 0
            if after == name:
                print(f"  {name!r} is still there -- nothing was deleted")
                return 1
            print(f"  deleted; the top row is now {after!r}")

        print(f"stopped after {limit} board(s), as asked")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
