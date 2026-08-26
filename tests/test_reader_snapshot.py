"""load_boards returns exactly what it returned before the refactor.

Board names, their order, how many circuits each has, and the warning list.
The order matters as much as the contents: build_schedule writes boards
children-first because TME cannot link a feeder to a board that does not exist
yet, and the reader must hand that order back untouched.

The snapshot holds real board names, so it lives inside the gitignored client
tree rather than in tests/.
"""

from __future__ import annotations

import json

import pytest

from tests.clientdata import SLUGS, SNAPSHOT_DIR, require

from tme.schedule import reader as excel_reader

SNAPSHOT = SNAPSHOT_DIR / "load_boards.json"


def summarise(workbook) -> dict:
    boards, warnings = excel_reader.load_boards(str(workbook))
    return {
        "boards": [
            {"name": b.name, "circuits": len(b.circuits)} for b in boards
        ],
        "warnings": list(warnings),
    }


@pytest.mark.parametrize("slug", SLUGS)
def test_reader_output_is_unchanged(slug: str) -> None:
    _, workbook = require(slug)
    if not SNAPSHOT.exists():
        pytest.skip(
            f"no snapshot at {SNAPSHOT}; generate it with the command in "
            f"Task 3 Step 2 of the layout refactor plan"
        )
    stored = json.loads(SNAPSHOT.read_text("utf-8"))
    if slug not in stored:
        pytest.skip(f"{slug} not in the snapshot")
    assert summarise(workbook) == stored[slug]
