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


def test_a_blank_system_takes_the_spare_system_passed_in() -> None:
    """The one branch no committed workbook reaches.

    build.py writes a System into every row, so reading a real schedule back
    never exercises the fallback -- a broken wiring would pass every other
    test in this file.
    """
    from tme.schedule.reader import _resolve_circuit
    from tme import config

    raw = {key: "" for key in config.COLUMN_KEYS}
    raw["circuit_number"] = "L1"
    raw["cable_spec"] = "SPARE"

    circuit = _resolve_circuit(2, raw, "Fire Alarm System")

    assert circuit.is_spare
    assert circuit.values["system"] == "Fire Alarm System"
