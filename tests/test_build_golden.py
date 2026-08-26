"""The schedule builder still produces the schedules we already approved.

Each committed BLD_*-schedule.xlsx was reviewed by hand before it was fed to
TME.  Rebuilding it from its YAML and comparing cell for cell is the only
regression net this pipeline has, and it covers everything from the YAML parse
through order_boards() to the workbook write.
"""

from __future__ import annotations

import openpyxl
import pytest

from tests.clientdata import SLUGS, require

import build_schedule


def cell_grid(path) -> list[list[str]]:
    """Every cell as text, so a stored 20 and a stored '20' compare equal."""
    sheet = openpyxl.load_workbook(path).active
    return [
        ["" if cell.value is None else str(cell.value) for cell in row]
        for row in sheet.iter_rows()
    ]


@pytest.mark.parametrize("slug", SLUGS)
def test_build_matches_the_approved_schedule(slug: str, tmp_path) -> None:
    drawing, approved = require(slug)
    rebuilt = tmp_path / "rebuilt.xlsx"

    build_schedule.build(str(drawing), str(rebuilt))

    assert cell_grid(rebuilt) == cell_grid(approved)
