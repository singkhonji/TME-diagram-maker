"""Where the client data lives, for the tests that need it.

The drawings and the schedules built from them are real project data and are
gitignored, so a fresh clone has none of it.  Tests that need it skip with a
message rather than fail -- a clone without client data is a valid clone.

Task 4 of the layout refactor moves this data; when it does, only CLIENT_ROOT
and the two path helpers change, and every test that uses them follows
automatically.
"""

from __future__ import annotations

from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent

CLIENT_ROOT = REPO_ROOT / "projects" / "mrdiy"
SNAPSHOT_DIR = CLIENT_ROOT / ".snapshots"

SLUGS = ("Y1", "Y2", "Y3", "Y4")


def building_dir(slug: str) -> Path:
    return CLIENT_ROOT / "buildings" / slug


def drawing_path(slug: str) -> Path:
    return building_dir(slug) / "drawing.yaml"


def workbook_path(slug: str) -> Path:
    return building_dir(slug) / "schedule.xlsx"


def require(slug: str) -> tuple[Path, Path]:
    """Return (drawing, workbook) for a building, or skip the test."""
    drawing, workbook = drawing_path(slug), workbook_path(slug)
    missing = [p.name for p in (drawing, workbook) if not p.exists()]
    if missing:
        pytest.skip(f"client data not in this clone: {', '.join(missing)}")
    return drawing, workbook
