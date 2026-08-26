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

# Pre-refactor layout: drawings/BLD_Y1.yaml, drawings/BLD_Y1-schedule.xlsx
CLIENT_ROOT = REPO_ROOT / "drawings"
SNAPSHOT_DIR = CLIENT_ROOT / ".snapshots"

SLUGS = ("Y1", "Y2", "Y3", "Y4")


def drawing_path(slug: str) -> Path:
    return CLIENT_ROOT / f"BLD_{slug}.yaml"


def workbook_path(slug: str) -> Path:
    return CLIENT_ROOT / f"BLD_{slug}-schedule.xlsx"


def require(slug: str) -> tuple[Path, Path]:
    """Return (drawing, workbook) for a building, or skip the test."""
    drawing, workbook = drawing_path(slug), workbook_path(slug)
    missing = [p.name for p in (drawing, workbook) if not p.exists()]
    if missing:
        pytest.skip(f"client data not in this clone: {', '.join(missing)}")
    return drawing, workbook
