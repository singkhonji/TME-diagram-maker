"""Client-level settings, read from projects/<client>/project.yaml.

Only what is agreed with the client or decided by us -- never anything read off
a drawing.  Cable specs in particular stay in each building's drawing.yaml: a
shared spec is a spec nobody re-read against the sheet, and re-reading is what
catches the ditto marks and the reversed label ordering that have each produced
a wrong reading once.

The file is optional.  Without one the built-in conventions apply, so examples/
and a fresh clone still work.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

from tme import conventions

FILENAME = "project.yaml"

# How far up from a building folder to look.  projects/<client>/buildings/<slug>
# puts project.yaml two levels above the building, but a loose workbook can sit
# anywhere, so the search is bounded rather than exact.
SEARCH_DEPTH = 4


@dataclass(frozen=True)
class ProjectConfig:
    client: str
    spare_system_fallback: str
    source: Path | None

    @classmethod
    def defaults(cls) -> "ProjectConfig":
        return cls(
            client="",
            spare_system_fallback=conventions.SPARE_SYSTEM_FALLBACK,
            source=None,
        )


def find(start) -> Path | None:
    """The nearest project.yaml at or above `start`, or None."""
    here = Path(start).resolve()
    for parent in (here, *here.parents)[:SEARCH_DEPTH + 1]:
        candidate = parent / FILENAME
        if candidate.is_file():
            return candidate
    return None


def load(start) -> ProjectConfig:
    """Read the project.yaml above a building folder, or fall back."""
    path = find(start)
    if path is None:
        return ProjectConfig.defaults()
    data = yaml.safe_load(path.read_text("utf-8")) or {}
    return ProjectConfig(
        client=str(data.get("client", "")),
        spare_system_fallback=str(
            data.get("spare_system_fallback", conventions.SPARE_SYSTEM_FALLBACK)
        ),
        source=path,
    )
