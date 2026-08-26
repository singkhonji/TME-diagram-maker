"""Where a run's files are, derived from the building folder it runs against.

These used to be bare filename constants in config.py, resolved against the
current working directory: .state.json and the fill log landed at the repo
root, shared between every building anyone ran.  Deriving them from the target
means two buildings can never overwrite each other's state, and a run started
from anywhere finds its own artefacts.

A run names its target.  There is no default workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from tme import project

WORKBOOK_NAME = "schedule.xlsx"
DRAWING_NAME = "drawing.yaml"
RUNS_DIR = "runs"


@dataclass(frozen=True)
class RunPaths:
    root: Path
    workbook: Path
    drawing: Path
    runs: Path

    @property
    def state(self) -> Path:
        return self.runs / "state.json"

    @property
    def log(self) -> Path:
        return self.runs / "fill.log"

    @property
    def intruder(self) -> Path:
        return self.runs / "interrupted_by.png"

    @property
    def probe(self) -> Path:
        return self.runs / "probe_results.json"

    @property
    def project(self) -> "project.ProjectConfig":
        """The client settings above this building, or the built-in defaults.

        Read here rather than by each caller so that build, fill and fill_all
        cannot resolve the same workbook under three different client rules.
        """
        return project.load(self.root)

    @classmethod
    def for_building(cls, folder) -> "RunPaths":
        """A projects/<client>/buildings/<slug>/ folder."""
        root = Path(folder)
        workbook = root / WORKBOOK_NAME
        if not workbook.exists():
            raise FileNotFoundError(
                f"{root} has no {WORKBOOK_NAME}; build it first with "
                f"python -m tme.cli.build {root}"
            )
        runs = root / RUNS_DIR
        runs.mkdir(parents=True, exist_ok=True)
        return cls(root=root, workbook=workbook, drawing=root / DRAWING_NAME,
                   runs=runs)

    @classmethod
    def for_workbook(cls, path) -> "RunPaths":
        """A named .xlsx; runs/ goes beside it.  Covers examples/."""
        workbook = Path(path)
        if not workbook.exists():
            raise FileNotFoundError(f"no workbook at {workbook}")
        root = workbook.parent
        runs = root / RUNS_DIR
        runs.mkdir(parents=True, exist_ok=True)
        return cls(root=root, workbook=workbook, drawing=root / DRAWING_NAME,
                   runs=runs)

    @classmethod
    def resolve(cls, target) -> "RunPaths":
        """Accept either form, so old invocations keep working."""
        path = Path(target)
        return cls.for_building(path) if path.is_dir() else cls.for_workbook(path)
