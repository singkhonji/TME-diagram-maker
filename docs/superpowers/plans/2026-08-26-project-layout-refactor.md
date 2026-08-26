# Project Layout Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split the reusable TME engine into a `tme/` package and move all client data into a gitignored `projects/<client>/buildings/<building>/` tree, so a new client job is a new folder.

**Architecture:** Ten flat root modules become one package with four layers — `tme/win/` (Windows input, clipboard, windows, session), `tme/schedule/` (YAML → workbook → `Board` objects), `tme/cli/` (the pipeline), `tme/tools/` (diagnostics). Client data leaves the repo tree into `projects/`, and run artefacts move beside the run that produced them. The two hand-written per-board shell loops are replaced by a Python driver that reads the board order out of the workbook.

**Tech Stack:** Python 3.14, pytest 9.1, openpyxl, PyYAML, pywin32, pypdfium2, Pillow.

## Global Constraints

- **Behaviour is frozen.** No change to key sequences, timings, verification logic, or any constant's value. Every module move is imports-only; any behavioural change belongs in a separate, later commit.
- **Never press OK.** Nothing in this repo may click TME's OK button. This refactor does not add a code path that could.
- **Never guess a delay.** No `time.sleep()` is added anywhere in this plan.
- **`git mv` for every move**, so file history survives.
- **`drawings/` is gitignored** — git cannot recover a mistake there. Task 4 copies, verifies, and only then deletes.
- **Client data never reaches a commit.** Board names, cable specs, drawing filenames and schedules are real project data. Test snapshots containing them live inside the gitignored tree.
- Python 3.14 is the interpreter (`python -V` → `Python 3.14.3`). pytest 9.1.1 is already installed.
- Run every command from the repo root: `C:/Users/123/Documents/Claude/Code/TME-diagram-maker`.

## Two corrections to the spec

The spec was written before the module was read line by line. Two things in it are wrong and this plan supersedes them:

1. **`TmeIoError` and its subclasses cannot live in `session.py`.** Spec Section 2 puts them there. But `input.py` (`_send`, line 193), `clipboard.py` (lines 252, 293, 334) and `window.py` (lines 453, 608) all raise them, and `session.py` imports all three — so errors in `session.py` is an import cycle. They get their own module, `tme/win/errors.py`.

2. **`Throttle` and `AbortWatch` get their own module too.** Spec Section 2 folds them into `session.py`. In the source they occupy lines 337–436, *between* the clipboard block and the window block. Leaving them in `session.py` would make `session.py` two non-contiguous source regions and break the concatenation proof that Section 6 of the spec depends on. As `tme/win/timing.py` every one of the six new files is one contiguous slice of the original, and the proof becomes a plain `diff`.

3. **`_user32` has to be lifted out of the input block.** `_user32 = ctypes.WinDLL("user32", use_last_error=True)` is line 94, inside what becomes `input.py` — but it is called from four of the six slices: `input.py` (`MapVirtualKeyW`, `SendInput`), `clipboard.py` (`GetClipboardSequenceNumber`), `timing.py` (`GetAsyncKeyState`) and `window.py` (`GetWindowRect`, `PrintWindow`, `ReleaseDC`). It becomes a one-line module, `tme/win/native.py`, that the other four import. This is the single line that crosses a slice boundary, and the concatenation proof in Step 5 expects exactly that one difference and nothing else.

   `from PIL import Image` at original lines 481 and 519 is a *function-local* import and stays inside the moved block untouched — do not hoist it to the header.

So `tme/win/` has seven modules, not four:

| File | Original lines | Contents |
|---|---|---|
| `native.py` | 94 (lifted) | `_user32`, the one symbol four slices share |
| `errors.py` | 45–88 | `TmeIoError`, `Aborted`, `FocusLost`, `CellReadTimeout`, `EditModeFailed`, `StrayWindow`, `WindowNotFound` |
| `input.py` | 89–198 | `SendInput` structs, `_key_event`, `_char_event`, `_send` |
| `clipboard.py` | 199–336 | `MISSING`, `clip_get`, `clipboard_sequence`, `clip_settle`, `clip_set`, `_is_message_box_text` |
| `timing.py` | 337–436 | `Throttle`, `AbortWatch` |
| `window.py` | 437–613 | `find_window`, `foreground_title`, `describe_window`, `capture_window`, `compare_images`, `looks_like`, `click`, `click_in_window`, `focus_window` |
| `session.py` | 614–1011 | `CellRead`, `TmeSession` |

Lines 1–44 (module docstring and imports) become `tme/win/__init__.py`, which also re-exports the ten names other modules actually use.

## File Structure

**Created:**

| Path | Responsibility |
|---|---|
| `tests/clientdata.py` | One place naming where client data lives; Task 4 edits this file and nothing else |
| `tests/test_imports.py` | Every module imports cleanly |
| `tests/test_build_golden.py` | YAML → workbook produces the committed schedule, cell for cell |
| `tests/test_reader_snapshot.py` | `load_boards` returns unchanged boards, circuit counts and warnings |
| `tme/__init__.py`, `tme/win/*`, `tme/schedule/*`, `tme/cli/*`, `tme/tools/*` | The package |
| `tme/paths.py` | `RunPaths` — where a run's artefacts go |
| `tme/conventions.py` | `SKIP_TOKEN`, `AUTO_NAME_TOKEN`, `SPARE_MARKER` |
| `tme/cli/fill_all.py` | One subprocess per board, order read from the workbook |
| `projects/mrdiy/project.yaml` | Client-level conventions |
| `pyproject.toml` | Package metadata, so `python -m tme.cli.fill` resolves |

**Moved:** `config.py` → `tme/config.py`; `tmeio.py` → six files under `tme/win/`; `grid.py` → `tme/grid.py`; `excel_reader.py` → `tme/schedule/reader.py`; `build_schedule.py` → `tme/schedule/build.py`; `run_fill.py` → `tme/cli/fill.py`; `pdfcrop.py` → `tme/cli/pdfcrop.py`; `probe.py`, `diagnose_row.py`, `delete_boards.py` → `tme/tools/`; `save_prompt_reference.png` → `tme/tools/assets/`; `drawings/*` → `projects/mrdiy/buildings/Y*/`.

**Deleted:** `run_y3.sh`, `run_y4.sh`, `y2_runs.log`, `.state.json`.

---

### Task 1: Test scaffolding and the import smoke test

The cheapest layer and the one that catches the most likely failure of this whole refactor — a module that no longer imports.

**Files:**
- Create: `tests/__init__.py`
- Create: `tests/clientdata.py`
- Create: `tests/test_imports.py`
- Create: `pytest.ini`

**Interfaces:**
- Produces: `tests.clientdata.CLIENT_ROOT` and `SNAPSHOT_DIR` (both `pathlib.Path`), `SLUGS` (a tuple of building slugs), `drawing_path(slug)`, `workbook_path(slug)`, and `require(slug)` → `(drawing, workbook)`, which skips the test when the data is not in this clone. Tasks 2, 3, 4 and 9 all use these.

- [ ] **Step 1: Create the pytest config**

`pytest.ini`:

```ini
[pytest]
testpaths = tests
python_files = test_*.py
addopts = -q
```

- [ ] **Step 2: Create the client-data locator**

This is the single point of truth for where client data sits. Task 4 changes this file and no other test file.

`tests/__init__.py`: empty file.

`tests/clientdata.py`:

```python
"""Where the client data lives, for the tests that need it.

The drawings and the schedules built from them are real project data and are
gitignored, so a fresh clone has none of it.  Tests that need it skip with a
message rather than fail -- a clone without client data is a valid clone.

Task 4 of the layout refactor moves this data; when it does, only CLIENT_ROOT
and BUILDINGS change, and every test that uses them follows automatically.
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
```

- [ ] **Step 3: Write the failing import test**

`tests/test_imports.py`:

```python
"""Every module imports cleanly.

Trivial, and it is the test that will actually fire during the layout refactor:
a moved module whose imports were not updated fails here in a second, instead of
failing in front of TME with rows already committed.
"""

from __future__ import annotations

import importlib

import pytest

MODULES = [
    "config",
    "tmeio",
    "grid",
    "excel_reader",
    "build_schedule",
    "pdfcrop",
    "run_fill",
    "probe",
    "delete_boards",
    "diagnose_row",
]


@pytest.mark.parametrize("name", MODULES)
def test_module_imports(name: str) -> None:
    importlib.import_module(name)
```

- [ ] **Step 4: Run it**

```bash
python -m pytest tests/test_imports.py -v
```

Expected: 10 passed. If any module fails to import on the *current* code, stop and report it — that is a pre-existing break, not something this refactor introduced, and it must be understood before anything moves.

- [ ] **Step 5: Commit**

```bash
git add pytest.ini tests/__init__.py tests/clientdata.py tests/test_imports.py
git commit -m "Add the import smoke test, before anything moves"
```

---

### Task 2: Golden test — YAML to workbook

Covers `build_schedule` end to end: parsing, the children-first board ordering, row generation, and the workbook write. This is the half of the pipeline that can be tested at all.

**Files:**
- Create: `tests/test_build_golden.py`

**Interfaces:**
- Consumes: `tests.clientdata.SLUGS`, `tests.clientdata.require`
- Produces: `tests.test_build_golden.cell_grid(path)` → `list[list[str]]`, reused by nothing else but kept small on purpose.

- [ ] **Step 1: Write the failing test**

`tests/test_build_golden.py`:

```python
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
```

- [ ] **Step 2: Run it**

```bash
python -m pytest tests/test_build_golden.py -v
```

Expected: 4 passed.

**If a building fails, do not touch the test.** Two possible causes, and they are handled differently:

- The committed `.xlsx` is stale — it was built before some later change to `build_schedule.py`. Confirm by reading the diff of the failing cells. If that is what happened, regenerate that building's workbook with `python build_schedule.py drawings/BLD_<slug>.yaml drawings/BLD_<slug>-schedule.xlsx` **as its own commit, before continuing**, and note in the commit message which cells changed and why.
- `build_schedule.py` has a real bug. Stop the refactor and report it.

Either way the golden files must be trustworthy before anything moves, or the net is worthless.

- [ ] **Step 3: Commit**

```bash
git add tests/test_build_golden.py
git commit -m "Pin the four approved schedules as a golden test"
```

---

### Task 3: Snapshot test — the workbook reader

`load_boards` is what every TME-facing entry point calls to decide what to type. Its output is pinned so the move cannot change it silently.

**Files:**
- Create: `tests/test_reader_snapshot.py`
- Create: `drawings/.snapshots/load_boards.json` (generated, gitignored — it is inside `drawings/`)

**Interfaces:**
- Consumes: `tests.clientdata.SLUGS`, `tests.clientdata.require`, `tests.clientdata.SNAPSHOT_DIR`
- Produces: nothing other tasks import.

- [ ] **Step 1: Write the test**

`tests/test_reader_snapshot.py`:

```python
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

import excel_reader

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
```

- [ ] **Step 2: Generate the snapshot**

```bash
python -c "
import json, pathlib, excel_reader
out = {}
for slug in ('Y1','Y2','Y3','Y4'):
    wb = pathlib.Path('drawings') / f'BLD_{slug}-schedule.xlsx'
    boards, warnings = excel_reader.load_boards(str(wb))
    out[slug] = {'boards': [{'name': b.name, 'circuits': len(b.circuits)} for b in boards], 'warnings': list(warnings)}
d = pathlib.Path('drawings/.snapshots'); d.mkdir(exist_ok=True)
(d / 'load_boards.json').write_text(json.dumps(out, indent=2, ensure_ascii=False), 'utf-8')
print(json.dumps({k: len(v['boards']) for k, v in out.items()}, indent=2))
"
```

Expected: a count per building, and `drawings/.snapshots/load_boards.json` written.

- [ ] **Step 3: Read the snapshot before trusting it**

```bash
python -c "
import json
d = json.load(open('drawings/.snapshots/load_boards.json', encoding='utf-8'))
for slug, v in d.items():
    print(slug, len(v['boards']), 'boards', sum(b['circuits'] for b in v['boards']), 'circuits', len(v['warnings']), 'warnings')
    print('   first:', v['boards'][0]['name'], ' last:', v['boards'][-1]['name'])
"
```

Check the last board of each building is its MSB or top-level feeder — children-first ordering means the board nothing feeds comes last. If it does not, the ordering is wrong *today* and that is a bug to report, not to snapshot.

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_reader_snapshot.py -v
```

Expected: 4 passed.

- [ ] **Step 5: Confirm the snapshot is not staged**

```bash
git status --short drawings/
```

Expected: no output. `drawings/` is gitignored, so the snapshot with its real board names stays local. If anything is listed, stop — client data is about to be committed.

- [ ] **Step 6: Commit**

```bash
git add tests/test_reader_snapshot.py
git commit -m "Pin what load_boards returns for the four buildings"
```

---

### Task 4: Move client data into projects/

The one step git cannot undo. Copy, verify, then delete.

**Files:**
- Create: `projects/mrdiy/buildings/Y1..Y4/` with `drawing.yaml`, `schedule.xlsx`, `runs/`
- Modify: `.gitignore`
- Modify: `tests/clientdata.py`
- Delete (after verification): `drawings/`, `y2_runs.log`, `.state.json`

**Interfaces:**
- Consumes: `tests.clientdata` from Task 1
- Produces: the `projects/<client>/buildings/<slug>/` layout every later task assumes.

- [ ] **Step 1: Copy the data into the new layout**

Copy, do not move. The originals stay until the tests confirm the copies load identically.

```bash
for s in Y1 Y2 Y3 Y4; do
  mkdir -p "projects/mrdiy/buildings/$s/runs"
  cp "drawings/BLD_$s.yaml"           "projects/mrdiy/buildings/$s/drawing.yaml"
  cp "drawings/BLD_$s-schedule.xlsx"  "projects/mrdiy/buildings/$s/schedule.xlsx"
done
mkdir -p projects/mrdiy/.snapshots
cp drawings/.snapshots/load_boards.json projects/mrdiy/.snapshots/load_boards.json
find projects -type f | sort
```

Expected: nine files — four `drawing.yaml`, four `schedule.xlsx`, one `load_boards.json`.

- [ ] **Step 2: Update .gitignore before anything else**

Replace the `drawings/` rule with `projects/`, so the new tree is protected from its first moment.

In `.gitignore`, replace:

```
# Read drawings and the schedules built from them are real client project data.
# This remote is public -- keep them local, and see CLAUDE.md.
drawings/
```

with:

```
# Client project data: drawings read by eye, the schedules built from them, and
# everything a run against them produced.  This remote is public -- keep the
# whole tree local, and see CLAUDE.md.  One rule so a new client inherits it.
projects/
```

- [ ] **Step 3: Prove the new tree is ignored**

```bash
git status --short && git check-ignore -v projects/mrdiy/buildings/Y4/drawing.yaml
```

Expected: `git status --short` shows only `.gitignore` modified (plus the untracked `run_y3.sh`/`run_y4.sh`), and `check-ignore` reports the `projects/` rule matching. If `projects/...` appears as untracked in `git status`, stop and fix `.gitignore` before continuing.

- [ ] **Step 4: Repoint the tests at the new layout**

In `tests/clientdata.py`, replace the layout block:

```python
# Pre-refactor layout: drawings/BLD_Y1.yaml, drawings/BLD_Y1-schedule.xlsx
CLIENT_ROOT = REPO_ROOT / "drawings"
SNAPSHOT_DIR = CLIENT_ROOT / ".snapshots"

SLUGS = ("Y1", "Y2", "Y3", "Y4")


def drawing_path(slug: str) -> Path:
    return CLIENT_ROOT / f"BLD_{slug}.yaml"


def workbook_path(slug: str) -> Path:
    return CLIENT_ROOT / f"BLD_{slug}-schedule.xlsx"
```

with:

```python
CLIENT_ROOT = REPO_ROOT / "projects" / "mrdiy"
SNAPSHOT_DIR = CLIENT_ROOT / ".snapshots"

SLUGS = ("Y1", "Y2", "Y3", "Y4")


def building_dir(slug: str) -> Path:
    return CLIENT_ROOT / "buildings" / slug


def drawing_path(slug: str) -> Path:
    return building_dir(slug) / "drawing.yaml"


def workbook_path(slug: str) -> Path:
    return building_dir(slug) / "schedule.xlsx"
```

- [ ] **Step 5: Run the whole suite against the copies**

```bash
python -m pytest -v
```

Expected: 18 passed (10 imports, 4 golden, 4 snapshot). Every one of the golden and snapshot tests is now reading the copies. If any skip, the copy did not land where `clientdata.py` expects — fix the paths, do not proceed.

- [ ] **Step 6: Delete the originals**

Only now, with the suite green against the copies.

```bash
rm -rf drawings
rm -f y2_runs.log .state.json
mv y3-fill.log projects/mrdiy/buildings/Y3/runs/fill.log
mv y4-fill.log projects/mrdiy/buildings/Y4/runs/fill.log
mv interrupted_by.png projects/mrdiy/buildings/Y4/runs/interrupted_by.png
ls
```

`y2_runs.log` (725 KB) and `.state.json` are stale run artefacts from a finished building; the two per-building logs and the interruption screenshot move to the building they belong to.

Expected `ls`: no `drawings`, no `*.log`, no `.state.json`, no `interrupted_by.png`.

- [ ] **Step 7: Re-run the suite**

```bash
python -m pytest -q
```

Expected: 18 passed. This is the proof the originals were not still being read.

- [ ] **Step 8: Commit**

```bash
git add .gitignore tests/clientdata.py
git commit -m "Move client data into projects/<client>/buildings/<building>/

drawings/ was flat and its filenames carried the building, not the job, so a
second client had nowhere to go.  The gitignore rule narrows to one line that
a new client inherits without anyone remembering to add it.

Copied rather than moved, and the originals deleted only after the golden and
snapshot tests passed against the copies -- drawings/ was gitignored, so git
could not have recovered a mistake there, and the four workbooks are the
output of reading drawings by eye and are not reproducible from the repo."
```

---

### Task 5: Package skeleton and the modules that do not touch TME

`config.py` is imported by everything, so it cannot move alone. This task moves it together with the pure modules, and updates the import line in the root modules that stay behind for now.

**Files:**
- Create: `pyproject.toml`, `tme/__init__.py`, `tme/schedule/__init__.py`, `tme/cli/__init__.py`, `tme/tools/__init__.py`
- Move: `config.py` → `tme/config.py`; `excel_reader.py` → `tme/schedule/reader.py`; `build_schedule.py` → `tme/schedule/build.py`; `pdfcrop.py` → `tme/cli/pdfcrop.py`
- Modify: `tmeio.py`, `grid.py`, `run_fill.py`, `probe.py`, `diagnose_row.py`, `delete_boards.py` (import lines only)
- Modify: `tests/test_imports.py`, `tests/test_build_golden.py`, `tests/test_reader_snapshot.py`

**Interfaces:**
- Produces: `tme.config`, `tme.schedule.reader` (with `Board`, `Circuit`, `ExcelStructureError`, `load_boards`, `print_report`), `tme.schedule.build` (with `DrawingError`, `circuit_number`, `board_rows`, `order_boards`, `build`), `tme.cli.pdfcrop`.

- [ ] **Step 1: Create the package metadata**

`pyproject.toml`:

```toml
[project]
name = "tme-diagram-maker"
version = "0.1.0"
description = "Fill Cubicost TME's Schematic Table from a Single Line Diagram"
requires-python = ">=3.12"
dependencies = [
    "openpyxl>=3.1",
    "pywin32>=306",
    "pillow>=10.0",
    "pypdfium2>=4.30",
    "PyYAML>=6.0",
]

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
include = ["tme*"]

[tool.setuptools.package-data]
"tme.tools" = ["assets/*.png"]
```

- [ ] **Step 2: Create the package directories**

```bash
mkdir -p tme/schedule tme/cli tme/tools/assets tme/win
printf '"""Fill Cubicost TME'"'"'s Schematic Table from an approved Excel schedule."""\n' > tme/__init__.py
printf '"""Reading a drawing YAML into a workbook, and a workbook into Board objects."""\n' > tme/schedule/__init__.py
printf '"""Command line entry points: build, fill, fill_all, pdfcrop."""\n' > tme/cli/__init__.py
printf '"""Diagnostics for when a run has gone wrong.  Not part of the pipeline."""\n' > tme/tools/__init__.py
ls tme
```

- [ ] **Step 3: Move the four modules**

```bash
git mv config.py         tme/config.py
git mv excel_reader.py   tme/schedule/reader.py
git mv build_schedule.py tme/schedule/build.py
git mv pdfcrop.py        tme/cli/pdfcrop.py
git status --short
```

- [ ] **Step 4: Fix the imports inside the moved modules**

In `tme/schedule/reader.py`, replace `import config` with:

```python
from tme import config
```

In `tme/schedule/build.py`, replace:

```python
import config
from excel_reader import ExcelStructureError, load_boards, print_report
```

with:

```python
from tme import config
from tme.schedule.reader import ExcelStructureError, load_boards, print_report
```

`tme/cli/pdfcrop.py` imports neither and needs no change.

- [ ] **Step 5: Fix the imports in the modules still at the root**

Each of these keeps its position for now; only the import lines change.

`tmeio.py` — replace `import config` with `from tme import config`.

`grid.py` — replace:

```python
import config
import tmeio
from excel_reader import Board, Circuit
```

with:

```python
import tmeio
from tme import config
from tme.schedule.reader import Board, Circuit
```

`run_fill.py` — replace:

```python
import config
import grid
import tmeio
from excel_reader import Board, ExcelStructureError, load_boards
```

with:

```python
import grid
import tmeio
from tme import config
from tme.schedule.reader import Board, ExcelStructureError, load_boards
```

`probe.py` — replace:

```python
import config
import tmeio
```

with:

```python
import tmeio
from tme import config
```

`diagnose_row.py` — replace:

```python
import config
import grid
import tmeio
from excel_reader import load_boards
```

with:

```python
import grid
import tmeio
from tme import config
from tme.schedule.reader import load_boards
```

`delete_boards.py` imports only `tmeio` and win32; no change.

- [ ] **Step 6: Update the test module list**

In `tests/test_imports.py`, replace the `MODULES` list with:

```python
MODULES = [
    "tme.config",
    "tme.schedule.reader",
    "tme.schedule.build",
    "tme.cli.pdfcrop",
    "tmeio",
    "grid",
    "run_fill",
    "probe",
    "delete_boards",
    "diagnose_row",
]
```

In `tests/test_build_golden.py`, replace `import build_schedule` with:

```python
from tme.schedule import build as build_schedule
```

In `tests/test_reader_snapshot.py`, replace `import excel_reader` with:

```python
from tme.schedule import reader as excel_reader
```

- [ ] **Step 7: Run the suite**

```bash
python -m pytest -v
```

Expected: 18 passed. A failure here is an import that was missed — the error names the module.

- [ ] **Step 8: Check the CLI still runs**

```bash
python -m tme.schedule.build projects/mrdiy/buildings/Y4/drawing.yaml /c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad/y4-check.xlsx
```

Expected: the same board/circuit counts it printed before. This confirms `python -m` resolves the moved module, which nothing in the test suite proves.

- [ ] **Step 9: Commit**

```bash
git add -A
git commit -m "Move config and the schedule modules into a tme/ package

config.py is imported by everything, so it could not move on its own; the
modules that stay at the root for now have had their import line updated in
the same commit.  Imports only -- no other line of any moved module changed."
```

---

### Task 6: Split tmeio.py into tme/win/

The riskiest step, and the one with the strongest proof. Every new file is one contiguous slice of the original, so a `diff` settles whether anything changed.

**Files:**
- Create: `tme/win/__init__.py`, `errors.py`, `input.py`, `clipboard.py`, `timing.py`, `window.py`, `session.py`
- Delete: `tmeio.py`
- Modify: `grid.py`, `run_fill.py`, `probe.py`, `diagnose_row.py`, `delete_boards.py` (import lines only)
- Modify: `tests/test_imports.py`

**Interfaces:**
- Produces: `tme.win` re-exporting exactly the ten names other modules use — `TmeIoError`, `Aborted`, `FocusLost`, `CellReadTimeout`, `EditModeFailed`, `StrayWindow`, `TmeSession`, `find_window`, `focus_window`, `capture_window`. Consumers keep writing `tmeio.TmeSession(...)` via `from tme import win as tmeio`, so no call site changes.

- [ ] **Step 1: Keep a copy of the original to diff against**

```bash
cp tmeio.py /c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad/tmeio-original.py
wc -l /c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad/tmeio-original.py
```

Expected: `1011`.

- [ ] **Step 2: Cut the six slices mechanically**

Do not retype any of it. This script cuts the exact line ranges and writes each slice below a marker line, so Step 5 can concatenate them back and diff.

```bash
python - <<'PY'
from pathlib import Path

src = Path("tmeio.py").read_text(encoding="utf-8").splitlines(keepends=True)
MARKER = "# --- BEGIN MOVED FROM tmeio.py ---\n"

# (filename, first line, last line) -- 1-based and inclusive, as read off the file
# native.py is written by hand in Step 3: it holds the one line that crosses a
# slice boundary, so it has no slice of its own.
SLICES = [
    ("errors.py",     45,  88),
    ("input.py",      89, 198),
    ("clipboard.py", 199, 336),
    ("timing.py",    337, 436),
    ("window.py",    437, 613),
    ("session.py",   614, 1011),
]

out = Path("tme/win")
out.mkdir(parents=True, exist_ok=True)
for name, first, last in SLICES:
    body = "".join(src[first - 1:last])
    (out / name).write_text(MARKER + body, encoding="utf-8")
    print(f"{name}: lines {first}-{last} ({last - first + 1})")
PY
```

Expected: six lines confirming the ranges. Each file currently has only the marker plus its slice — no imports yet, so nothing imports. That is fixed in the next step.

- [ ] **Step 3: Add a docstring and imports above the marker in each file**

Insert each block **above** the existing `# --- BEGIN MOVED FROM tmeio.py ---` line. Nothing below the marker is touched in this task.

`tme/win/errors.py`:

```python
"""What the TME I/O layer raises.

Its own module because every layer below raises from it -- input, clipboard and
window all do -- and session imports all three.  Anywhere else is a cycle.
"""

from __future__ import annotations

# --- BEGIN MOVED FROM tmeio.py ---
```

`tme/win/native.py` — this one has no marker and no moved block; it is the single lifted line:

```python
"""The user32 handle, shared by every layer that calls into it.

Its own module because four of the six modules below call it -- input for
SendInput and MapVirtualKeyW, clipboard for GetClipboardSequenceNumber, timing
for GetAsyncKeyState, window for PrintWindow.  It was line 94 of tmeio.py,
inside what is now input.py, and it is the only line this split moves across a
file boundary.
"""

from __future__ import annotations

import ctypes

_user32 = ctypes.WinDLL("user32", use_last_error=True)
```

`tme/win/input.py`:

```python
"""Synthetic keyboard input through win32 SendInput.

The Schematic Table is a custom-drawn canvas with no accessibility tree, so
these structs are the only way to reach it.  Nothing here has changed since it
was written, and nothing here should.
"""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes

from tme.win.errors import TmeIoError
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
```

**Delete line 94 (`_user32 = ctypes.WinDLL(...)`) from inside this file's moved block** — it now lives in `native.py`. Confirm with `grep -c "WinDLL" tme/win/input.py` → `0`.

`tme/win/clipboard.py`:

```python
"""Reading and writing the clipboard, which is how cells are read and written.

The settle protocol here guards a real race: TME can still be finishing its own
clipboard write when we set ours, and the paste that follows would carry the
previously read cell's text.  See the comments on CLIPBOARD_SETTLE_SAMPLES.
"""

from __future__ import annotations

import time

import win32clipboard

from tme import config
from tme.win.errors import TmeIoError
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
```

`tme/win/timing.py`:

```python
"""Pacing and interruption: the adaptive key throttle and the F12 abort watch.

Both are loops over measured time rather than guessed time -- the throttle
derives its delay from observed latency and may only ever slow down, and the
abort watch polls a key rather than installing a hook.
"""

from __future__ import annotations

import threading
import time

from tme import config
from tme.win.errors import Aborted
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
```

No `win32api` here: `AbortWatch` polls `_user32.GetAsyncKeyState`, which sees the
key regardless of which window has focus, and needs no extra dependency.

`tme/win/window.py`:

```python
"""Finding, focusing, capturing and comparing the TME windows.

The capture and compare pair exists so a dialog can be identified by its pixels
rather than its title -- TME reuses the title "OK" for more than one question,
and answering an unrecognised one would be a decision this code has no business
making.
"""

from __future__ import annotations

import ctypes
import os
import tempfile
import time
from ctypes import wintypes

import win32api
import win32con
import win32gui

from tme import config
from tme.win.errors import FocusLost, WindowNotFound
from tme.win.input import _key_event, _send
from tme.win.native import _user32

# --- BEGIN MOVED FROM tmeio.py ---
```

`from PIL import Image` (and `ImageChops`, `ImageStat`) appear *inside*
`capture_window` and `compare_images` in the moved block. Leave them there —
hoisting them to this header would be an edit, and the diff in Step 5 would
catch it as one.

`tme/win/session.py`:

```python
"""TmeSession: the one object that knows how TME responds.

This is the file that changes when TME changes.  Everything it stands on --
SendInput, the clipboard protocol, window handling -- does not.
"""

from __future__ import annotations

import time
from dataclasses import dataclass

import win32con
import win32gui

from tme import config
from tme.win.clipboard import (
    MISSING,
    _is_message_box_text,
    clip_get,
    clip_set,
    clip_settle,
)
from tme.win.errors import (
    Aborted,
    CellReadTimeout,
    EditModeFailed,
    FocusLost,
    StrayWindow,
    TmeIoError,
    WindowNotFound,
)
from tme.win.input import _char_event, _key_event, _send
from tme.win.timing import AbortWatch, Throttle
from tme.win.window import (
    capture_window,
    click_in_window,
    describe_window,
    find_window,
    focus_window,
    foreground_title,
    looks_like,
)

# --- BEGIN MOVED FROM tmeio.py ---
```

Importing the underscore-prefixed `_send`, `_key_event`, `_char_event` and `_is_message_box_text` across modules is ugly. Leave it. Renaming them is a behavioural-surface change and belongs in a separate commit after this one is proven; doing it here would destroy the diff proof this whole task rests on.

- [ ] **Step 4: Write the package `__init__` with the original docstring**

`tme/win/__init__.py` takes lines 1–26 of the original — the module docstring, which documents the whole layer — and re-exports the ten names other modules use.

```python
"""Low-level I/O against the TME System Diagram dialog.

TME's Schematic Table is a custom-drawn canvas: UI Automation reports
``ControlType: Custom(50025)``, no Name, not keyboard focusable, so nothing can
be read or written through the accessibility tree.  The only channel that works
is synthetic input plus the clipboard.

The one rule everything here follows: **never sleep and hope**.  Every action is
followed by evidence that it landed.  A cell read polls the clipboard until TME
answers, which doubles as the pacing mechanism -- if the app stutters for 800 ms
we wait exactly 800 ms, instead of guessing 100 ms and corrupting the row or
guessing 2000 ms and wasting an hour over a full project.

Two keys are deliberately never sent:

``Ctrl+A``
    Inside a cell editor it selects the text; outside one it selects *every row*
    in the grid, and the paste that follows would overwrite rows already
    verified.  It is also unnecessary -- F2 opens the editor with the old text
    already selected, so a paste replaces it either way (measured 2026-08-12).

``Esc`` on its own
    It cancels the whole System Diagram dialog and discards every row entered so
    far.  It is only ever sent after a successful cell read, which is positive
    proof that a cell editor is open for it to close.

The names re-exported below are the ten this package's callers actually use.
They are re-exported so ``from tme import win as tmeio`` keeps every existing
call site working unchanged -- which is what makes the split provably a move.
"""

from __future__ import annotations

from tme.win.errors import (
    Aborted,
    CellReadTimeout,
    EditModeFailed,
    FocusLost,
    StrayWindow,
    TmeIoError,
    WindowNotFound,
)
from tme.win.session import CellRead, TmeSession
from tme.win.window import capture_window, find_window, focus_window

__all__ = [
    "Aborted",
    "CellRead",
    "CellReadTimeout",
    "EditModeFailed",
    "FocusLost",
    "StrayWindow",
    "TmeIoError",
    "TmeSession",
    "WindowNotFound",
    "capture_window",
    "find_window",
    "focus_window",
]
```

- [ ] **Step 5: Prove it is a move and nothing else**

Concatenate the six slices in original order, strip each file's header down to the marker, and diff against lines 45–1011 of the original.

```bash
python - <<'PY'
from pathlib import Path

SCRATCH = Path("/c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad")
MARKER = "# --- BEGIN MOVED FROM tmeio.py ---\n"
ORDER = ["errors.py", "input.py", "clipboard.py", "timing.py", "window.py", "session.py"]

original = Path("tmeio.py").read_text(encoding="utf-8").splitlines(keepends=True)[44:]

moved = []
for name in ORDER:
    lines = Path("tme/win", name).read_text(encoding="utf-8").splitlines(keepends=True)
    moved += lines[lines.index(MARKER) + 1:]

(SCRATCH / "original-body.py").write_text("".join(original), encoding="utf-8")
(SCRATCH / "moved-body.py").write_text("".join(moved), encoding="utf-8")
print(f"original {len(original)} lines, moved {len(moved)} lines")
PY
```

Then:

```bash
diff /c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad/original-body.py /c/Users/123/AppData/Local/Temp/claude/C--Users-123-Documents-Claude-Code-TME-diagram-maker/a5a9bc09-a4dc-4ced-bdc6-f74dfdcd26c6/scratchpad/moved-body.py
```

Expected: exactly one difference — the `_user32 = ctypes.WinDLL(...)` line, which moved above the marker in `input.py` (Step 3). Nothing else.

**Any other difference means something was edited that was not meant to be.** Do not reconcile it by editing the diff output. Re-cut the slice from `tmeio-original.py` and redo Step 3 for that file.

Save the diff output — the commit message in Step 9 quotes it.

- [ ] **Step 6: Delete tmeio.py**

```bash
git rm tmeio.py
```

- [ ] **Step 7: Repoint the consumers**

Each of these keeps every `tmeio.X` call site unchanged; only the import line moves.

In `grid.py`, `run_fill.py`, `probe.py`, `diagnose_row.py` and `delete_boards.py`, replace `import tmeio` with:

```python
from tme import win as tmeio
```

- [ ] **Step 8: Update the test module list and run the suite**

In `tests/test_imports.py`, replace `"tmeio"` with the six new modules:

```python
MODULES = [
    "tme.config",
    "tme.schedule.reader",
    "tme.schedule.build",
    "tme.cli.pdfcrop",
    "tme.win",
    "tme.win.errors",
    "tme.win.native",
    "tme.win.input",
    "tme.win.clipboard",
    "tme.win.timing",
    "tme.win.window",
    "tme.win.session",
    "grid",
    "run_fill",
    "probe",
    "delete_boards",
    "diagnose_row",
]
```

```bash
python -m pytest -v
```

Expected: 25 passed (17 imports, 4 golden, 4 snapshot).

- [ ] **Step 9: Commit, quoting the proof**

```bash
git add -A
git commit -m "Split tmeio.py into tme/win/, as a pure move

1011 lines doing four jobs: SendInput plumbing, the clipboard settle protocol,
window discovery and comparison, and TmeSession -- the only one of the four
that changes when TME changes.

Six files, not the four the spec named.  Errors need their own module because
input, clipboard and window all raise from it and session imports all three,
so anywhere else is a cycle.  Throttle and AbortWatch need their own because
they sit between the clipboard and window blocks in the source, and keeping
every file one contiguous slice is what makes the next paragraph possible.

No automated test can drive TME, so this was verified by concatenating the six
slices in source order and diffing against the original.  The only difference
is the _user32 = ctypes.WinDLL(...) line, which had to move above the import
marker in input.py.  Nothing else changed.

Callers do 'from tme import win as tmeio' and every tmeio.X call site is
untouched.  The cross-module underscore imports are deliberate: renaming them
here would have destroyed the diff."
```

---

### Task 7: Move grid, the CLI and the tools into the package

**Files:**
- Move: `grid.py` → `tme/grid.py`; `run_fill.py` → `tme/cli/fill.py`; `probe.py`, `diagnose_row.py`, `delete_boards.py` → `tme/tools/`; `save_prompt_reference.png` → `tme/tools/assets/`
- Create: `tme/cli/build.py`
- Modify: `tme/win/window.py` (resource lookup for the reference image)
- Modify: `tests/test_imports.py`

**Interfaces:**
- Consumes: `tme.win`, `tme.config`, `tme.schedule.reader` from Tasks 5 and 6
- Produces: `tme.grid.SchematicGrid`, `tme.grid.RowMismatch`, `tme.cli.fill.main(argv)`, `tme.cli.build.main(argv)`.

- [ ] **Step 1: Move the files**

```bash
git mv grid.py          tme/grid.py
git mv run_fill.py      tme/cli/fill.py
git mv probe.py         tme/tools/probe.py
git mv diagnose_row.py  tme/tools/diagnose_row.py
git mv delete_boards.py tme/tools/delete_boards.py
git mv save_prompt_reference.png tme/tools/assets/save_prompt_reference.png
git status --short
```

- [ ] **Step 2: Fix the imports in the moved modules**

`tme/grid.py` — replace:

```python
from tme import win as tmeio
from tme import config
from tme.schedule.reader import Board, Circuit
```

with:

```python
from tme import config
from tme import win as tmeio
from tme.schedule.reader import Board, Circuit
```

(Alphabetical only — the content is already correct from Tasks 5 and 6.)

`tme/cli/fill.py` — replace:

```python
import grid
from tme import win as tmeio
from tme import config
from tme.schedule.reader import Board, ExcelStructureError, load_boards
```

with:

```python
from tme import config
from tme import grid
from tme import win as tmeio
from tme.schedule.reader import Board, ExcelStructureError, load_boards
```

`tme/tools/diagnose_row.py` — replace `import grid` with `from tme import grid`.

`tme/tools/probe.py` and `tme/tools/delete_boards.py` need no further change; their imports were fixed in Tasks 5 and 6.

- [ ] **Step 3: Load the reference image as a package resource**

The filename constant becomes a resource lookup, so a run started from a building folder still finds it. Today it resolves against the working directory and would fail mid-run, when the save prompt appears.

In `tme/win/window.py`, add to the import block above the marker:

```python
from importlib import resources
```

and add this function above the marker:

```python
def save_prompt_reference() -> str:
    """Filesystem path to the stored picture of TME's "Save current project?".

    A package resource, not a working-directory-relative filename: a run
    started from a building folder must still find it, and the alternative
    failure mode is a FileNotFoundError raised in the middle of a fill, when
    the prompt appears -- not at startup, where it would be harmless.
    """
    ref = resources.files("tme.tools") / "assets" / "save_prompt_reference.png"
    return str(ref)
```

- [ ] **Step 4: Use it at the call site**

In `tme/win/session.py`, find the `looks_like(` call (original line 746) and replace the argument that reads `config.SAVE_PROMPT_REFERENCE` with `save_prompt_reference()`. Add `save_prompt_reference` to the `from tme.win.window import (...)` block at the top of the file.

This is the first intentional behavioural change in the refactor, and it is confined to one argument. Note it in the commit message.

- [ ] **Step 5: Add the build CLI wrapper**

`tme/cli/build.py`:

```python
"""Build a schedule workbook from a drawing YAML.

    python -m tme.cli.build projects/<client>/buildings/<slug>

A thin wrapper so every entry point lives under tme.cli, including the one that
never touches TME.
"""

from __future__ import annotations

from tme.schedule.build import main

if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 6: Update the test module list**

In `tests/test_imports.py`, replace the trailing root-module entries:

```python
    "grid",
    "run_fill",
    "probe",
    "delete_boards",
    "diagnose_row",
```

with:

```python
    "tme.grid",
    "tme.cli.fill",
    "tme.cli.build",
    "tme.tools.probe",
    "tme.tools.delete_boards",
    "tme.tools.diagnose_row",
```

The list's closing `]` is already on the next line — do not add a second one. The list is 18 entries when you are done; check with `python -c "import tests.test_imports as t; print(len(t.MODULES))"`.

- [ ] **Step 7: Run the suite**

```bash
python -m pytest -v
```

Expected: 26 passed (18 imports, 4 golden, 4 snapshot).

- [ ] **Step 8: Check the reference image resolves from another directory**

The whole point of Step 3. Run it from somewhere that is not the repo root.

```bash
cd /c/Users/123 && python -c "
import sys; sys.path.insert(0, '/c/Users/123/Documents/Claude/Code/TME-diagram-maker')
from tme.win.window import save_prompt_reference
import os
p = save_prompt_reference()
print(p, os.path.exists(p))
"; cd /c/Users/123/Documents/Claude/Code/TME-diagram-maker
```

Expected: a path ending in `tme/tools/assets/save_prompt_reference.png` and `True`.

- [ ] **Step 9: Check the root is clean**

```bash
ls *.py 2>/dev/null; echo "--- no .py above means the root is clear ---"
```

Expected: no Python files at the repo root.

- [ ] **Step 10: Commit**

```bash
git add -A
git commit -m "Move grid, the CLI and the diagnostics into the package

cli/ is the pipeline; tools/ is what you reach for when a run has gone wrong.
Different callers, different risk, different rules -- diagnose_row refuses to
write unless the focused row matches.

One deliberate behavioural change, confined to one argument: the save-prompt
reference image is now loaded as a package resource instead of by a bare
relative filename.  It resolved against the working directory before, which
only worked because every invocation happened to start at the repo root; from
a building folder it would have raised FileNotFoundError in the middle of a
fill, when the prompt appears."
```

---

### Task 8: Split config, add conventions, RunPaths and project.yaml

**Files:**
- Modify: `tme/config.py` (remove three groups of constants)
- Create: `tme/conventions.py`, `tme/paths.py`
- Create: `projects/mrdiy/project.yaml`
- Modify: `tme/schedule/reader.py`, `tme/schedule/build.py`, `tme/cli/fill.py`, `tme/win/session.py`, `tme/tools/probe.py`
- Create: `tests/test_paths.py`

**Interfaces:**
- Produces: `tme.conventions.SKIP_TOKEN`, `SPARE_MARKER`, `AUTO_NAME_TOKEN`; `tme.paths.RunPaths` with classmethods `for_building(path)` and `for_workbook(path)` and attributes `root`, `workbook`, `drawing`, `runs`, `state`, `log`, `intruder`, `probe`.

- [ ] **Step 1: Write the failing test for RunPaths**

`tests/test_paths.py`:

```python
"""RunPaths puts a run's artefacts beside the run that produced them.

Before this, .state.json and the fill log landed in the current working
directory -- which meant the repo root, mixed in with the source, and shared
between every building anyone happened to run.
"""

from __future__ import annotations

import pytest

from tme.paths import RunPaths


def test_for_building_derives_everything_from_the_folder(tmp_path):
    building = tmp_path / "Y9"
    building.mkdir()
    (building / "schedule.xlsx").write_bytes(b"")
    (building / "drawing.yaml").write_text("boards: []", "utf-8")

    paths = RunPaths.for_building(building)

    assert paths.workbook == building / "schedule.xlsx"
    assert paths.drawing == building / "drawing.yaml"
    assert paths.state == building / "runs" / "state.json"
    assert paths.log == building / "runs" / "fill.log"
    assert paths.intruder == building / "runs" / "interrupted_by.png"
    assert paths.probe == building / "runs" / "probe_results.json"


def test_for_building_creates_the_runs_folder(tmp_path):
    building = tmp_path / "Y9"
    building.mkdir()
    (building / "schedule.xlsx").write_bytes(b"")

    RunPaths.for_building(building)

    assert (building / "runs").is_dir()


def test_for_building_refuses_a_folder_with_no_schedule(tmp_path):
    building = tmp_path / "Y9"
    building.mkdir()

    with pytest.raises(FileNotFoundError, match="schedule.xlsx"):
        RunPaths.for_building(building)


def test_for_workbook_puts_runs_beside_the_named_file(tmp_path):
    workbook = tmp_path / "example-schedule.xlsx"
    workbook.write_bytes(b"")

    paths = RunPaths.for_workbook(workbook)

    assert paths.workbook == workbook
    assert paths.log == tmp_path / "runs" / "fill.log"


def test_resolve_accepts_either_a_folder_or_a_workbook(tmp_path):
    building = tmp_path / "Y9"
    building.mkdir()
    (building / "schedule.xlsx").write_bytes(b"")

    assert RunPaths.resolve(building).workbook == building / "schedule.xlsx"
    assert RunPaths.resolve(building / "schedule.xlsx").workbook == (
        building / "schedule.xlsx"
    )
```

- [ ] **Step 2: Run it and watch it fail**

```bash
python -m pytest tests/test_paths.py -v
```

Expected: FAIL — `ModuleNotFoundError: No module named 'tme.paths'`.

- [ ] **Step 3: Write RunPaths**

`tme/paths.py`:

```python
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
```

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_paths.py -v
```

Expected: 5 passed.

- [ ] **Step 5: Move the schedule conventions out of config**

`tme/conventions.py`:

```python
"""Conventions of our own Excel schedule format.

Not measured TME behaviour and not client-specific -- these are the sentinels
we chose, and they change only if we change the schedule format.
"""

from __future__ import annotations

# "use whatever TME already put there" -- do not write, but still verify.
SKIP_TOKEN = "[SKIP]"

# Marks a circuit row: TME generates the Name itself from DB + circuit number.
AUTO_NAME_TOKEN = "[AUTO GENERATED BY ROW B]"

# A circuit whose Cable Specification is exactly this is a spare way.
SPARE_MARKER = "SPARE"
```

Delete those three constants and their comments from `tme/config.py`.

- [ ] **Step 6: Delete the file constants from config**

Remove this whole block from `tme/config.py`:

```python
# --------------------------------------------------------------------------
# Files
# --------------------------------------------------------------------------

PROBE_RESULTS_FILE = "probe_results.json"
STATE_FILE = ".state.json"
DEFAULT_WORKBOOK = "examples/example-schedule.xlsx"
INTRUDER_SHOT_FILE = "interrupted_by.png"
```

(Including the `DEFAULT_WORKBOOK` comment above it.) Also remove `SAVE_PROMPT_REFERENCE`, replaced by `window.save_prompt_reference()` in Task 7.

Keep `SPARE_SYSTEM_FALLBACK` in `config.py` for now — Step 8 moves it.

- [ ] **Step 7: Repoint every reference**

```bash
grep -rn "SKIP_TOKEN\|AUTO_NAME_TOKEN\|SPARE_MARKER\|STATE_FILE\|PROBE_RESULTS_FILE\|INTRUDER_SHOT_FILE\|DEFAULT_WORKBOOK\|SAVE_PROMPT_REFERENCE" tme/
```

For each hit:

- `config.SKIP_TOKEN` / `config.AUTO_NAME_TOKEN` / `config.SPARE_MARKER` → `conventions.X`, adding `from tme import conventions` to that module.
- `config.INTRUDER_SHOT_FILE` in `tme/win/session.py` (two hits, original lines 691 and 769) → a `intruder_shot` attribute passed into `TmeSession.__init__` with a default of `"interrupted_by.png"`, set by `fill.py` from `paths.intruder`. Add the parameter; do not read a global.
- `config.STATE_FILE` in `tme/cli/fill.py` (line 105) → `paths.state`.
- `config.PROBE_RESULTS_FILE` in `tme/tools/probe.py` → `paths.probe`.
- `config.DEFAULT_WORKBOOK` in `tme/cli/fill.py` (line 54) → see Step 9.

- [ ] **Step 8: Add project.yaml and load the client conventions**

`projects/mrdiy/project.yaml`:

```yaml
# Client-level conventions.  Agreed with the client or decided by us -- never
# anything read off a drawing.  Cable specs deliberately stay in each building's
# drawing.yaml: a shared spec is a spec nobody re-read against the sheet, and
# re-reading is what catches the ditto marks and the reversed label ordering.
client: "MR. D.I.Y. Trading (Thailand)"

# Spare ways take the same System as the real ways in the same board, so one
# board never mixes Power System and Lighting System (decided 2026-08-12).
spare_system_fallback: "Power System"
```

Move `SPARE_SYSTEM_FALLBACK` out of `tme/config.py` and read it here instead. In `tme/schedule/build.py`, add:

```python
def load_project(building_dir: Path) -> dict:
    """Read the client-level project.yaml above a building folder.

    Falls back to the built-in defaults when there is none, so examples/ and a
    fresh clone still build.
    """
    candidate = Path(building_dir).parent.parent / "project.yaml"
    if not candidate.exists():
        return {"spare_system_fallback": "Power System"}
    return yaml.safe_load(candidate.read_text("utf-8")) or {}
```

and pass `project["spare_system_fallback"]` into `board_rows` in place of the `config.SPARE_SYSTEM_FALLBACK` reference.

- [ ] **Step 9: Give the build CLI the new target argument**

Not cosmetic. `build.main` currently derives its output from the YAML's stem:

```python
    out = args.out or str(Path(args.drawing).with_name(
        Path(args.drawing).stem + "-schedule.xlsx"
    ))
```

Under the new layout the YAML is `drawing.yaml`, so that produces
`drawing-schedule.xlsx` — not the `schedule.xlsx` that `RunPaths` looks for.
Left alone, `build` and `fill` would disagree about the filename.

In `tme/schedule/build.py`, replace the argument and the `out` derivation:

```python
    parser.add_argument(
        "target",
        help="a building folder (projects/<client>/buildings/<slug>) whose "
             "drawing.yaml is read and schedule.xlsx written, or a path to a "
             "single YAML file",
    )
    parser.add_argument("-o", "--out", default=None,
                        help="output workbook (default: schedule.xlsx in the "
                             "building folder)")
    parser.add_argument("--allow-warnings", action="store_true",
                        help="write the workbook even if the read-back warns")
    args = parser.parse_args(argv)

    target = Path(args.target)
    if target.is_dir():
        drawing = target / "drawing.yaml"
        out = args.out or str(target / "schedule.xlsx")
    else:
        drawing = target
        out = args.out or str(drawing.with_name(drawing.stem + "-schedule.xlsx"))
```

and replace `build(args.drawing, out)` with `build(str(drawing), out)`.

The single-YAML branch is kept so `examples/` and any loose file still build.

- [ ] **Step 10: Give the fill CLI the new target argument**

In `tme/cli/fill.py`, replace:

```python
    parser.add_argument("workbook", nargs="?", default=config.DEFAULT_WORKBOOK)
```

with:

```python
    parser.add_argument(
        "target",
        help="a building folder (projects/<client>/buildings/<slug>) or a "
             "workbook path. Run artefacts -- state, log, screenshots -- go "
             "into runs/ beside it.",
    )
```

and after the parse, before `load_boards`:

```python
    try:
        paths = RunPaths.resolve(args.target)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
```

Then replace `load_boards(args.workbook)` with `load_boards(str(paths.workbook))`, `Path(config.STATE_FILE)` with `paths.state`, and `"workbook": args.workbook` in the state payload with `"workbook": str(paths.workbook)`.

Add the import: `from tme.paths import RunPaths`.

- [ ] **Step 11: Run the whole suite**

```bash
python -m pytest -v
```

Expected: 31 passed (18 imports, 4 golden, 4 snapshot, 5 paths).

Then check both CLIs agree on the workbook filename, which is what Step 9 was for:

```bash
python -m tme.cli.build projects/mrdiy/buildings/Y4 && python -m tme.cli.fill_all projects/mrdiy/buildings/Y4 --dry-run 2>/dev/null || ls projects/mrdiy/buildings/Y4
```

Expected: `build` rewrites `projects/mrdiy/buildings/Y4/schedule.xlsx` (not `drawing-schedule.xlsx`) and prints the same board and circuit counts as before. `fill_all` does not exist until Task 9, so the fallback `ls` is what you will see — check the folder holds `drawing.yaml`, `schedule.xlsx` and `runs/`, and nothing named `drawing-schedule.xlsx`.

- [ ] **Step 12: Check the golden test still passes after the config split**

The spare-system fallback moved from a constant to a YAML value. If its value changed by accident, the golden test fails on any building with a spare way — which is the point of having it. Confirm it did not:

```bash
python -m pytest tests/test_build_golden.py -v
```

Expected: 4 passed. A failure here means the fallback value or the conventions moved with the wrong string; compare against `git show HEAD~1:tme/config.py`.

- [ ] **Step 13: Commit**

```bash
git add -A
git commit -m "Split config three ways and derive run paths from the target

config.py held measured TME behaviour, our own schedule sentinels, and four
bare artefact filenames in one module.  The filenames were relative and
resolved against the working directory, so .state.json and the fill log landed
at the repo root and every building overwrote the last one's state.

RunPaths derives them from the building folder instead, and a run now names its
target -- DEFAULT_WORKBOOK is gone.  Each constant kept the comment that
records why it has the value it has."
```

---

### Task 9: The fill_all driver

**Files:**
- Create: `tme/cli/fill_all.py`
- Create: `tests/test_fill_all.py`
- Delete: `run_y3.sh`, `run_y4.sh`

**Interfaces:**
- Consumes: `tme.paths.RunPaths`, `tme.schedule.reader.load_boards`
- Produces: `tme.cli.fill_all.board_order(workbook)` → `list[str]`, and `main(argv)`.

- [ ] **Step 1: Write the failing test**

The subprocess loop cannot be tested without TME, but the part that has actually caused a wrong run — the board order — can be.

`tests/test_fill_all.py`:

```python
"""The board order comes from the workbook, not from a list kept by hand.

run_y4.sh repeated its 16 board names by hand, in an order that had to match
what order_boards() computed.  A YAML edit not mirrored into the .sh raised no
error -- it ran a parent board before its child, and TME had nothing to link
the feeder to, mid-run, with earlier rows already committed and unremovable.
"""

from __future__ import annotations

import pytest

from tests.clientdata import SLUGS, require

from tme.cli.fill_all import board_order


@pytest.mark.parametrize("slug", SLUGS)
def test_board_order_comes_from_the_workbook(slug: str) -> None:
    _, workbook = require(slug)
    order = board_order(workbook)
    assert order, "no boards read from the workbook"
    assert len(order) == len(set(order)), "a board name appears twice"


def test_from_board_drops_everything_before_it() -> None:
    from tme.cli.fill_all import drop_until

    boards = ["A", "B", "C", "D"]
    assert drop_until(boards, "C") == ["C", "D"]
    assert drop_until(boards, None) == boards


def test_from_board_that_is_not_there_raises() -> None:
    from tme.cli.fill_all import drop_until

    with pytest.raises(ValueError, match="no board named"):
        drop_until(["A", "B"], "Z")
```

- [ ] **Step 2: Run it and watch it fail**

```bash
python -m pytest tests/test_fill_all.py -v
```

Expected: FAIL — `ModuleNotFoundError: No module named 'tme.cli.fill_all'`.

- [ ] **Step 3: Write the driver**

`tme/cli/fill_all.py`:

```python
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
```

- [ ] **Step 4: Run the test**

```bash
python -m pytest tests/test_fill_all.py -v
```

Expected: 6 passed.

- [ ] **Step 5: Check --dry-run against the real workbook**

```bash
python -m tme.cli.fill_all projects/mrdiy/buildings/Y4 --dry-run
```

Expected: 16 boards listed. Compare against the `BOARDS=(...)` array in `run_y4.sh` — same names, same order. If they differ, **do not adjust the driver to match the shell script**: the workbook is the source of truth, and a difference means the shell script had drifted, which is the failure this task exists to remove. Record the difference in the commit message.

- [ ] **Step 6: Delete the shell scripts**

```bash
rm -f run_y3.sh run_y4.sh
ls *.sh 2>/dev/null; echo "--- no .sh above means both are gone ---"
```

- [ ] **Step 7: Run the whole suite**

```bash
python -m pytest -q
```

Expected: 37 passed.

- [ ] **Step 8: Commit**

```bash
git add -A
git commit -m "Read the board order from the workbook instead of a shell array

run_y3.sh and run_y4.sh each repeated a board list that order_boards() had
already computed and written into the workbook.  A YAML edit not mirrored into
the .sh did not raise anything -- it ran a parent before its child, and TME had
nothing to link the feeder to, mid-run, with earlier rows already committed.

One subprocess per board is kept: a stop never rewrites a board that already
verified.  Being Python rather than a shell pipeline also removes the
exit-status trap in CLAUDE.md, where a failed run came back as success."
```

---

### Task 10: Update README.md and CLAUDE.md

The documentation currently describes a layout that no longer exists, and warns about a shell trap that no longer applies.

**Files:**
- Modify: `README.md`
- Modify: `CLAUDE.md`

- [ ] **Step 1: Find every stale path and command**

```bash
grep -n "drawings/\|run_fill.py\|build_schedule.py\|excel_reader.py\|pdfcrop.py\|delete_boards.py\|diagnose_row.py\|probe.py\|tmeio.py\|\.state\.json\|run_y[0-9]\.sh\|tee" README.md CLAUDE.md
```

Every hit needs updating. Work through the list.

- [ ] **Step 2: Update the pipeline diagram in CLAUDE.md**

Replace:

```
drawing PDF -> pdfcrop.py -> read by eye -> drawings/<sheet>.yaml
            -> build_schedule.py -> <sheet>-schedule.xlsx
            -> run_fill.py -> TME
```

with:

```
drawing PDF -> tme.cli.pdfcrop -> read by eye
            -> projects/<client>/buildings/<slug>/drawing.yaml
            -> tme.cli.build   -> .../schedule.xlsx
            -> tme.cli.fill_all -> TME
```

- [ ] **Step 3: Update the separability paragraph in CLAUDE.md**

Replace:

```
Every stage is separable. `excel_reader.py`, `build_schedule.py` and `pdfcrop.py`
run without TME open; only `run_fill.py`, `delete_boards.py`, `diagnose_row.py`
and `probe.py` touch the application.
```

with:

```
Every stage is separable. `tme/schedule/` and `tme/cli/pdfcrop.py` run without
TME open; only `tme/cli/fill.py`, `tme/cli/fill_all.py` and the three modules in
`tme/tools/` touch the application. `tme/win/` is the whole Windows surface --
if anything here ever needs porting, that is the directory.
```

- [ ] **Step 4: Rewrite the shell-traps section in CLAUDE.md**

Replace the whole "Shell traps that have burned this project" section with:

```markdown
## Traps that have burned this project

**`delete_boards.py` only deletes the topmost board.** A partial board anywhere
else has to be removed by hand in TME: click its board row, then `Delete Row`.

**Resume rather than restart.** `python -m tme.cli.fill_all <building> --from BOARD`
picks up where a stop left off. It runs one process per board, so boards already
verified are never rewritten -- but the skipped boards must already exist in the
project, or a feeder pointing at one has nothing to link to.

**Do not pipe a fill through `tee`, and do not follow it with `; echo`.** A shell
pipeline reports the *last* command's exit status, so a failed run comes back as
success. This is why `fill_all` is Python and writes to `runs/fill.log` itself;
the trap is still live for anything you run by hand.
```

- [ ] **Step 5: Add a layout section to CLAUDE.md**

Insert after the "What this is" section:

```markdown
## Where things are

`tme/` is the engine and is client-agnostic. `projects/<client>/buildings/<slug>/`
holds one building: the `drawing.yaml` read off the sheet, the `schedule.xlsx`
built from it, and a `runs/` folder for everything a fill against it produced.

`projects/` is gitignored in one line, so a new client inherits the protection.
Nothing under it may reach a public remote -- see the last section of this file.

A new client is a new folder: `projects/<slug>/project.yaml` for the conventions
agreed with them, then one folder per building. Cable specs stay in each
building's `drawing.yaml` on purpose. A shared spec is a spec nobody re-read
against the sheet in front of them, and re-reading is the step that catches the
traps in "Reading a drawing" below.
```

- [ ] **Step 6: Update README.md commands**

Every command in README.md changes form. The mapping:

| Old | New |
|---|---|
| `python pdfcrop.py <pdf> ...` | `python -m tme.cli.pdfcrop <pdf> ...` |
| `python build_schedule.py drawings/X.yaml out.xlsx` | `python -m tme.cli.build projects/<client>/buildings/<slug>` |
| `python excel_reader.py <xlsx>` | `python -m tme.schedule.reader <xlsx>` |
| `python run_fill.py <xlsx> [flags]` | `python -m tme.cli.fill <building> [flags]` |
| `python probe.py ...` | `python -m tme.tools.probe ...` |
| `python delete_boards.py ...` | `python -m tme.tools.delete_boards ...` |
| `python diagnose_row.py ...` | `python -m tme.tools.diagnose_row ...` |
| (nothing) | `python -m tme.cli.fill_all <building> [flags]` |

Add a short section for `fill_all` covering `--from`, `--dry-run`, and that it writes `runs/fill.log`.

- [ ] **Step 7: Add a testing section to README.md**

```markdown
## Tests

```bash
python -m pytest
```

Four things are covered, and it is worth being clear about what is not:

- every module imports
- `RunPaths` derives the right artefact locations
- rebuilding each building's schedule from its YAML reproduces the approved
  workbook cell for cell
- `load_boards` returns the same boards, circuit counts and warnings as before

The golden and snapshot tests need the client data in `projects/`, and skip with
a message when it is absent -- a clone without it is a valid clone.

Nothing from `tme/grid.py` inward is covered. No test can drive TME. The first
run after any change to `tme/win/` or `tme/grid.py` should be a single board,
against a project file you have backed up, with someone watching -- not a
sixteen-board sweep.
```

- [ ] **Step 8: Check nothing stale is left**

```bash
grep -n "drawings/\|run_fill.py\|build_schedule.py\|excel_reader.py\|run_y[0-9]\.sh\|\.state\.json" README.md CLAUDE.md
```

Expected: no output, except where a line is deliberately quoting the old layout as history.

- [ ] **Step 9: Run the suite one last time**

```bash
python -m pytest -q && git status --short
```

Expected: 37 passed, and `git status` shows only the two documentation files modified.

- [ ] **Step 10: Commit**

```bash
git add README.md CLAUDE.md
git commit -m "Document the package layout and retire the shell-trap section

The pipeline diagram, every command, and the separability paragraph described
a flat repo that no longer exists.  The tee/pipeline warning stays, narrowed:
fill_all is Python and writes its own log, but the trap is still live for
anything run by hand.

Adds a section saying what the tests cover and, more usefully, what they do
not -- nothing from grid.py inward, so the first run after touching tme/win/
is one board against a backed-up project file."
```

---

## Verification checklist

After Task 10, before the first real run:

- [ ] `python -m pytest` → 37 passed
- [ ] `ls *.py *.sh` at the repo root → nothing
- [ ] `git status --short` → clean; nothing under `projects/` listed
- [ ] `python -m tme.cli.fill_all projects/mrdiy/buildings/Y4 --dry-run` → 16 boards, children first
- [ ] `git log --oneline -10` → one commit per task, each with its reasoning

**The first run against TME must be one board on a project file backed up first.** Layers 1–3 of the test net stop at `Board` objects; everything from `tme/grid.py` inward rests on Task 6's diff and on someone watching. TME has no undo, and a wrong row is a row that has to be deleted by hand.
