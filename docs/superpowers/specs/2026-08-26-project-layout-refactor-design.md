# Project layout refactor — design

Date: 2026-08-26
Status: approved, not yet implemented

## Goal

Separate the reusable TME engine from per-client project data, so a new client
job is a new folder rather than a new set of files scattered through the repo
root. Along the way, remove three pieces of duplication that have each already
had a chance to cause a wrong run.

Nothing about how the filler talks to TME changes. Every measured behaviour
recorded in README.md and in the module docstrings survives this refactor
unaltered — that is a constraint, not an aspiration, and Section 6 says how it
is enforced.

## Current state

Ten Python modules sit flat at the repo root and import each other by bare name
(`import config`, `import tmeio`). Client data lives in `drawings/`, which holds
four buildings of one job (MR. D.I.Y.: `BLD_Y1` .. `BLD_Y4`), each building a
YAML read off a drawing plus the schedule built from it. Run artefacts —
`.state.json`, `y2_runs.log`, `y3-fill.log`, `y4-fill.log`, `interrupted_by.png`
— land in the repo root, mixed in with the source.

Two hand-written shell scripts, `run_y3.sh` and `run_y4.sh`, are near-identical
and each hardcodes a board list.

Three specific problems this design addresses:

1. **`drawings/` cannot hold a second client.** The folder is flat and the
   filenames carry the building, not the job. A second client's `BLD_A1` would
   sit beside MR. D.I.Y.'s `BLD_Y1` with nothing distinguishing them.
2. **The board list is stored twice.** `build_schedule.order_boards()` computes
   the children-first order and writes it into the workbook; `run_y4.sh` then
   repeats those 16 board names by hand. A YAML edit that is not mirrored into
   the `.sh` does not raise an error — it runs a parent board before its child,
   and TME has nothing to link the feeder to. That failure lands mid-run, after
   earlier rows are already committed and unremovable.
3. **`config.py` mixes three kinds of value.** Measured TME behaviour, our own
   schedule conventions, and run-artefact filenames all sit in one 11 KB module.
   The filenames are bare and relative, so they resolve against the current
   working directory; the code only works because every invocation happens to
   `cd` to the repo root first.

## Section 1 — Folder layout

```
TME-diagram-maker/
├─ tme/                        the engine, one package
├─ projects/                   all client data (gitignored)
│   └─ mrdiy/
│       ├─ project.yaml        client-level conventions
│       └─ buildings/
│           ├─ Y1/  drawing.yaml  schedule.xlsx  runs/
│           ├─ Y2/  drawing.yaml  schedule.xlsx  runs/
│           ├─ Y3/  drawing.yaml  schedule.xlsx  runs/
│           └─ Y4/  drawing.yaml  schedule.xlsx  runs/
├─ examples/                   synthetic schedule, committed
├─ templates/
├─ tests/
├─ docs/
├─ README.md  CLAUDE.md  requirements.txt  pyproject.toml
```

Each building folder owns its own `runs/`, holding what currently litters the
root: `state.json`, `fill.log`, `interrupted_by.png`, `delete_popup_*.png`,
`probe_results.json`. A run's artefacts belong beside the run, not beside the
code.

`.gitignore` replaces the `drawings/` line with `projects/`. One line covers the
whole client tree, and a new client inherits the protection without anyone
remembering to add it. The root-level `*.log` and `*.png` rules can then narrow,
because artefacts no longer land at the root.

Building folder names drop the `BLD_` prefix (`Y1`, not `BLD_Y1`): the client
folder above them already says whose buildings these are.

## Section 2 — Engine package

```
tme/
├─ __init__.py
├─ config.py             measured TME behaviour only
├─ conventions.py        schedule conventions
├─ paths.py              RunPaths
├─ win/
│   ├─ __init__.py
│   ├─ input.py          ← tmeio.py: SendInput structs, key/char events, click
│   ├─ clipboard.py      ← tmeio.py: clip_get/clip_set/clip_settle, sequence
│   ├─ window.py         ← tmeio.py: find/focus/capture/compare/looks_like
│   └─ session.py        ← tmeio.py: TmeSession, CellRead, Throttle, AbortWatch,
│                          and the TmeIoError hierarchy
├─ grid.py               ← grid.py
├─ schedule/
│   ├─ __init__.py
│   ├─ reader.py         ← excel_reader.py
│   └─ build.py          ← build_schedule.py
├─ cli/
│   ├─ __init__.py
│   ├─ fill.py           ← run_fill.py
│   ├─ fill_all.py       NEW — replaces run_y3.sh and run_y4.sh
│   ├─ build.py          thin wrapper over schedule/build.py
│   └─ pdfcrop.py        ← pdfcrop.py
└─ tools/
    ├─ __init__.py
    ├─ probe.py          ← probe.py
    ├─ diagnose_row.py   ← diagnose_row.py
    ├─ delete_boards.py  ← delete_boards.py
    └─ assets/
        └─ save_prompt_reference.png
```

### Why `tmeio.py` splits four ways

It is 38 KB and roughly 900 lines doing four jobs: raw win32 `SendInput`
plumbing, clipboard read/write with the settle protocol, window discovery and
image comparison, and `TmeSession` — the state machine that actually knows how
TME behaves.

Only `session.py` changes when TME changes. The `SendInput` struct definitions
have not been touched since they were written and never should be. Keeping them
in the same file as the logic that gets edited means every session-logic edit is
an edit in a file containing the input layer.

The `TmeIoError` hierarchy goes with `session.py` because that is what raises
almost all of it, and every consumer that catches those already imports the
session.

The subpackage is `win/`, not `io/`. `tme/io/` would shadow the stdlib `io`
module for anything doing a relative import inside it — legal under Python 3
absolute imports, but a trap nobody should have to remember. `win/` also says
what the layer is: the Windows-specific half, the part that would have to be
rewritten wholesale if TME ever ran anywhere else.

### Why `tools/` is separate from `cli/`

`cli/` is the pipeline you run to get work done. `tools/` is what you reach for
when a run has gone wrong — `probe.py` measures TME behaviour, `diagnose_row.py`
replays one row with tracing, `delete_boards.py` cleans up a partial board. They
are diagnostics with different callers, different risk, and different rules
(`diagnose_row.py` refuses to write unless the focused row matches).

### Package assets

`save_prompt_reference.png` becomes a package asset loaded through
`importlib.resources`, not a bare relative filename. Today
`config.SAVE_PROMPT_REFERENCE = "save_prompt_reference.png"` is resolved against
the current working directory at open time, which is fine only because every
invocation starts at the repo root. Once runs happen from a building folder,
that path fails — and it fails mid-run, when the save prompt appears, not at
startup.

`interrupted_by.png` at the root is a run artefact and moves to `runs/`.
`save_prompt_reference.png` is a code asset and moves into the package.

## Section 3 — Config split

`config.py` splits by what makes a value change.

| Destination | Contents | Changes when |
|---|---|---|
| `tme/config.py` | `COLUMNS`, `COLUMN_KEYS`, `COLUMN_LABELS`, `COLUMN_INDEX`, `N_COLUMNS`, `WRITE_KEYS`, `VERIFY_KEYS`, `NEVER_F2_KEYS`, `CIRCUIT_DEFAULTS`, `IGNORE_ON_VERIFY`, `DB_DEFAULT_NAME_PREFIX`, `WINDOW_TITLE_MATCH`, `FOCUS_REACQUIRE_TIMEOUT_S`, all `KEY_DELAY_*` / `READ_TIMEOUT_*` / `CLIPBOARD_*` / `EXPECT_EMPTY_TIMEOUT_S`, `ABORT_VK`, `ABORT_KEY_NAME`, all `SAVE_PROMPT_*` | TME is upgraded, or a value is re-measured |
| `tme/conventions.py` | `SKIP_TOKEN`, `AUTO_NAME_TOKEN`, `SPARE_MARKER` | our own schedule format changes |
| `projects/<client>/project.yaml` | `spare_system_fallback`, circuit-number prefix rule | a new client, or a new agreement with this one |

Every comment in `config.py` moves with the constant it explains. Those comments
record measurements and the reasoning behind them — the throttle floor block
alone is 30 lines of evidence — and a constant separated from its evidence is a
constant someone will "simplify" later.

`SAVE_PROMPT_REFERENCE` stops being a filename constant and becomes a resource
lookup in `win/window.py`.

### `project.yaml`

```yaml
client: "MR. D.I.Y. Trading (Thailand)"
spare_system_fallback: "Power System"
circuit_prefix: section   # section letter prefixes the circuit ref (L1, P3, AC2)
```

Loaded once by `schedule/build.py`, passed down as a value. No module-level
import of client data.

### `cables:` stays per-building — decision and rationale

The obvious candidate for sharing was the `cables:` anchor block, which each
building YAML repeats. Comparing them shows real overlap but real divergence:
14 anchors in Y1, 13 in Y2 and Y3, 15 in Y4. Y1 against Y4 differs by three
entries — Y4 adds `fp_240` and `dl_10`, Y1 has `gate`, which Y4 does not.

They stay per-building anyway, and the reason is not the divergence.
`BLD_Y4.yaml` records its own rule: everything was re-checked against that sheet
rather than carried over.

A shared cable spec is a spec nobody re-read against the sheet in front of them.
That runs directly against the reading discipline in CLAUDE.md — ditto marks on
breaker ratings, rating labels sitting to the left of their breaker, cable-spec
labels ordered opposite to the feeders they serve. Every one of those has already
produced a wrong reading once. Deduplicating the specs would remove the step that
catches them.

So `project.yaml` carries only what is agreed with the client, never anything
read off a drawing.

## Section 4 — `RunPaths`

Run-artefact filenames leave `config.py` and become a small value object in
`tme/paths.py`, built from a building folder:

```python
paths = RunPaths.for_building("projects/mrdiy/buildings/Y4")
paths.workbook   # projects/mrdiy/buildings/Y4/schedule.xlsx
paths.drawing    # projects/mrdiy/buildings/Y4/drawing.yaml
paths.state      # projects/mrdiy/buildings/Y4/runs/state.json
paths.log        # projects/mrdiy/buildings/Y4/runs/fill.log
paths.intruder   # projects/mrdiy/buildings/Y4/runs/interrupted_by.png
paths.probe      # projects/mrdiy/buildings/Y4/runs/probe_results.json
```

`RunPaths.for_building` creates `runs/` if absent and raises if `schedule.xlsx`
is missing, so a mistyped path fails before any key reaches TME.

There is also `RunPaths.for_workbook(path_to_xlsx)`, which puts `runs/` beside
the named workbook. This keeps the existing invocation style working and covers
`examples/example-schedule.xlsx`.

`DEFAULT_WORKBOOK` is deleted. A run names its target; there is no default.

### CLI surface

Every entry point accepts a building folder *or* a workbook path and picks the
matching `RunPaths` constructor:

```
python -m tme.cli.build    projects/mrdiy/buildings/Y4
python -m tme.cli.fill     projects/mrdiy/buildings/Y4 --start-at MSB-Y4 --limit-boards 1
python -m tme.cli.fill_all projects/mrdiy/buildings/Y4 --auto-dismiss-save
python -m tme.cli.pdfcrop  <pdf> --scale 2.6 --rotate cw
```

All existing flags on `run_fill.py` keep their names and meanings.

## Section 5 — `fill_all`

`tme/cli/fill_all.py` replaces both shell scripts.

Behaviour:

- Read the board order from the workbook via `schedule.reader.load_boards`. The
  workbook already stores boards children-first, because `build.order_boards()`
  wrote them that way. No board list is kept anywhere by hand.
- For each board, run one subprocess:
  `[sys.executable, "-m", "tme.cli.fill", <target>, "--start-at", <board>, "--limit-boards", "1", ...]`,
  forwarding any pass-through flags. One process per board preserves the
  isolation the shell loop had: a stop never rewrites a board that already
  verified.
- Append each board's stdout and stderr to `paths.log`, with the same
  `=== BOARD x ===` / `=== OK x ===` / `=== STOPPED at x (exit N) ===` markers
  the shell version wrote, so existing log-reading habits still work.
- Stop at the first non-zero exit and propagate that exit code.
- `--from BOARD` starts partway down the list, for resuming.
- `--dry-run` prints the board order and exits without touching TME.

Because this is Python rather than a shell pipeline, the exit-status trap
documented in CLAUDE.md — a pipeline reporting the last command's status, so a
failed run comes back as success — cannot occur. That CLAUDE.md warning gets
rewritten to say the runner is now Python and why.

`run_y3.sh` and `run_y4.sh` are deleted.

## Section 6 — Test net

The repo has no tests. Moving a 38 KB win32 module with no safety net is the
main risk in this design. Four layers, built **before** anything moves.

1. **Golden test on the pure path.** Run `schedule.build` against each of the
   four building YAMLs and assert every cell value matches the `schedule.xlsx`
   currently committed for that building. This covers YAML → ordering → rows →
   workbook end to end, which is the half of the pipeline that can be tested at
   all. Fixtures read the real client YAMLs from `projects/`; the test skips with
   a clear message when `projects/` is absent, so a fresh clone is not broken by
   data it cannot have.
2. **Snapshot of `load_boards`.** For each of the four workbooks, assert the
   board count, circuit count, board names in order, and the warning list are
   unchanged.
3. **Import smoke test.** Import every module in `tme/`. Catches import breakage
   from the move, which is the most likely failure and the cheapest to detect.
4. **Cut-and-paste discipline for the TME-facing code.** No automated test can
   drive TME. Instead, the `tmeio.py` split is constrained to be pure movement:
   zero logic edits, only the import header of each new file differing. It is
   verified mechanically — concatenate the four new files in the original order,
   strip the added import headers, and diff against the original `tmeio.py`. A
   non-empty diff means something was changed that was not meant to be.

   The same rule applies to every other module move in step 3: imports change,
   nothing else does. Any behavioural change is a separate, later commit.

   Layer 4 is a one-off check performed during the split commit, not a test that
   lives on: once `tmeio.py` is gone there is nothing left to diff against. Its
   output — the empty diff — is quoted in that commit message as the evidence.

Layer 4 is the one that matters most and costs least. It converts "we were
careful" into something checkable.

### What the test net does not cover

Layers 1–3 exercise the YAML → workbook → `Board` objects path only. Everything
from `grid.py` inward is unverified by anything except layer 4's discipline and
a human watching the first real run. The first run after this refactor should be
a single board on a throwaway TME project, with the project file backed up
first, and not a 16-board sweep.

## Section 7 — Migration order

One step per commit, tests green at each step.

1. Write test layers 1–3 against the **current** flat code. Establishes the
   baseline; nothing has moved yet.
2. Move client data into `projects/mrdiy/buildings/Y1..Y4/`, renaming
   `BLD_Y4.yaml` → `Y4/drawing.yaml` and `BLD_Y4-schedule.xlsx` →
   `Y4/schedule.xlsx`. Change `.gitignore` from `drawings/` to `projects/`.
   Point the test fixtures at the new paths.
3. Create the package skeleton and `git mv` each module into place. Fix imports.
   Run the tests.
4. Split `tmeio.py` into `win/`. Verify with the concatenation diff from layer 4.
5. Split `config.py`; add `conventions.py`, `paths.py`, and
   `projects/mrdiy/project.yaml`.
6. Add `fill_all.py`; delete `run_y3.sh` and `run_y4.sh`.
7. Update README.md and CLAUDE.md to the new layout, commands, and shell-trap
   wording.

Use `git mv` throughout so file history survives.

**Handling step 2 safely.** `drawings/` is gitignored, so git cannot recover a
mistake there. That step copies rather than moves; the originals are deleted only
after the tests confirm the new paths load identically. The four `.xlsx` files
are the output of hours of reading drawings by eye and are not reproducible from
anything in the repo.

Existing run artefacts at the root (`y2_runs.log`, `y3-fill.log`, `y4-fill.log`,
`.state.json`, `interrupted_by.png`) are historical. They move into the matching
building's `runs/`, except `y2_runs.log` and `.state.json`, which are stale and
get deleted.

## Out of scope

- Any change to key sequences, timings, or verification logic. Behaviour is
  frozen for this refactor.
- Automated testing of TME-facing code.
- Deduplicating cable specs across buildings — deliberately rejected, Section 3.
- Reading the new client's drawings. That is the next piece of work and starts
  from a clean `projects/<newclient>/` created by this design.

## Open question for implementation

The new client's folder name. `projects/mrdiy/` is short and filesystem-safe;
the full client string lives in `project.yaml` under `client:`. The new job needs
the same treatment — a short slug for the folder, the real name in the YAML.
