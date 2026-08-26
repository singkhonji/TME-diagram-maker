# CLAUDE.md

Notes for Claude Code working in this repo. README.md is the source of truth for
how the tool works and what TME does; this file is about how to work on it
safely.

## What this is

A pipeline that turns an electrical Single Line Diagram into filled-in rows of
Cubicost TME's Schematic Table:

```
drawing PDF -> tme.cli.pdfcrop -> read by eye
            -> projects/<client>/buildings/<slug>/drawing.yaml
            -> tme.cli.build    -> .../schedule.xlsx
            -> tme.cli.fill_all -> TME
```

Every stage is separable. `tme/schedule/` and `tme/cli/pdfcrop.py` run without
TME open; only `tme/cli/fill.py`, `tme/cli/fill_all.py` and the three modules in
`tme/tools/` touch the application. `tme/win/` is the whole Windows surface --
if anything here ever needs porting, that is the directory.

## Where things are

`tme/` is the engine and is client-agnostic. `projects/<client>/buildings/<slug>/`
holds one building: the `drawing.yaml` read off the sheet, the `schedule.xlsx`
built from it, and a `runs/` folder for everything a fill against it produced.

`projects/` is gitignored in one line, so a new client inherits the protection.
Nothing under it reaches the remote at all -- see the last section of this file.

A new client is a new folder: `projects/<slug>/project.yaml` for the conventions
agreed with them, then one folder per building. Cable specs stay in each
building's `drawing.yaml` on purpose. A shared spec is a spec nobody re-read
against the sheet in front of them, and re-reading is the step that catches the
traps in "Reading a drawing" below.

## The three rules that matter most

**TME has no undo.** Rows are committed to the project the moment they are
created. Neither `Esc` nor the Cancel button removes them. Before any run that
writes, the project file must be backed up, and a board that was only partly
filled has to be deleted before that board is retried — TME stops to ask about a
repeated name otherwise.

**Never press OK.** The script fills the table; a human reviews it and commits
it. Nothing in this repo may click that button.

**Never guess a delay.** Every wait is a poll against evidence that the action
landed — usually the clipboard. `time.sleep()` as a way of hoping TME kept up is
the bug this whole design exists to avoid.

## Working on the automation

The Schematic Table is a custom-drawn canvas: UI Automation reports
`ControlType: Custom(50025)`, no Name, not focusable. Synthetic keys plus the
clipboard are the only channel. Everything known about how it responds was
measured, not assumed, and lives in README.md's behaviour tables and in the
module docstrings. **Read those before changing key sequences** — several of the
rules there look like superstition and are not (`Ctrl+A` outside an editor
selects every row; a bare `Esc` cancels the whole dialog and discards the run).

When something goes wrong, do not tune constants against the symptom. Two
sessions were lost that way. What worked:

1. Run with `--verbose` — every key sent and every value read back is logged.
2. Find a row that **worked** and put its trace next to the row that failed.
   Identical key sequences with different outcomes narrow it to one variable
   immediately.
3. `tme/tools/diagnose_row.py` replays one row's write/verify with tracing. It refuses to
   write unless the focused row's Circuit Number matches, so point it at a row
   that is already broken and due for deletion.

Do not read cells in the `NEVER_F2_KEYS` columns. F2 on Conduit Size raises
"Please enter an integer between [10,5000]" and on Elevation "Elevation input
error", and the dialog then blocks everything after it.

## Traps that have burned this project

**`tme/tools/delete_boards.py` only deletes the topmost board.** A partial board
anywhere else has to be removed by hand in TME: click its board row, then
`Delete Row`.

**Resume rather than restart.** `python -m tme.cli.fill_all <building> --from BOARD`
picks up where a stop left off. It runs one process per board, so boards already
verified are never rewritten -- but the skipped boards must already exist in the
project, or a feeder pointing at one has nothing to link to.

**Do not pipe a fill through `tee`, and do not follow it with `; echo`.** A shell
pipeline reports the *last* command's exit status, so a failed run comes back as
success. This is why `fill_all` is Python and writes to `runs/fill.log` itself;
the trap is still live for anything you run by hand.

## Reading a drawing

These PDFs are CAD exports with the text converted to outlines — `pdfplumber`
finds zero characters per page and `pdftotext` returns nothing. There is no
extraction shortcut; the labels have to be looked at. `tme.cli.pdfcrop --scale 2.6`
is legible on an A1 sheet, and `--rotate cw` turns the vertical load
descriptions upright.

Traps worth re-checking on every new sheet, all of which have already caused a
wrong reading once:

- breaker ratings use ditto marks, and the rating label sits to the **left** of
  the breaker it belongs to
- cable-spec labels can be ordered **opposite** to the feeders they serve —
  trace the height of each horizontal run, then sanity-check against the breaker
  rating
- the same board name can appear twice on one sheet for two different boards
- circuit references restart in every section (`L`, `P`, `AC`), so they need the
  section as a prefix before TME sees them

Anything read but uncertain goes in the YAML as a `check:` note rather than a
silent guess.

## Conventions

Comments explain **why**, especially where the code looks over-careful — most of
it is guarding against a measured TME behaviour, and a future reader who
simplifies it will walk back into the same trap. Where a rule can be enforced in
code instead of documented (`NEVER_F2_KEYS`, the topological board ordering),
enforce it in code.

Client drawings and filled schedules are real project data. They never go in a
commit. `projects/` is gitignored wholesale for exactly this, and the rule holds
regardless of whether the remote is private today -- visibility is one settings
click away from changing, and history cannot be un-pushed.

A consequence worth knowing when someone new joins: cloning this repo gets them
the engine and nothing else. The drawings, the built schedules and the run logs
have to be handed over separately.
