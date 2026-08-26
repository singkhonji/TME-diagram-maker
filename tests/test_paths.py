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


def test_project_yaml_overrides_the_spare_system_rule(tmp_path):
    """The client rule reaches the reader, not just the config file.

    Worth a test of its own: no committed workbook exercises it.  build.py
    writes a System into every row, so the fallback branch never fires when
    reading one back, and a broken wiring would look fine.
    """
    from tme import project

    client = tmp_path / "acme"
    building = client / "buildings" / "B1"
    building.mkdir(parents=True)
    (building / "schedule.xlsx").write_bytes(b"")
    (client / "project.yaml").write_text(
        'client: "Acme"\nspare_system_fallback: "Fire Alarm System"\n', "utf-8"
    )

    settings = RunPaths.for_building(building).project

    assert settings.client == "Acme"
    assert settings.spare_system_fallback == "Fire Alarm System"


def test_a_building_with_no_project_yaml_falls_back_to_the_convention(tmp_path):
    from tme import conventions

    building = tmp_path / "loose"
    building.mkdir()
    (building / "schedule.xlsx").write_bytes(b"")

    settings = RunPaths.for_building(building).project

    assert settings.source is None
    assert settings.spare_system_fallback == conventions.SPARE_SYSTEM_FALLBACK
