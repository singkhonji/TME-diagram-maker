"""Every module imports cleanly.

Trivial, and it is the test that will actually fire during the layout refactor:
a moved module whose imports were not updated fails here in a second, instead of
failing in front of TME with rows already committed.
"""

from __future__ import annotations

import importlib

import pytest

MODULES = [
    "tme.config",
    "tme.conventions",
    "tme.project",
    "tme.paths",
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
    "tme.grid",
    "tme.cli.fill",
    "tme.cli.build",
    "tme.cli.fill_all",
    "tme.tools.probe",
    "tme.tools.delete_boards",
    "tme.tools.diagnose_row",
]


@pytest.mark.parametrize("name", MODULES)
def test_module_imports(name: str) -> None:
    importlib.import_module(name)
