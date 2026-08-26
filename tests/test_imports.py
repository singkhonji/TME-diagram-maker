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
