"""Build a schedule workbook from a drawing YAML.

    python -m tme.cli.build projects/<client>/buildings/<slug>

A thin wrapper so every entry point lives under tme.cli, including the one that
never touches TME.
"""

from __future__ import annotations

from tme.schedule.build import main

if __name__ == "__main__":
    raise SystemExit(main())
