"""The board order comes from the workbook, not from a list kept by hand.

run_y4.sh repeated its 16 board names by hand, in an order that had to match
what order_boards() computed.  A YAML edit not mirrored into the .sh raised no
error -- it ran a parent board before its child, and TME had nothing to link
the feeder to, mid-run, with earlier rows already committed and unremovable.
"""

from __future__ import annotations

import pytest

from tests.clientdata import SLUGS, require

from tme.cli.fill_all import board_order, drop_until


@pytest.mark.parametrize("slug", SLUGS)
def test_board_order_comes_from_the_workbook(slug: str) -> None:
    _, workbook = require(slug)
    order = board_order(workbook)
    assert order, "no boards read from the workbook"
    assert len(order) == len(set(order)), "a board name appears twice"


@pytest.mark.parametrize("slug", SLUGS)
def test_the_board_nothing_feeds_comes_last(slug: str) -> None:
    """Children first is the whole point: TME cannot link to a board that is
    not there yet.  The MSB feeds everything and is fed by nothing, so it is
    the one board that must be created last."""
    _, workbook = require(slug)
    assert board_order(workbook)[-1].startswith("MSB-")


def test_from_board_drops_everything_before_it() -> None:
    boards = ["A", "B", "C", "D"]
    assert drop_until(boards, "C") == ["C", "D"]
    assert drop_until(boards, None) == boards


def test_from_board_that_is_not_there_raises() -> None:
    with pytest.raises(ValueError, match="no board named"):
        drop_until(["A", "B"], "Z")
